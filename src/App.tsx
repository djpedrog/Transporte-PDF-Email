import React, { useState, useCallback, useMemo, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { 
  FileBox, 
  Upload, 
  Play, 
  CheckCircle2, 
  AlertCircle, 
  Download, 
  Mail, 
  FileText, 
  Trash2, 
  Copy,
  ChevronRight,
  Loader2,
  ExternalLink,
  ShieldAlert
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import JSZip from 'jszip';
import { saveAs } from 'file-saver';
import { 
  TransportRecord, 
  FirmRecord, 
  ProcessingError, 
  normalizeKey, 
  generateTransportPdf, 
  cleanFilename, 
  generateEml,
  checkLibraries
} from './services';

type Tab = 'upload' | 'preview' | 'processing' | 'summary';

export default function App() {
  const [activeTab, setActiveTab] = useState<Tab>('upload');
  const [exportFile, setExportFile] = useState<File | null>(null);
  const [firmsFile, setFirmsFile] = useState<File | null>(null);
  
  const [transportData, setTransportData] = useState<TransportRecord[]>([]);
  const [firmsData, setFirmsData] = useState<FirmRecord[]>([]);
  const [isParsing, setIsParsing] = useState(false);
  
  const [diagnostics, setDiagnostics] = useState<{
    headerRow: number;
    columns: string[];
    exportRows: number;
    firmsRows: number;
    libHealth: { jspdf: boolean; autotable: boolean };
    amountSamples: { raw: any; parsed: number }[];
  }>({ 
    headerRow: -1, 
    columns: [], 
    exportRows: 0, 
    firmsRows: 0,
    libHealth: { jspdf: false, autotable: false },
    amountSamples: []
  });
  
  const [logs, setLogs] = useState<string[]>([]);
  const [errors, setErrors] = useState<ProcessingError[]>([]);
  const [progress, setProgress] = useState(0);
  const [isProcessing, setIsProcessing] = useState(false);
  const [results, setResults] = useState<{
    pdfs: { cod: string; name: string; blob: Blob }[];
    emls: { cod: string; name: string; blob: Blob }[];
    unmapped: string[];
  }>({ pdfs: [], emls: [], unmapped: [] });

  const [sentEmails, setSentEmails] = useState<Record<string, boolean>>({});

  const addLog = (msg: string) => setLogs(prev => [...prev, `${new Date().toLocaleTimeString()} - ${msg}`]);

  useEffect(() => {
    const health = checkLibraries();
    setDiagnostics(prev => ({ ...prev, libHealth: health }));
    addLog(`Aplicação iniciada. jsPDF: ${health.jspdf}, AutoTable: ${health.autotable}`);

    // Load sent states from localStorage
    const saved: Record<string, boolean> = {};
    for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (key && key.startsWith('emailSent_')) {
          const cod = key.replace('emailSent_', '');
          saved[cod] = true;
        }
    }
    setSentEmails(saved);
  }, []);

  const toggleEmailSent = (cod: string) => {
    const isSent = sentEmails[cod];
    if (!isSent) {
      localStorage.setItem(`emailSent_${cod}`, '1');
      setSentEmails(prev => ({ ...prev, [cod]: true }));
    }
  };

  const resetSentEmails = () => {
    const keysToRemove: string[] = [];
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      if (key && key.startsWith('emailSent_')) {
        keysToRemove.push(key);
      }
    }
    keysToRemove.forEach(k => localStorage.removeItem(k));
    setSentEmails({});
    addLog('Estados de envio de email limpos.');
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>, type: 'export' | 'firms') => {
    const file = e.target.files?.[0];
    if (file) {
      if (type === 'export') setExportFile(file);
      else setFirmsFile(file);
    }
  };

  const parseExcels = async () => {
    if (!exportFile || !firmsFile) return;
    setIsParsing(true);
    addLog('A ler ficheiros Excel...');

    try {
      // 1. Parse EXPORT_TRANSPORTES
      const exportBuf = await exportFile.arrayBuffer();
      const exportWb = XLSX.read(exportBuf, { cellDates: true });
      const exportSheet = exportWb.Sheets['Data'] || exportWb.Sheets[exportWb.SheetNames[0]];
      
      // Read as 2D array to find header
      const exportRaw = XLSX.utils.sheet_to_json(exportSheet, { header: 1, raw: false }) as string[][];
      
      const requiredHeaders = ["Cliente", "Data do documento", "Referência", "Montante em moeda interna", "Texto", "Chave referência 3"];
      let headerIdx = -1;
      let foundHeaders: string[] = [];

      for (let i = 0; i < exportRaw.length; i++) {
        const row = exportRaw[i].map(c => String(c || '').trim());
        // Detailed log for debugging if needed (invisible to user unless in logs)
        const matches = requiredHeaders.filter(h => row.some(cell => cell.includes(h)));
        if (matches.length >= 5) {
          headerIdx = i;
          foundHeaders = row;
          break;
        }
      }

      if (headerIdx === -1) {
        throw new Error("Não foi encontrada a linha de cabeçalhos no EXPORT_TRANSPORTES. Verifique se as colunas obrigatórias existem (pode ser necessário remover linhas de topo manuais se o formato for muito complexo).");
      }

      const rowsData = exportRaw.slice(headerIdx + 1);
      const amountDiagnostics: { raw: any, parsed: number }[] = [];

      const mappedTransport: TransportRecord[] = rowsData
        .filter(row => row.length > 0 && row.some(cell => cell !== null && cell !== ''))
        .map(row => {
          const obj: any = {};
          foundHeaders.forEach((h, idx) => {
            if (h) obj[h.trim()] = row[idx];
          });

          // ROBUST AMOUNT PARSING
          const rawVal = obj["Montante em moeda interna"];
          let parsedVal = 0;
          if (typeof rawVal === 'number') {
            parsedVal = rawVal;
          } else if (typeof rawVal === 'string') {
            const trimmed = rawVal.trim().replace(/[^\d,.+-]/g, ''); // Remove currency symbols or spaces
            if (trimmed.includes('.') && trimmed.includes(',')) {
              // Format 1.234,56
              parsedVal = parseFloat(trimmed.replace(/\./g, '').replace(',', '.'));
            } else if (trimmed.includes(',')) {
              // Format 1234,56
              parsedVal = parseFloat(trimmed.replace(',', '.'));
            } else {
              // Format 1234.56 or 1234
              parsedVal = parseFloat(trimmed);
            }
          }
          
          if (amountDiagnostics.length < 3) {
            amountDiagnostics.push({ raw: rawVal, parsed: parsedVal });
          }

          return {
            Cliente: normalizeKey(obj["Cliente"]),
            'Data do documento': obj["Data do documento"] || '',
            Referência: String(obj["Referência"] || '').trim(),
            'Montante em moeda interna': parsedVal,
            Texto: String(obj["Texto"] || '').trim(),
            'Chave referência 3': normalizeKey(obj["Chave referência 3"])
          };
        });

      // Sanity check for x100 bug
      const suspicous = mappedTransport.filter(r => Math.abs(r['Montante em moeda interna']) > 100000);
      if (suspicous.length > (mappedTransport.length * 0.5) && mappedTransport.length > 10) {
        addLog("[AVISO CRÍTICO] Detetados montantes invulgarmente elevados. Verifique se o separador decimal foi processado corretamente.");
      }

      // 2. Parse FirmasTransportes_Emails
      const firmsBuf = await firmsFile.arrayBuffer();
      const firmsWb = XLSX.read(firmsBuf);
      const firmsSheet = firmsWb.Sheets['Folha1'] || firmsWb.Sheets[firmsWb.SheetNames[0]];
      const firmsRaw = XLSX.utils.sheet_to_json(firmsSheet, { raw: false }) as any[];
      
      const mappedFirms: FirmRecord[] = firmsRaw.map(row => {
        const findVal = (keyBase: string) => {
          const key = Object.keys(row).find(k => k.trim().toLowerCase() === keyBase.toLowerCase());
          return key ? row[key] : undefined;
        };

        return {
          Cod: normalizeKey(findVal("Cod")),
          Nome: String(findVal("Nome") || '').trim(),
          Para1: findVal("Para1"),
          Para2: findVal("Para2"),
          Para3: findVal("Para3"),
          Para4: findVal("Para4"),
          Para5: findVal("Para5"),
          Conhecimento1: findVal("Conhecimento1"),
          Conhecimento2: findVal("Conhecimento2"),
          Conhecimento3: findVal("Conhecimento3"),
          Conhecimento4: findVal("Conhecimento4"),
          Conhecimento5: findVal("Conhecimento5"),
          Conhecimento6: findVal("Conhecimento6"),
        };
      });

      setTransportData(mappedTransport);
      setFirmsData(mappedFirms);
      setDiagnostics(prev => ({
        ...prev,
        headerRow: headerIdx + 1,
        columns: foundHeaders.filter(h => h !== ''),
        exportRows: mappedTransport.length,
        firmsRows: mappedFirms.length,
        amountSamples: amountDiagnostics
      }));

      addLog(`Leitura concluída. ${mappedTransport.length} linhas de transporte, ${mappedFirms.length} firmas.`);
      if (mappedTransport.length === 0) {
        throw new Error("O ficheiro EXPORT_TRANSPORTES não contém dados válidos após o cabeçalho.");
      }
      setActiveTab('preview');
    } catch (err: any) {
      addLog(`Erro na leitura: ${err.message}`);
      alert(err.message);
    } finally {
      setIsParsing(false);
    }
  };

  const stats = useMemo(() => {
    const clients = new Set(transportData.map(r => normalizeKey(r.Cliente)).filter(c => c !== ''));
    const firmCods = new Set(firmsData.map(f => normalizeKey(f.Cod)));
    const unmapped = Array.from(clients).filter(c => !firmCods.has(c));
    return {
      totalRows: transportData.length,
      uniqueClients: clients.size,
      unmappedCount: unmapped.length,
      unmappedList: unmapped
    };
  }, [transportData, firmsData]);

  const processEverything = async () => {
    setIsProcessing(true);
    setActiveTab('processing');
    setLogs([]);
    setErrors([]);
    setProgress(0);
    addLog('A iniciar processamento em massa...');

    const generatedPdfs: typeof results.pdfs = [];
    const generatedEmls: typeof results.emls = [];
    const unmappedCods: string[] = [];

    const clients = Array.from(new Set(transportData.map(r => normalizeKey(r.Cliente)).filter(c => c !== '')));
    const total = clients.length;

    for (let i = 0; i < total; i++) {
      const cod = clients[i];
      const firm = firmsData.find(f => normalizeKey(f.Cod) === cod);
      const records = transportData.filter(r => normalizeKey(r.Cliente) === cod);

      if (!firm) {
        addLog(`[AVISO] Cod ${cod} não encontrado na base de dados de firmas.`);
        unmappedCods.push(cod);
        setErrors(prev => [...prev, { type: 'MAP_MISSING', message: `Cliente ${cod} sem correspondência na base de dados.`, cod }]);
        
        const fakeFirm: FirmRecord = { Cod: cod, Nome: '(Desconhecido)' };
        try {
          const pdfBlob = await generateTransportPdf(cod, fakeFirm, records);
          const pdfName = cleanFilename(`Relatório Transp. Aberto ${cod} Desconhecido.pdf`);
          generatedPdfs.push({ cod, name: pdfName, blob: pdfBlob });
        } catch (e: any) {
          setErrors(prev => [...prev, { type: 'MIXING_DETECTED', message: e.message, cod }]);
        }
      } else {
        try {
          const pdfBlob = await generateTransportPdf(cod, firm, records);
          const pdfName = cleanFilename(`Relatório Transp. Aberto ${firm.Cod} ${firm.Nome}.pdf`);
          generatedPdfs.push({ cod, name: pdfName, blob: pdfBlob });

          if (normalizeKey(firm.Cod) !== cod) {
             throw new Error(`Sanity check falhou: tentativa de anexar PDF do cod ${cod} ao email do cod ${firm.Cod}`);
          }

          const emlBlob = await generateEml(firm, pdfBlob, pdfName);
          const emlName = cleanFilename(`Email Draft ${firm.Cod} ${firm.Nome}.eml`);
          generatedEmls.push({ cod, name: emlName, blob: emlBlob });
          
          addLog(`Processado: ${cod} - ${firm.Nome}`);
        } catch (e: any) {
          addLog(`[ERRO] Falha no processamento de ${cod}: ${e.message}`);
          setErrors(prev => [...prev, { type: 'MIXING_DETECTED', message: e.message, cod }]);
        }
      }
      setProgress(Math.round(((i + 1) / total) * 100));
    }

    setResults({ pdfs: generatedPdfs, emls: generatedEmls, unmapped: unmappedCods });
    setIsProcessing(false);
    setActiveTab('summary');
    addLog('Processamento concluído.');
  };

  const downloadZip = async () => {
    const zip = new JSZip();
    const pdfFolder = zip.folder("PDFs");
    const emlFolder = zip.folder("Emails");

    results.pdfs.forEach(p => pdfFolder?.file(p.name, p.blob));
    results.emls.forEach(e => emlFolder?.file(e.name, e.blob));

    const content = await zip.generateAsync({ type: "blob" });
    saveAs(content, `Export_Transportes_${new Date().toISOString().split('T')[0]}.zip`);
  };

  const copyToClipboard = (text: string) => {
    navigator.clipboard.writeText(text);
  };
  // Exportar apenas o PDF correspondente ao cod (Transportista)
const downloadSinglePdf = (cod: string) => {
  const pdf = results.pdfs.find(p => p.cod === cod);

  if (!pdf) {
    alert('PDF não encontrado para este código. Confirma se já foi processado.');
    return;
  }

  const url = URL.createObjectURL(pdf.blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = pdf.name;
  document.body.appendChild(a);
  a.click();
  a.remove();

  // evita leaks
  setTimeout(() => URL.revokeObjectURL(url), 0);
};

  return (
    <div className="min-h-screen py-6 px-4 sm:px-6 lg:px-8 max-w-6xl mx-auto">
      <header className="mb-6 flex items-center justify-between border-b border-slate-300 pb-4">
        <div className="flex items-center space-x-3">
          <div className="p-2 bg-[#2F5F8F] text-white rounded shadow-sm">
            <FileBox className="w-6 h-6" />
          </div>
          <div>
            <h1 className="text-xl font-bold tracking-tight text-[#1B1F23]">Transportes | PDF & Email</h1>
            <p className="text-[11px] text-slate-500 uppercase tracking-wide">SAP Interface Edition</p>
          </div>
        </div>
        <div className="hidden sm:block text-right">
          <p className="text-[11px] text-slate-500">Versão 3.5.0</p>
          <p className="text-[10px] text-slate-400">© 2026 SumolCompal</p>
        </div>
      </header>

      <nav className="flex items-center space-x-px border border-slate-300 bg-white p-0.5 rounded shadow-sm mb-6 overflow-hidden">
        {[
          { id: 'upload', label: '1. Ficheiros', icon: Upload },
          { id: 'preview', label: '2. Validar', icon: ShieldAlert },
          { id: 'processing', label: '3. Processar', icon: Play },
          { id: 'summary', label: '4. Resumo', icon: CheckCircle2 },
        ].map((tab) => (
          <button
            key={tab.id}
            onClick={() => (activeTab === 'summary' || activeTab === 'preview' || tab.id === 'upload') && setActiveTab(tab.id as Tab)}
            disabled={(tab.id === 'processing' && !isProcessing) || (tab.id === 'summary' && results.pdfs.length === 0)}
            className={`
              flex-1 flex items-center justify-center space-x-2 px-3 py-1.5 text-xs font-semibold transition-all
              ${activeTab === tab.id ? 'bg-[#2F5F8F] text-white' : 'text-[#4B5563] hover:bg-[#EAF2FF]'}
              disabled:opacity-40 disabled:cursor-not-allowed border-r border-slate-200 last:border-r-0
            `}
          >
            <tab.icon className="w-3.5 h-3.5" />
            <span>{tab.label}</span>
          </button>
        ))}
      </nav>

      <main className="sap-card overflow-hidden min-h-[450px]">
        <AnimatePresence mode="wait">
          {activeTab === 'upload' && (
            <motion.div 
              key="upload"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              className="p-6"
            >
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div className="space-y-3">
                  <label className="block text-[11px] font-bold text-slate-500 uppercase tracking-wider">1. EXPORT_TRANSPORTES.xlsx</label>
                  <div 
                    className={`
                      relative group cursor-pointer border rounded p-6 transition-all text-center
                      ${exportFile ? 'border-indigo-400 bg-indigo-50/30' : 'border-slate-300 bg-white hover:border-indigo-400'}
                    `}
                  >
                    <input type="file" accept=".xlsx" onChange={(e) => handleFileUpload(e, 'export')} className="absolute inset-0 w-full h-full opacity-0 cursor-pointer" />
                    <FileText className={`w-10 h-10 mx-auto mb-3 ${exportFile ? 'text-indigo-600' : 'text-slate-400'}`} />
                    <p className="text-xs font-bold text-slate-900">{exportFile ? exportFile.name : 'Selecionar Export de Transportes'}</p>
                    <p className="text-[10px] text-slate-500 mt-1 uppercase">Colunas: Cliente, Data, Referência, Montante...</p>
                  </div>
                </div>

                <div className="space-y-3">
                  <label className="block text-[11px] font-bold text-slate-500 uppercase tracking-wider">2. FirmasTransportes_Emails.xlsx</label>
                  <div 
                    className={`
                      relative group cursor-pointer border rounded p-6 transition-all text-center
                      ${firmsFile ? 'border-emerald-400 bg-emerald-50/30' : 'border-slate-300 bg-white hover:border-emerald-400'}
                    `}
                  >
                    <input type="file" accept=".xlsx" onChange={(e) => handleFileUpload(e, 'firms')} className="absolute inset-0 w-full h-full opacity-0 cursor-pointer" />
                    <Mail className={`w-10 h-10 mx-auto mb-3 ${firmsFile ? 'text-emerald-600' : 'text-slate-400'}`} />
                    <p className="text-xs font-bold text-slate-900">{firmsFile ? firmsFile.name : 'Selecionar Base de Dados de Emails'}</p>
                    <p className="text-[10px] text-slate-500 mt-1 uppercase">Colunas: Cod, Nome, Para1-5, Conhecimento...</p>
                  </div>
                </div>
              </div>

              <div className="mt-8 flex justify-center">
                <button
                  onClick={parseExcels}
                  disabled={!exportFile || !firmsFile || isParsing}
                  className="sap-btn-primary px-10 shadow-sm"
                >
                  {isParsing ? <Loader2 className="w-4 h-4 animate-spin mr-2" /> : <ChevronRight className="w-4 h-4 mr-2" />}
                  <span>Carregar Ficheiros</span>
                </button>
              </div>
            </motion.div>
          )}

          {activeTab === 'preview' && (
            <motion.div 
              key="preview"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              className="p-6"
            >
              <div className="grid grid-cols-1 md:grid-cols-3 gap-px bg-slate-300 border border-slate-300 mb-6 overflow-hidden rounded">
                <div className="p-4 bg-white">
                  <p className="text-[10px] font-bold text-slate-500 uppercase tracking-wider mb-1">Total Registos</p>
                  <p className="text-2xl font-bold text-[#2F5F8F]">{stats.totalRows}</p>
                </div>
                <div className="p-4 bg-white">
                  <p className="text-[10px] font-bold text-slate-500 uppercase tracking-wider mb-1">Transportistas</p>
                  <p className="text-2xl font-bold text-[#2F5F8F]">{stats.uniqueClients}</p>
                </div>
                <div className="p-4 bg-white">
                  <p className="text-[10px] font-bold text-slate-500 uppercase tracking-wider mb-1">Sem Contacto</p>
                  <p className="text-2xl font-bold text-[#B00020]">{stats.unmappedCount}</p>
                </div>
              </div>

              <div className="bg-[#1B1F23] rounded p-4 mb-6 text-slate-300 font-mono text-[11px] border border-black shadow-sm">
                <h3 className="text-[#FFEB3B] font-bold mb-3 flex items-center space-x-2 uppercase text-[10px] tracking-widest">
                  <AlertCircle className="w-3.5 h-3.5" />
                  <span>Painel de Diagnóstico</span>
                </h3>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div className="space-y-1">
                    <p><span className="text-slate-500">Linha Header (Export):</span> {diagnostics.headerRow}</p>
                    <p><span className="text-slate-500">Registos Export:</span> {diagnostics.exportRows}</p>
                    <p><span className="text-slate-500">Registos Firmas:</span> {diagnostics.firmsRows}</p>
                    <div className="flex space-x-4 mt-2">
                      <p><span className="text-slate-500">jsPDF:</span> {diagnostics.libHealth.jspdf ? <span className="text-emerald-400">OK</span> : <span className="text-red-400">ERRO</span>}</p>
                      <p><span className="text-slate-500">AutoTable:</span> {diagnostics.libHealth.autotable ? <span className="text-emerald-400">OK</span> : <span className="text-red-400">ERRO</span>}</p>
                    </div>
                  </div>
                  <div className="space-y-1">
                    <p className="text-slate-500">Colunas Detetadas:</p>
                    <div className="flex flex-wrap gap-1 mt-1 mb-3">
                      {diagnostics.columns.map((c, i) => (
                        <span key={i} className="px-1.5 py-0.5 bg-slate-800 rounded border border-slate-700 text-[10px]">{c}</span>
                      ))}
                    </div>
                    {diagnostics.amountSamples.length > 0 && (
                      <div className="mt-4 p-3 bg-slate-800/50 rounded-xl border border-slate-700">
                        <p className="text-indigo-400 font-bold mb-2 uppercase text-[10px]">Amostras de Montantes (Parse)</p>
                        <div className="space-y-1 text-[10px]">
                          {diagnostics.amountSamples.map((s, i) => (
                            <div key={i} className="flex justify-between border-b border-slate-700 pb-1 last:border-0">
                              <span className="text-slate-400">Original: "{String(s.raw)}"</span>
                              <span className="text-emerald-400">Parsed: {s.parsed.toFixed(2)}</span>
                            </div>
                          ))}
                        </div>
                      </div>
                    )}
                  </div>
                </div>
              </div>

              {stats.unmappedCount > 0 && (
                <div className="bg-orange-50 border border-orange-200 rounded-2xl p-6 mb-8">
                  <div className="flex items-start space-x-4">
                    <AlertCircle className="w-6 h-6 text-orange-600 flex-shrink-0 mt-1" />
                    <div>
                      <h3 className="font-bold text-orange-900">Aviso: Códigos não mapeados</h3>
                      <p className="text-sm text-orange-700 mb-3">Encontramos códigos no export que não estão no ficheiro de firmas. Serão gerados PDFs genéricos.</p>
                      <div className="flex flex-wrap gap-2 text-xs font-mono">
                        {stats.unmappedList.map(c => (
                          <span key={c} className="px-2 py-1 bg-orange-100 border border-orange-200 rounded text-orange-800">{c}</span>
                        ))}
                      </div>
                    </div>
                  </div>
                </div>
              )}

              <div className="flex justify-center space-x-3">
                <button onClick={() => setActiveTab('upload')} className="sap-btn-secondary px-6">Voltar</button>
                <button
                  onClick={processEverything}
                  className="sap-btn-primary px-10"
                >
                  <Play className="w-4 h-4 mr-2" />
                  <span>PROCESSAR TUDO</span>
                </button>
              </div>
            </motion.div>
          )}

          {activeTab === 'processing' && (
            <motion.div 
              key="processing"
              className="p-12 text-center"
            >
              <div className="max-w-md mx-auto py-8">
                <Loader2 className="w-12 h-12 text-[#2F5F8F] animate-spin mx-auto mb-6" />
                <h2 className="text-lg font-bold text-slate-900 mb-6 uppercase tracking-wider">A Gerar PDFs e Emails...</h2>
                <div className="w-full bg-slate-200 h-2 rounded overflow-hidden mb-3 border border-slate-300">
                  <motion.div 
                    className="h-full bg-[#2F5F8F]"
                    initial={{ width: 0 }}
                    animate={{ width: `${progress}%` }}
                  />
                </div>
                <p className="text-xs font-bold text-[#2F5F8F]">{progress}% concluído</p>
                
                <div className="mt-8 text-left bg-[#1B1F23] rounded p-3 font-mono text-[10px] text-slate-400 h-40 overflow-y-auto border border-black shadow-inner">
                  {logs.map((log, i) => (
                    <div key={i} className="mb-0.5 border-l border-slate-700 pl-2">{log}</div>
                  ))}
                  {isProcessing && <div className="animate-pulse">_</div>}
                </div>
              </div>
            </motion.div>
          )}

          {activeTab === 'summary' && (
            <motion.div 
              key="summary"
              className="p-6"
            >
              <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 mb-8 border-b border-slate-200 pb-6">
                <div>
                  <h2 className="text-xl font-bold text-[#1B1F23]">Processamento Concluído</h2>
                  <p className="text-xs text-slate-500 uppercase tracking-wide">{results.pdfs.length} Relatórios PDF Gerados | {results.emls.length} Rascunhos de Email</p>
                </div>
                <button
                  onClick={downloadZip}
                  className="sap-btn-primary h-12 px-8 flex items-center space-x-2"
                >
                  <Download className="w-5 h-5" />
                  <span className="text-sm font-bold">BAIXAR ZIP COMPLETO</span>
                </button>
              </div>

              <div className="space-y-6">
                {errors.length > 0 && (
                  <section>
                    <h3 className="text-[10px] font-bold text-[#B00020] uppercase tracking-widest mb-3 flex items-center space-x-2">
                       <ShieldAlert className="w-3.5 h-3.5" />
                       <span>Relatório de Erros / Alertas</span>
                    </h3>
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-2">
                      {errors.map((err, i) => (
                        <div key={i} className="p-2 bg-[#FFF3CD] border border-[#FFEEBA] rounded flex items-start space-x-3 text-[11px]">
                          <AlertCircle className="w-4 h-4 text-[#B00020] mt-0.5 flex-shrink-0" />
                          <div>
                            <span className="font-bold text-[#1B1F23] mr-1">{err.cod || 'ERRO'}:</span>
                            <span className="text-slate-700">{err.message}</span>
                          </div>
                        </div>
                      ))}
                    </div>
                  </section>
                )}

                <section>
                  <div className="flex items-center justify-between mb-3 sap-table-header p-2 rounded-t">
  <div className="flex items-center space-x-4">
    <h3 className="text-[11px] font-bold uppercase tracking-widest flex items-center space-x-2">
      <Mail className="w-3.5 h-3.5" />
      <span>Rascunhos de Email (Outlook Drafts)</span>
    </h3>
  </div>

  <button
    onClick={resetSentEmails}
    className="text-[9px] text-slate-500 hover:text-red-600 transition-colors flex items-center space-x-1 border border-slate-300 bg-white px-2 py-0.5 rounded shadow-sm"
    title="Limpar marcações de enviado"
    type="button"
  >
    <Trash2 className="w-3 h-3" />
    <span>Limpar Histórico de Envio</span>
  </button>
</div>
                  
                  <div className="border border-slate-300 bg-white rounded-b overflow-hidden max-h-[380px] overflow-y-auto">
                    {results.emls.map((eml, i) => {
                      const firm = firmsData.find(f => normalizeKey(f.Cod) === eml.cod);
                      if (!firm) return null;
                      const to = [firm.Para1, firm.Para2, firm.Para3, firm.Para4, firm.Para5].filter(x => x).join(', ');
                      const cc = [firm.Conhecimento1, firm.Conhecimento2, firm.Conhecimento3, firm.Conhecimento4, firm.Conhecimento5, firm.Conhecimento6].filter(x => x).join(', ');
                      const today = new Intl.DateTimeFormat('pt-PT').format(new Date());
                      
                      const mailtoBody = `Exmos. ${firm.Nome}\n\nSegue em anexo ficheiro com os documentos em aberto à data de ${today}\nEstamos disponíveis para qualquer esclarecimento adicional que considerem relevante.\n\nAtentamente\nA equipa AFSN\n\nEm caso de dúvidas contactar faturacao@sumolcompal.pt`;
                      const mailtoUrl = `mailto:${to}?cc=${cc}&subject=${encodeURIComponent(`Relatório de PA´s em aberto de ${firm.Nome}`)}&body=${encodeURIComponent(mailtoBody)}`;
                      const isSent = sentEmails[eml.cod] === true;

                      return (
                        <div key={i} className="sap-table-row p-3 hover:bg-[#EAF2FF] transition-all group">
                          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                            <div className="flex-1">
                              <div className="flex items-center space-x-2 mb-0.5">
                                <span className={`font-bold text-[13px] transition-colors ${isSent ? 'text-slate-400' : 'text-[#1B1F23]'}`}>{firm.Nome}</span>
                                <span className={`text-[11px] px-1.5 border transition-colors rounded ${isSent ? 'bg-slate-50 text-slate-400 border-slate-200' : 'bg-slate-100 text-slate-600 border-slate-300 font-mono'}`}>{firm.Cod}</span>
                                {isSent && <span className="sap-badge-sent text-[10px]"><CheckCircle2 className="w-3 h-3 mr-1" /> ENVIADO</span>}
                              </div>
                              <div className="text-[11px] text-slate-500 truncate max-w-lg">Para: {to}</div>
                            </div>
                            <div className="flex items-center space-x-1">
                              <button onClick={() => copyToClipboard(to)} title="Copiar destinatários" className="sap-btn-secondary p-1">
                                <Copy className="w-3.5 h-3.5" />
                              </button>
                              <a href={mailtoUrl} title="Mailto Link" className="sap-btn-secondary p-1">
                                <ExternalLink className="w-3.5 h-3.5 text-[#2F5F8F]" />
                              </a>
                              <button
  type="button"
  onClick={() => downloadSinglePdf(eml.cod)}
  title="Exportar apenas o PDF"
  className="sap-btn-secondary p-1"
>
  <span className="text-[10px] font-bold">PDF</span>
</button>

                              <a 
                                href={URL.createObjectURL(eml.blob)} 
                                download={eml.name} 
                                onClick={() => toggleEmailSent(eml.cod)}
                                title={isSent ? "Reabrir Rascunho" : "Gerar .EML com Anexo"} 
                                className={`sap-btn-primary px-3 space-x-1.5 transition-all ${isSent ? 'btn-sent opacity-90' : ''}`}
                              >
                                {isSent ? <CheckCircle2 className="w-3.5 h-3.5" /> : <Download className="w-3.5 h-3.5" />}
                                <span className="text-[11px] font-bold uppercase">{isSent ? 'Enviado' : 'Enviar E-mail'}</span>
                              </a>
                            </div>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </section>
              </div>

              <div className="mt-8 flex justify-center border-t border-slate-200 pt-6">
                 <button onClick={() => window.location.reload()} className="text-[11px] font-bold text-slate-400 hover:text-red-600 flex items-center space-x-2 uppercase tracking-widest">
                   <Trash2 className="w-3.5 h-3.5" />
                   <span>Resetar Aplicação</span>
                 </button>
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </main>

      <footer className="mt-8 text-center text-slate-500 text-[10px] uppercase tracking-widest space-y-1 pb-8">
        <p>Desenvolvido por Pedro Gameiro para uso interno (Sumol+Compal)</p>
        <p>Processamento Local Pela Equipa AFSN</p>
      </footer>
    </div>
  );
}
