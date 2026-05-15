import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';

export interface TransportRecord {
  Cliente: string;
  'Data do documento': any;
  Referência: string;
  'Montante em moeda interna': number;
  Texto: string;
  'Chave referência 3': string;
}

export interface FirmRecord {
  Cod: string;
  Nome: string;
  Para1?: string;
  Para2?: string;
  Para3?: string;
  Para4?: string;
  Para5?: string;
  Conhecimento1?: string;
  Conhecimento2?: string;
  Conhecimento3?: string;
  Conhecimento4?: string;
  Conhecimento5?: string;
  Conhecimento6?: string;
}

export interface ProcessingError {
  type: 'MAP_MISSING' | 'MIXING_DETECTED' | 'SANITY_FAIL' | 'INVALID_EMAIL' | 'LIB_ERROR';
  message: string;
  cod?: string;
}

export function checkLibraries(): { jspdf: boolean; autotable: boolean } {
  return {
    jspdf: typeof jsPDF !== 'undefined',
    autotable: typeof autoTable === 'function'
  };
}

export function normalizeKey(val: any): string {
  if (val === null || val === undefined) return '';
  return String(val).trim();
}

export function formatCurrencyValue(val: number): string {
  return new Intl.NumberFormat('pt-PT', {
    style: 'decimal',
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(val);
}

export function cleanFilename(name: string): string {
  let cleaned = name.replace(/[\\/:*?"<>|]/g, '');
  cleaned = cleaned.replace(/\s+/g, ' ').trim();
  return cleaned.substring(0, 120);
}

// Helpers for Date Formatting
export function formatDateDDMMYYYY(val: any): string {
  if (!val) return '';
  // Try to parse as date. If it's a number (Excel date), it might need conversion but usually XLSX handles it if raw:false is used.
  // If it's already a string, try to check if it's already formatted
  const d = new Date(val);
  if (isNaN(d.getTime())) return String(val).trim();
  const day = String(d.getDate()).padStart(2, '0');
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const year = d.getFullYear();
  return `${day}/${month}/${year}`;
}

export function formatDateTimeDDMMYYYY(date: Date): string {
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = date.getFullYear();
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  return `${day}/${month}/${year}, ${hours}:${minutes}`;
}

const SAP_COLORS = {
  BLUE: [0, 103, 160] as [number, number, number],
  NEGATIVE: [187, 0, 0] as [number, number, number],
  HIGHLIGHT: [239, 217, 166] as [number, number, number],
  GRAY_LINE: [204, 204, 204] as [number, number, number],
  TEXT_WHITE: [255, 255, 255] as [number, number, number],
  TEXT_BLACK: [51, 51, 51] as [number, number, number],
};

export async function generateTransportPdf(
  cod: string,
  firm: FirmRecord,
  records: TransportRecord[]
): Promise<Blob> {
  if (typeof autoTable !== 'function') {
    throw new Error('Biblioteca AutoTable não carregada ou incompatível.');
  }

  // A4, portrait, unit 'pt'
  const doc = new jsPDF({
    orientation: 'portrait',
    unit: 'pt',
    format: 'a4'
  });
  
  const margin = 36;
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();

  const now = new Date();
  const generationFormatted = formatDateTimeDDMMYYYY(now);

  // 1. Anti-mixing Validation
  for (const record of records) {
    if (normalizeKey(record.Cliente) !== cod) {
      throw new Error(`Mistura de códigos detetada no PDF ${cod}: cliente inesperado ${record.Cliente}`);
    }
  }

  // Header - Simple Layout Reverted
  doc.setFontSize(20);
  doc.setFont('helvetica', 'bold');
  doc.setTextColor(SAP_COLORS.TEXT_BLACK[0], SAP_COLORS.TEXT_BLACK[1], SAP_COLORS.TEXT_BLACK[2]);
  doc.text('Relatório de PA´s em aberto', margin, 50);
  
  doc.setFontSize(10);
  doc.setFont('helvetica', 'normal');
  doc.text(`Cod: ${firm.Cod}`, margin, 70);
  doc.text(`Nome: ${firm.Nome}`, margin, 83);
  doc.text(`Gerado em: ${generationFormatted}`, margin, 96);

  let currentY = 120;

  // Group by "Chave referência 3"
  const groups: Record<string, TransportRecord[]> = {};
  records.forEach(r => {
    const key = normalizeKey(r['Chave referência 3']) || '(Sem Chave)';
    if (!groups[key]) groups[key] = [];
    groups[key].push(r);
  });

  const sortedGroupKeys = Object.keys(groups).sort((a, b) => {
    if (a === '(Sem Chave)') return 1;
    if (b === '(Sem Chave)') return -1;
    return a.localeCompare(b);
  });

  let totalCliente = 0;

  sortedGroupKeys.forEach((groupKey, groupIdx) => {
    const groupRecords = groups[groupKey].sort((a, b) => {
      const dateA = a['Data do documento'] || '';
      const dateB = b['Data do documento'] || '';
      if (dateA !== dateB) return String(dateA).localeCompare(String(dateB));
      return (a.Referência || '').localeCompare(b.Referência || '');
    });

    const subtotalGroup = groupRecords.reduce((sum, r) => sum + (Number(r['Montante em moeda interna']) || 0), 0);
    totalCliente += subtotalGroup;

    // Check vertical space before group title
    if (currentY > pageHeight - 100) {
      doc.addPage();
      currentY = 50;
    }

    doc.setFontSize(10);
    doc.setFont('helvetica', 'bold');
    doc.text(`Transporte: ${groupKey}`, margin, currentY);
    currentY += 10;

    const tableData = groupRecords.map(r => [
      firm.Cod,
      firm.Nome,
      formatDateDDMMYYYY(r['Data do documento']),
      r.Referência,
      formatCurrencyValue(r['Montante em moeda interna']),
      r.Texto
    ]);

    // Insert Group Total row into the table data (Simplified to just "TOTAL")
    tableData.push([
      '', // Cod
      '', // Nome
      '', // Data
      'TOTAL', // Referência/Label (Just "TOTAL")
      formatCurrencyValue(subtotalGroup), // Montante
      ''  // Texto
    ]);

    autoTable(doc, {
      startY: currentY,
      margin: { left: margin, right: margin },
      head: [['Cod', 'Nome', 'Data Doc.', 'Referência', 'Montante', 'Texto']],
      body: tableData,
      theme: 'striped',
      styles: { 
        fontSize: 8, 
        cellPadding: 3, 
        overflow: 'linebreak', 
        valign: 'top',
        lineColor: SAP_COLORS.GRAY_LINE,
        lineWidth: 0.1,
      },
      headStyles: { 
        fillColor: SAP_COLORS.BLUE, 
        textColor: SAP_COLORS.TEXT_WHITE, 
        fontStyle: 'bold' 
      },
      alternateRowStyles: {
        fillColor: [250, 250, 250]
      },
      columnStyles: {
        0: { cellWidth: 45 },
        1: { cellWidth: 90 },
        2: { cellWidth: 55 },
        3: { cellWidth: 70 },
        4: { cellWidth: 70, halign: 'right' },
        5: { cellWidth: 'auto' },
      },
      didParseCell: (data) => {
        // Highlighting for the Group Total row
        if (data.section === 'body' && data.row.index === tableData.length - 1) {
          data.cell.styles.fontStyle = 'bold';
          data.cell.styles.fillColor = SAP_COLORS.HIGHLIGHT;
          // Apply thick top border for total row
          data.cell.styles.lineWidth = { top: 1.5, bottom: 0.1, left: 0.1, right: 0.1 };
        }

        if (data.section === 'body' && data.column.index === 4) {
          // Negative check
          const valStr = String(data.cell.raw).replace(/\./g, '').replace(',', '.');
          const val = parseFloat(valStr);
          if (val < 0) {
            data.cell.styles.textColor = SAP_COLORS.NEGATIVE;
          }
        }
      },
      didDrawPage: (data: any) => {
        const str = "Página " + doc.internal.getNumberOfPages();
        doc.setFontSize(8);
        doc.setTextColor(150);
        doc.text(str, margin, pageHeight - 20);
      }
    });

    currentY = (doc as any).lastAutoTable.finalY + 20;

    if (currentY > pageHeight - 60 && groupIdx < sortedGroupKeys.length - 1) {
      doc.addPage();
      currentY = 50;
    }
  });

  // Final Total
  if (currentY > pageHeight - 50) {
    doc.addPage();
    currentY = 50;
  }
  
  // Highlight background for final total
  doc.setFillColor(SAP_COLORS.HIGHLIGHT[0], SAP_COLORS.HIGHLIGHT[1], SAP_COLORS.HIGHLIGHT[2]);
  doc.rect(margin, currentY - 12, pageWidth - (margin * 2), 20, 'F');

  doc.setFontSize(11);
  doc.setFont('helvetica', 'bold');
  doc.setTextColor(SAP_COLORS.TEXT_BLACK[0], SAP_COLORS.TEXT_BLACK[1], SAP_COLORS.TEXT_BLACK[2]);
  doc.setLineWidth(2);
  doc.setDrawColor(SAP_COLORS.GRAY_LINE[0], SAP_COLORS.GRAY_LINE[1], SAP_COLORS.GRAY_LINE[2]);
  doc.line(margin, currentY - 12, pageWidth - margin, currentY - 12);
  doc.text(`TOTAL DO CLIENTE: ${formatCurrencyValue(totalCliente)}`, pageWidth - margin - 5, currentY + 2, { align: 'right' });

  return doc.output('blob');
}

export function generateEml(
  firm: FirmRecord,
  pdfBlob: Blob,
  pdfName: string,
  subjectSuffix?: string
): Promise<Blob> {
  const to = [firm.Para1, firm.Para2, firm.Para3, firm.Para4, firm.Para5]
    .filter(e => e && e.trim() !== '')
    .join(', ');
  
  const cc = [firm.Conhecimento1, firm.Conhecimento2, firm.Conhecimento3, firm.Conhecimento4, firm.Conhecimento5, firm.Conhecimento6]
    .filter(e => e && e.trim() !== '')
    .join(', ');

  const now = new Date();
  const dateFormatted = now.toUTCString();
  const today = formatDateDDMMYYYY(now);
  const suffix = subjectSuffix ? ` (${subjectSuffix})` : '';
  const subject = `Relatório de PA´s em aberto de ${firm.Nome}${suffix}`;
  
  const bodyText = `Exmos. ${firm.Nome}\r\n\r\nSegue em anexo ficheiro com os documentos em aberto à data de ${today}\r\nEstamos disponíveis para qualquer esclarecimento adicional que considerem relevante.\r\n\r\nAtentamente\r\nA equipa AFSN\r\n\r\nEm caso de dúvidas contactar faturacao@sumolcompal.pt`;
  
  const bodyHtml = `
    <html>
    <body style="font-family: sans-serif; font-size: 14px; color: #333; line-height: 1.5;">
      <p>Exmos. <strong>${firm.Nome}</strong></p>
      <p>Segue em anexo ficheiro com os documentos em aberto à data de <strong>${today}</strong></p>
      <p>Estamos disponíveis para qualquer esclarecimento adicional que considerem relevante.</p>
      <p>Atentamente,<br><strong>A equipa AFSN</strong></p>
      <br>
      <div style="background-color: #FFF3CD; border: 1px solid #FFEEBA; padding: 12px; border-radius: 6px; font-size: 12px; color: #333;">
        <strong>Aviso:</strong> Em caso de dúvidas contactar <a href="mailto:faturacao@sumolcompal.pt" style="color: #856404; font-weight: bold;">faturacao@sumolcompal.pt</a>
      </div>
    </body>
    </html>
  `.trim();

  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onload = () => {
      const base64 = (reader.result as string).split(',')[1];
      const mixedBoundary = `----=_Mixed_Part_${Date.now()}_${Math.random().toString(36).substring(2)}`;
      const altBoundary = `----=_Alt_Part_${Date.now()}_${Math.random().toString(36).substring(2)}`;
      
      const eml = [
        `X-Unsent: 1`,
        `Date: ${dateFormatted}`,
        `From: "AutoTransp App" <noreply@local.app>`,
        `To: ${to}`,
        `Cc: ${cc}`,
        `Subject: ${subject}`,
        `MIME-Version: 1.0`,
        `Content-Type: multipart/mixed; boundary="${mixedBoundary}"`,
        ``,
        `--${mixedBoundary}`,
        `Content-Type: multipart/alternative; boundary="${altBoundary}"`,
        ``,
        `--${altBoundary}`,
        `Content-Type: text/plain; charset="utf-8"`,
        `Content-Transfer-Encoding: 8bit`,
        ``,
        bodyText,
        ``,
        `--${altBoundary}`,
        `Content-Type: text/html; charset="utf-8"`,
        `Content-Transfer-Encoding: 8bit`,
        ``,
        bodyHtml,
        ``,
        `--${altBoundary}--`,
        ``,
        `--${mixedBoundary}`,
        `Content-Type: application/pdf; name="${pdfName}"`,
        `Content-Transfer-Encoding: base64`,
        `Content-Disposition: attachment; filename="${pdfName}"`,
        ``,
        base64,
        ``,
        `--${mixedBoundary}--`
      ].join('\r\n');

      resolve(new Blob([eml], { type: 'message/rfc822' }));
    };
    reader.readAsDataURL(pdfBlob);
  });
}
