# Transportes | PDF & Email (SAP Interface Edition)

**Site (GitHub Pages):** https://djpedrog.github.io/Transporte-PDF-Email/

Aplicação web (processamento local no browser) para gerar **PDFs de documentos em aberto** por transportista a partir do `EXPORT_TRANSPORTES.xlsx` e preparar **rascunhos de email (.EML)** com o PDF em anexo, usando a base `FirmasTransportes_Emails.xlsx`.

## Como usar
1. Selecionar:
   - `EXPORT_TRANSPORTES.xlsx`
   - `FirmasTransportes_Emails.xlsx`
2. Clicar **Carregar Ficheiros** → validar
3. Clicar **PROCESSAR TUDO**
4. No **Resumo**:
   - **PDF**: descarrega apenas o PDF dessa entrada
   - **Enviar E-mail**: descarrega o `.EML` com PDF anexado (abre no Outlook Web e clica **Enviar**)
   - **ZIP Completo**: descarrega PDFs + EMLs separados por pastas

## Entradas e separação por centros
Alguns transportistas geram entradas adicionais por “centro” (detetado na coluna **Referência**), mantendo também a versão **ALL** (tudo junto):

- **92000923** → `1090/` e `1330/`
- **92000112** → `1020/` e `1010/`

Nas variantes por centro (ex.: `1090/`), o PDF inclui:
- linhas cujo **Referência** contém o centro
- **+** linhas do mesmo cliente com a mesma **Chave referência 3** (para manter o agrupamento consistente)

O assunto do `.EML` inclui o identificador do centro (ex.: `(1090)`); na versão ALL mantém o assunto normal.

## Estado “ENVIADO”
O estado “ENVIADO” é controlado por entrada (ALL vs centros) e guarda-se localmente no browser.
Ao carregar novos ficheiros, o histórico é limpo automaticamente.
Existe botão **Limpar Histórico de Envio** no Resumo.

## Saídas
- PDFs: `Relatório Transp. Aberto <Cod> <Nome> [- <Centro>].pdf`
- EMLs: `Email Draft <Cod> <Nome> [- <Centro>].eml`
- ZIP: `Export_Transportes_YYYY-MM-DD.zip`

## Privacidade / Boas práticas
Não versionar dados reais no repositório (Excels, PDFs, EMLs, ZIPs).

