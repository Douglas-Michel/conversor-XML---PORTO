import * as XLSX from 'xlsx';
import { format } from 'date-fns';
import { NotaFiscal } from './xmlParser';

function getTodayBRDate(): string {
  // Local browser date (avoids UTC shift from toISOString)
  return format(new Date(), 'dd/MM/yyyy');
}

export function exportToExcel(notas: NotaFiscal[], fileName: string = 'notas_fiscais') {
  const today = getTodayBRDate();

  const data = notas.map((nota) => ({
    'Data': today,
    'Tipo NF': nota.tipoOperacao,
    'Fornecedor/Cliente': nota.fornecedorCliente,
    'Nº NF-e': nota.tipo === 'NF-e' ? nota.numero : '',
    'Nº CT-e': nota.numeroCTe || nota.nfeReferenciada || '',
    'Valor': nota.valorTotal,
    'PIS': nota.valorPIS,
    'P': nota.flagPIS ? 'X' : '',
    'Alíq. COF': nota.aliquotaCOFINS,
    'COFINS': nota.valorCOFINS,
    'C': nota.flagCOFINS ? 'X' : '',
    'Alíq. IPI': nota.aliquotaIPI,
    'IPI': nota.valorIPI,
    'I': nota.flagIPI ? 'X' : '',
    'Alíq. ICMS': nota.aliquotaICMS,
    'ICMS': nota.valorICMS,
    'IC': nota.flagICMS ? 'X' : '',
    'Alíq. DIFAL': nota.aliquotaDIFAL,
    'DIFAL': nota.valorDIFAL,
    'Ano': nota.ano,
    'Reduz ICMS': nota.reducaoICMS,
  }));

  const worksheet = XLSX.utils.json_to_sheet(data);
  
  const columnWidths = [
    { wch: 12 },  // Data
    { wch: 10 },  // Tipo NF (Entrada/Saída)
    { wch: 40 },  // Fornecedor/Cliente
    { wch: 12 },  // Nº NF-e
    { wch: 12 },  // Nº CT-e
    { wch: 15 },  // Valor
    { wch: 12 },  // PIS
    { wch: 4 },   // P
    { wch: 10 },  // Alíq. COF
    { wch: 12 },  // COFINS
    { wch: 4 },   // C
    { wch: 10 },  // Alíq. IPI
    { wch: 12 },  // IPI
    { wch: 4 },   // I
    { wch: 10 },  // Alíq. ICMS
    { wch: 12 },  // ICMS
    { wch: 4 },   // IC
    { wch: 10 },  // Alíq. DIFAL
    { wch: 12 },  // DIFAL
    { wch: 6 },   // Ano
    { wch: 12 },  // Reduz ICMS
  ];
  worksheet['!cols'] = columnWidths;

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Notas Fiscais');

  const summary = createSummary(notas);
  const summarySheet = XLSX.utils.json_to_sheet(summary);
  summarySheet['!cols'] = [{ wch: 25 }, { wch: 18 }];
  XLSX.utils.book_append_sheet(workbook, summarySheet, 'Resumo');

  const timestamp = format(new Date(), 'yyyy-MM-dd');
  XLSX.writeFile(workbook, `${fileName}_${timestamp}.xlsx`);
}

function createSummary(notas: NotaFiscal[]) {
  const totalNotas = notas.length;
  const totalNFe = notas.filter(n => n.tipo === 'NF-e').length;
  const totalCTe = notas.filter(n => n.tipo === 'CT-e').length;
  const entradas = notas.filter(n => n.tipoOperacao === 'Entrada');
  const saidas = notas.filter(n => n.tipoOperacao === 'Saída');
  
  const sumValues = (arr: NotaFiscal[]) => ({
    total: arr.reduce((sum, n) => sum + n.valorTotal, 0),
    icms: arr.reduce((sum, n) => sum + n.valorICMS, 0),
    pis: arr.reduce((sum, n) => sum + n.valorPIS, 0),
    cofins: arr.reduce((sum, n) => sum + n.valorCOFINS, 0),
    ipi: arr.reduce((sum, n) => sum + n.valorIPI, 0),
    difal: arr.reduce((sum, n) => sum + n.valorDIFAL, 0),
  });

  const totaisEntrada = sumValues(entradas);
  const totaisSaida = sumValues(saidas);

  return [
    { 'Descrição': 'Total de Documentos', 'Valor': totalNotas },
    { 'Descrição': 'Total NF-e', 'Valor': totalNFe },
    { 'Descrição': 'Total CT-e', 'Valor': totalCTe },
    { 'Descrição': '', 'Valor': '' },
    { 'Descrição': '--- ENTRADAS ---', 'Valor': '' },
    { 'Descrição': 'Qtd. Entradas', 'Valor': entradas.length },
    { 'Descrição': 'Valor Total Entradas', 'Valor': totaisEntrada.total },
    { 'Descrição': 'ICMS Entradas', 'Valor': totaisEntrada.icms },
    { 'Descrição': 'PIS Entradas', 'Valor': totaisEntrada.pis },
    { 'Descrição': 'COFINS Entradas', 'Valor': totaisEntrada.cofins },
    { 'Descrição': 'IPI Entradas', 'Valor': totaisEntrada.ipi },
    { 'Descrição': 'DIFAL Entradas', 'Valor': totaisEntrada.difal },
    { 'Descrição': '', 'Valor': '' },
    { 'Descrição': '--- SAÍDAS ---', 'Valor': '' },
    { 'Descrição': 'Qtd. Saídas', 'Valor': saidas.length },
    { 'Descrição': 'Valor Total Saídas', 'Valor': totaisSaida.total },
    { 'Descrição': 'ICMS Saídas', 'Valor': totaisSaida.icms },
    { 'Descrição': 'PIS Saídas', 'Valor': totaisSaida.pis },
    { 'Descrição': 'COFINS Saídas', 'Valor': totaisSaida.cofins },
    { 'Descrição': 'IPI Saídas', 'Valor': totaisSaida.ipi },
    { 'Descrição': 'DIFAL Saídas', 'Valor': totaisSaida.difal },
  ];
}
