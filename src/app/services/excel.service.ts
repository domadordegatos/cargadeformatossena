import { Injectable } from '@angular/core';
import * as XLSX from 'xlsx';
import PizZip from 'pizzip';
import { saveAs } from 'file-saver';
import { Estudiante, SeguimientoData } from '../models/estudiante.model';

const EXAMPLE_NAME = '(EJEMPLO) PEREZ GARCIA JUAN CAMILO';
const EXAMPLE_DOC = '0000000000';
const EXAMPLE_ROW = [
  '1', EXAMPLE_NAME, '',
  EXAMPLE_DOC, 'TI', '',
  'GARCIA LOPEZ MARIA FERNANDA', 'CC', '37865412', 'BUCARAMANGA',
  'ejemplo@correo.com', '3101234567', '3209876543', 'CRA 10 # 20-30 BARRIO CENTRO',
  '2889741', '15', '07',
  'TECNOLOGIA EN ANALISIS Y DESARROLLO DE SOFTWARE', 'BUCARAMANGA', 'SANTANDER'
];

const HEADERS = [
  'grupo', 'nombre', 'firma',
  'documento', 'tipo', 'otro',
  'acudiente', 'tipopadre', 'docacudiente', 'lugardoc',
  'correo', 'telefono1', 'telefono2', 'direccion',
  'ficha', 'dia', 'mes',
  'programa', 'ciudadmunicipio', 'regional'
];

const COL_WIDTHS: Record<string, number> = {
  grupo: 8, nombre: 35, firma: 14,
  documento: 14, tipo: 8, otro: 8,
  acudiente: 35, tipopadre: 12, docacudiente: 14, lugardoc: 18,
  correo: 28, telefono1: 14, telefono2: 14, direccion: 35,
  ficha: 12, dia: 6, mes: 6,
  programa: 45, ciudadmunicipio: 20, regional: 18
};

interface CellImage {
  buffer: ArrayBuffer;
  extension: string;
}

function colToLetter(idx: number): string {
  let s = '';
  let n = idx;
  while (n >= 0) {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
  }
  return s;
}

function parseXmlAttr(xml: string, regex: RegExp): string[] {
  const results: string[] = [];
  let m;
  while ((m = regex.exec(xml)) !== null) {
    results.push(m[1]);
  }
  return results;
}

function extractCellImages(data: Uint8Array, colIndex: number): Map<number, CellImage> {
  const result = new Map<number, CellImage>();
  try {
    const zip = new PizZip(data);

    const sheetXml = zip.file('xl/worksheets/sheet1.xml')?.asText();
    if (!sheetXml) return result;

    const colLetter = colToLetter(colIndex);
    const cellVmMap = new Map<number, number>();
    const cellRegex = new RegExp('<c r="' + colLetter + '(\\d+)"[^>]*vm="(\\d+)"', 'g');
    let match;
    while ((match = cellRegex.exec(sheetXml)) !== null) {
      cellVmMap.set(parseInt(match[1], 10) - 2, parseInt(match[2], 10));
    }
    if (!cellVmMap.size) return result;

    const metaXml = zip.file('xl/metadata.xml')?.asText();
    const rvXml = zip.file('xl/richData/rdrichvalue.xml')?.asText();
    const rvRelXml = zip.file('xl/richData/richValueRel.xml')?.asText();
    const rvRelsXml = zip.file('xl/richData/_rels/richValueRel.xml.rels')?.asText();
    if (!metaXml || !rvXml || !rvRelXml || !rvRelsXml) return result;

    // vm (1-indexed) -> richdata value index
    const rvbMatches = parseXmlAttr(metaXml, /rvb i="(\d+)"/g);
    const vmToRv = new Map<number, number>();
    rvbMatches.forEach((val, i) => vmToRv.set(i + 1, parseInt(val, 10)));

    // richdata value index -> rel index (first <v> in each <rv>)
    const rvToRel = new Map<number, number>();
    const rvRegex = /<rv s="\d+">([\s\S]*?)<\/rv>/g;
    let rvIdx = 0;
    let rvMatch;
    while ((rvMatch = rvRegex.exec(rvXml)) !== null) {
      const firstV = /<v>(\d+)<\/v>/.exec(rvMatch[1]);
      if (firstV) rvToRel.set(rvIdx, parseInt(firstV[1], 10));
      rvIdx++;
    }

    // rel index -> rId
    const relIds = parseXmlAttr(rvRelXml, /r:id="([^"]+)"/g);

    // rId -> target path
    const rIdToTarget = new Map<string, string>();
    const relTargetRegex = /Id="([^"]+)"[^>]*Target="([^"]+)"/g;
    let rtMatch;
    while ((rtMatch = relTargetRegex.exec(rvRelsXml)) !== null) {
      rIdToTarget.set(rtMatch[1], rtMatch[2]);
    }

    cellVmMap.forEach((vm, rowIdx) => {
      const rvIndex = vmToRv.get(vm);
      if (rvIndex === undefined) return;
      const relIndex = rvToRel.get(rvIndex);
      if (relIndex === undefined || relIndex >= relIds.length) return;
      const rId = relIds[relIndex];
      const target = rIdToTarget.get(rId);
      if (!target) return;

      const normalizedPath = ('xl/richData/' + target.replace(/^\.\.\//, '')).replace('richData/media/', 'media/');
      const imgFile = zip.file(normalizedPath);
      if (!imgFile) return;

      const ext = normalizedPath.split('.').pop() || 'png';
      result.set(rowIdx, { buffer: imgFile.asArrayBuffer(), extension: ext });
    });
  } catch (_) { /* silently fail */ }
  return result;
}

function filterExampleRows(rows: any[]): number[] {
  const indices: number[] = [];
  for (let i = 0; i < rows.length; i++) {
    const name = String(rows[i]['nombre'] || '').trim().toUpperCase();
    const doc = String(rows[i]['documento'] || '').trim();
    if (!(name.includes('(EJEMPLO)') || doc === EXAMPLE_DOC)) {
      indices.push(i);
    }
  }
  return indices;
}

@Injectable({ providedIn: 'root' })
export class ExcelService { // v2

  async downloadTemplate(): Promise<void> {
    const { Workbook } = await import('exceljs');
    const wb = new Workbook();
    const ws = wb.addWorksheet('Estudiantes');

    const headerRow = ws.addRow(HEADERS);
    headerRow.eachCell(cell => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF39A900' } } as any;
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border = {
        bottom: { style: 'medium', color: { argb: 'FF2D8600' } }
      };
      cell.protection = { locked: true };
    });

    const exRow = ws.addRow(EXAMPLE_ROW);
    exRow.eachCell(cell => {
      cell.font = { italic: true, color: { argb: 'FF9CA3AF' }, size: 10 };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF8E1' } } as any;
      cell.protection = { locked: false };
    });
    ws.getCell(2, 1).note = 'FILA DE EJEMPLO - puede dejarla o borrarla.\nLa aplicacion la ignora automaticamente al cargar el archivo.';
    ws.getCell(1, HEADERS.indexOf('firma') + 1).note = 'FIRMA - Inserte una imagen de la firma del estudiante en esta celda.\nEste campo es opcional y se usara para el formato de Seguimiento.';

    HEADERS.forEach((h, i) => {
      ws.getColumn(i + 1).width = COL_WIDTHS[h] || 14;
    });

    const docTypes = ['"TI,CC,CE,PPT"'];
    const tipoCol = HEADERS.indexOf('tipo') + 1;
    const tipoPadreCol = HEADERS.indexOf('tipopadre') + 1;

    for (let r = 2; r <= 502; r++) {
      ws.getCell(r, tipoCol).dataValidation = {
        type: 'list', allowBlank: true, formulae: docTypes,
        showErrorMessage: true, errorTitle: 'Tipo invalido',
        error: 'Seleccione: TI, CC, CE o PPT'
      };
      ws.getCell(r, tipoPadreCol).dataValidation = {
        type: 'list', allowBlank: true, formulae: docTypes,
        showErrorMessage: true, errorTitle: 'Tipo invalido',
        error: 'Seleccione: TI, CC, CE o PPT'
      };
      if (r > 2) {
        ws.getCell(r, tipoCol).protection = { locked: false };
        ws.getCell(r, tipoPadreCol).protection = { locked: false };
      }
    }

    for (let r = 3; r <= 502; r++) {
      for (let c = 1; c <= HEADERS.length; c++) {
        ws.getCell(r, c).protection = { locked: false };
      }
    }

    await ws.protect('', {
      selectLockedCells: true, selectUnlockedCells: true,
      formatCells: true, formatColumns: true, formatRows: true,
      insertRows: true, deleteRows: true, sort: true, autoFilter: true
    });

    const seg = wb.addWorksheet('SEG');
    const segFields = [
      'Momento N° 1 - Planeación de la Etapa Productiva Media Técnica (Se realiza por una única vez)',
      'Fecha inicio Etapa Productiva: (DD/MM/AA)',
      'Fecha Fin de Etapa Productiva: (DD/MM/AA)',
      'Fecha de ingreso ARL:',
      'Ciudad',
      'fecha planeación de actividades',
      ' forma',
      'Momento N° 2 - Seguimiento Etapa Productiva Media Técnica',
      'Fecha inicio Etapa Productiva:',
      'Fecha del momento de seguimiento:',
      'Ciudad',
      'fecha de diligenciamiento',
      ' forma',
      'Momento N° 3 - Evaluación etapa productiva media técnica',
      'Fecha inicio etapa productiva: (DD/MM/AA)',
      'Fecha de fin de la ejecución de la etapa productiva:',
      'Ciudad',
      'fecha de diligenciamiento',
      ' forma',
    ];
    segFields.forEach((label, i) => {
      const row = seg.addRow([label]);
      if (label.startsWith('Momento')) {
        row.getCell(1).font = { bold: true, size: 11 };
        row.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F5E0' } } as any;
        seg.mergeCells(i + 1, 1, i + 1, 9);
      } else {
        row.getCell(1).font = { size: 10 };
      }
      row.getCell(2).protection = { locked: false };
    });
    seg.getColumn(1).width = 45;
    seg.getColumn(2).width = 25;
    seg.getCell('A1').note = 'Llene la columna B con los datos correspondientes.\nEstos campos se usaran para completar el formato de Seguimiento.\nSi deja un campo vacio, no se reemplazara en el formato.';

    const buf = await wb.xlsx.writeBuffer();
    saveAs(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), 'plantilla_estudiantes.xlsx');
  }

  private static readonly REQUIRED_COLS = ['nombre', 'documento'];
  private static readonly KNOWN_COLS = new Set(HEADERS);

  parseExcel(file: File): Promise<Estudiante[]> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target?.result as ArrayBuffer);
          const workbook = XLSX.read(data, { type: 'array' });
          const sheet = workbook.Sheets[workbook.SheetNames[0]];
          const rows: any[] = XLSX.utils.sheet_to_json(sheet, { defval: '' });

          if (!rows.length) {
            reject(new ExcelValidationError('El archivo Excel esta vacio. No se encontraron filas de datos.'));
            return;
          }

          const fileHeaders = Object.keys(rows[0]).map(h => h.trim().toLowerCase());

          const missingRequired = ExcelService.REQUIRED_COLS.filter(
            col => !fileHeaders.includes(col)
          );
          if (missingRequired.length) {
            reject(new ExcelValidationError(
              `El archivo no tiene las columnas obligatorias: ${missingRequired.join(', ')}.\n\n` +
              `Columnas encontradas: ${fileHeaders.join(', ')}.\n\n` +
              `Descargue la plantilla en blanco para ver el formato correcto.`
            ));
            return;
          }

          const matchedCols = fileHeaders.filter(h => ExcelService.KNOWN_COLS.has(h) || h === 'mes ');
          if (matchedCols.length < 3) {
            reject(new ExcelValidationError(
              `El archivo no parece tener el formato correcto. Solo se reconocieron ${matchedCols.length} columnas: ${matchedCols.join(', ')}.\n\n` +
              `Se esperan columnas como: ${HEADERS.slice(0, 8).join(', ')}, etc.\n\n` +
              `Descargue la plantilla en blanco para ver el formato correcto.`
            ));
            return;
          }

          const filtered = rows.filter(row => {
            const name = String(row['nombre'] || '').trim().toUpperCase();
            const doc = String(row['documento'] || '').trim();
            return !(name.includes('(EJEMPLO)') || doc === EXAMPLE_DOC);
          });

          const estudiantes: Estudiante[] = filtered.map(row => ({
            grupo: String(row['grupo'] || '').trim(),
            nombre: String(row['nombre'] || '').trim(),
            documento: String(row['documento'] || '').trim(),
            tipo: String(row['tipo'] || '').trim(),
            otro: String(row['otro'] || '').trim(),
            acudiente: String(row['acudiente'] || '').trim(),
            tipopadre: String(row['tipopadre'] || '').trim(),
            docacudiente: String(row['docacudiente'] || '').trim(),
            lugardoc: String(row['lugardoc'] || '').trim(),
            correo: String(row['correo'] || '').trim(),
            telefono1: String(row['telefono1'] || '').trim(),
            telefono2: String(row['telefono2'] || '').trim(),
            direccion: String(row['direccion'] || '').trim(),
            ficha: String(row['ficha'] || '').trim(),
            dia: String(row['dia'] || '').trim(),
            mes: String(row['mes'] || row['mes '] || '').trim(),
            programa: String(row['programa'] || '').trim(),
            ciudadmunicipio: String(row['ciudadmunicipio'] || '').trim(),
            regional: String(row['regional'] || '').trim(),
            selected: false,
          }));

          const valid = estudiantes.filter(e => e.nombre || e.documento);
          if (!valid.length) {
            reject(new ExcelValidationError(
              'El archivo tiene las columnas correctas pero todas las filas estan vacias.\n\n' +
              'Verifique que el archivo contenga datos de estudiantes.'
            ));
            return;
          }

          const firmaCol = HEADERS.indexOf('firma');
          if (firmaCol >= 0) {
            try {
              const images = extractCellImages(data, firmaCol);
              const exampleRows = filterExampleRows(rows);
              for (let i = 0; i < valid.length && i < exampleRows.length; i++) {
                const rowIdx = exampleRows[i];
                const img = images.get(rowIdx);
                if (img) {
                  valid[i].firmaBuffer = img.buffer;
                  valid[i].firmaExtension = img.extension;
                }
              }
            } catch (_) { /* firma extraction is optional */ }
          }

          resolve(valid);
        } catch (err) {
          reject(err);
        }
      };
      reader.onerror = () => reject(reader.error);
      reader.readAsArrayBuffer(file);
    });
  }

  parseSeguimientoData(file: File): Promise<SeguimientoData | null> {
    return new Promise((resolve) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target?.result as ArrayBuffer);
          const wb = XLSX.read(data, { type: 'array' });
          const seg = wb.Sheets['SEG'];
          if (!seg) { resolve(null); return; }

          const cell = (ref: string) => {
            const c = seg[ref];
            if (!c) return '';
            if (c.w) return String(c.w).trim();
            if (c.t === 'n' && c.v > 10000 && c.v < 3000000) {
              const d = XLSX.SSF.parse_date_code(c.v);
              if (d) return `${String(d.d).padStart(2,'0')}/${String(d.m).padStart(2,'0')}/${d.y}`;
            }
            return String(c.v || '').trim();
          };

          resolve({
            inicioEP: cell('B2'),
            finEP: cell('B3'),
            fechaARL: cell('B4'),
            ciudadPlan: cell('B5'),
            fechaPlan: cell('B6'),
            formaPlan: cell('B7'),
            inicioEPSeg: cell('B9'),
            fechaMomentoSeg: cell('B10'),
            ciudadSeg: cell('B11'),
            fechaDiligSeg: cell('B12'),
            formaSeg: cell('B13'),
            inicioEPEval: cell('B15'),
            finEPEval: cell('B16'),
            ciudadEval: cell('B17'),
            fechaDiligEval: cell('B18'),
            formaEval: cell('B19'),
          });
        } catch (_) { resolve(null); }
      };
      reader.onerror = () => resolve(null);
      reader.readAsArrayBuffer(file);
    });
  }
}

export class ExcelValidationError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'ExcelValidationError';
  }
}
