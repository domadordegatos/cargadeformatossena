import { Injectable } from '@angular/core';
import { Estudiante, SeguimientoData } from '../models/estudiante.model';
import PizZip from 'pizzip';
import { saveAs } from 'file-saver';

interface GrupoData {
  numero: string;
  estudiantes: Estudiante[];
}

function setCellValue(xml: string, cellRef: string, value: string, styleIdx?: number): string {
  if (!value) return xml;
  const escaped = value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

  const cellRegex = new RegExp(`<c r="${cellRef}"([^/]*?)/>|<c r="${cellRef}"([^>]*?)>[\\s\\S]*?</c>`);
  const match = xml.match(cellRegex);

  let sAttr = '';
  if (styleIdx !== undefined) {
    sAttr = ` s="${styleIdx}"`;
  } else if (match) {
    const sMatch = match[0].match(/ s="(\d+)"/);
    if (sMatch) sAttr = ` s="${sMatch[1]}"`;
  }

  const inlineVal = `<c r="${cellRef}"${sAttr} t="inlineStr"><is><t>${escaped}</t></is></c>`;
  if (match) {
    return xml.replace(match[0], inlineVal);
  }
  return xml;
}



const INFO_CELLS: { row: number; col: string; field: keyof Estudiante | 'docLabel' }[] = [];
for (let i = 0; i < 5; i++) {
  const row = 15 + i;
  INFO_CELLS.push(
    { row, col: 'B', field: 'nombre' },
    { row, col: 'D', field: 'docLabel' as any },
    { row, col: 'F', field: 'telefono1' },
    { row, col: 'H', field: 'direccion' },
    { row, col: 'J', field: 'correo' },
  );
}

const PLAN_NAMES = [
  { cell: 'A45', idx: 0 }, { cell: 'C45', idx: 1 }, { cell: 'E45', idx: 2 },
  { cell: 'G45', idx: 3 }, { cell: 'I45', idx: 4 },
];
const SEG_NAMES = [
  { cell: 'A35', idx: 0 }, { cell: 'D35', idx: 1 }, { cell: 'H35', idx: 2 },
  { cell: 'L35', idx: 3 }, { cell: 'O35', idx: 4 },
];
const EVAL_NAMES = [
  { cell: 'A57', idx: 0 }, { cell: 'D57', idx: 1 }, { cell: 'H57', idx: 2 },
  { cell: 'L57', idx: 3 }, { cell: 'O57', idx: 4 },
];
const EVAL_JUICIO = [
  { cell: 'A50', idx: 0 }, { cell: 'A51', idx: 1 }, { cell: 'A52', idx: 2 },
  { cell: 'A53', idx: 3 }, { cell: 'A54', idx: 4 },
];

function getDocLabel(est: Estudiante): string {
  const tipo = est.tipo || est.otro || '';
  if (!tipo && !est.documento) return '';
  if (!tipo) return est.documento;
  return `${tipo} ${est.documento}`;
}

function getImageMime(ext: string): string {
  if (ext === 'jpeg' || ext === 'jpg') return 'image/jpeg';
  if (ext === 'png') return 'image/png';
  if (ext === 'gif') return 'image/gif';
  return 'image/png';
}

@Injectable({ providedIn: 'root' })
export class SeguimientoService {

  private defaultTemplateBuffer: ArrayBuffer | null = null;
  private customTemplateBuffer: ArrayBuffer | null = null;

  async loadTemplate(): Promise<void> {
    const resp = await fetch('assets/templates/formato_seguimiento.xlsx');
    this.defaultTemplateBuffer = await resp.arrayBuffer();
  }

  setCustomTemplate(buffer: ArrayBuffer | null) {
    this.customTemplateBuffer = buffer;
  }

  get hasCustomTemplate(): boolean {
    return this.customTemplateBuffer !== null;
  }

  private get activeTemplate(): ArrayBuffer {
    return this.customTemplateBuffer || this.defaultTemplateBuffer!;
  }

  private segData: SeguimientoData | null = null;

  setSeguimientoData(data: SeguimientoData | null) {
    this.segData = data;
  }

  async generateForGroups(
    grupos: GrupoData[],
    onProgress?: (current: number, total: number) => void
  ): Promise<void> {
    if (!this.defaultTemplateBuffer) await this.loadTemplate();

    if (grupos.length === 1) {
      const buf = this.fillWorkbook(grupos[0]);
      onProgress?.(1, 1);
      saveAs(
        new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
        `Seguimiento_Grupo_${grupos[0].numero}.xlsx`
      );
      return;
    }

    const zip = new PizZip();
    for (let i = 0; i < grupos.length; i++) {
      const buf = this.fillWorkbook(grupos[i]);
      zip.file(`Seguimiento_Grupo_${grupos[i].numero}.xlsx`, buf);
      onProgress?.(i + 1, grupos.length);
    }

    const content = zip.generate({ type: 'blob', mimeType: 'application/zip' });
    saveAs(content, 'Seguimiento_Grupos.zip');
  }

  private fillWorkbook(grupo: GrupoData): ArrayBuffer {
    const zip = new PizZip(this.activeTemplate);
    const students = grupo.estudiantes.slice(0, 5);
    this.fillInformacion(zip, grupo.numero, students);
    this.fillSegFields(zip);
    this.fillSignatures(zip, 'xl/worksheets/sheet2.xml', PLAN_NAMES, students);
    this.fillSignatures(zip, 'xl/worksheets/sheet3.xml', SEG_NAMES, students);
    this.fillEvaluacion(zip, students);
    this.insertFirmaImages(zip, students);
    this.setFirmaRowHeights(zip);

    return zip.generate({ type: 'arraybuffer' }) as ArrayBuffer;
  }

  private setFirmaRowHeights(zip: any) {
    const targets = [
      { path: 'xl/worksheets/sheet2.xml', row: 44 },
      { path: 'xl/worksheets/sheet3.xml', row: 34 },
      { path: 'xl/worksheets/sheet4.xml', row: 56 },
    ];
    for (const t of targets) {
      const file = zip.file(t.path);
      if (!file) continue;
      let xml = file.asText();
      const rowRegex = new RegExp(`<row r="${t.row}"([^>]*)>`);
      const match = xml.match(rowRegex);
      if (match) {
        let attrs = match[1];
        attrs = attrs.replace(/ ht="[^"]*"/, '').replace(/ customHeight="[^"]*"/, '');
        xml = xml.replace(match[0], `<row r="${t.row}"${attrs} ht="58" customHeight="1">`);
      } else {
        const nextRow = new RegExp(`<row r="${t.row + 1}"`);
        const nm = xml.match(nextRow);
        if (nm) {
          xml = xml.replace(nm[0], `<row r="${t.row}" ht="58" customHeight="1" x14ac:dyDescent="0.25"/>${nm[0]}`);
        }
      }
      zip.file(t.path, xml);
    }
  }

  private fillInformacion(zip: any, grupoNum: string, students: Estudiante[]) {
    const file = zip.file('xl/worksheets/sheet1.xml');
    if (!file) return;
    let xml = file.asText();

    xml = setCellValue(xml, 'C8', students[0]?.regional || '');
    xml = setCellValue(xml, 'C10', students[0]?.programa || '');
    xml = setCellValue(xml, 'K10', grupoNum);

    for (let i = 0; i < students.length; i++) {
      const est = students[i];
      const row = 15 + i;
      xml = setCellValue(xml, `B${row}`, est.nombre);
      xml = setCellValue(xml, `D${row}`, getDocLabel(est));
      xml = setCellValue(xml, `F${row}`, est.telefono1);
      xml = setCellValue(xml, `H${row}`, est.direccion);
      xml = setCellValue(xml, `J${row}`, est.correo);
    }

    zip.file('xl/worksheets/sheet1.xml', xml);
  }

  private fillSegFields(zip: any) {
    const d = this.segData;
    if (!d) return;

    // Hoja 2 - Planeación
    const f2 = zip.file('xl/worksheets/sheet2.xml');
    if (f2) {
      let xml = f2.asText();
      if (d.inicioEP) xml = setCellValue(xml, 'B8', d.inicioEP);
      if (d.finEP) xml = setCellValue(xml, 'F8', d.finEP);
      if (d.fechaARL) xml = setCellValue(xml, 'I8', d.fechaARL);
      if (d.ciudadPlan || d.fechaPlan || d.formaPlan) {
        const ciudad = d.ciudadPlan || '____________';
        const fecha = d.fechaPlan || '___/___/____';
        const forma = d.formaPlan || 'presencial ___ o virtual __';
        xml = setCellValue(xml, 'A49',
          `Ciudad: ${ciudad} y fecha planeacion de actividades: ${fecha} de forma ${forma}`);
      }
      zip.file('xl/worksheets/sheet2.xml', xml);
    }

    // Hoja 3 - Seguimiento
    const f3 = zip.file('xl/worksheets/sheet3.xml');
    if (f3) {
      let xml = f3.asText();
      if (d.inicioEPSeg || d.inicioEP) xml = setCellValue(xml, 'B8', d.inicioEPSeg || d.inicioEP);
      if (d.fechaMomentoSeg) xml = setCellValue(xml, 'O8', d.fechaMomentoSeg);
      if (d.ciudadSeg || d.fechaDiligSeg || d.formaSeg) {
        const ciudad = d.ciudadSeg || d.ciudadPlan || '____________';
        const fecha = d.fechaDiligSeg || '___/___/____';
        const forma = d.formaSeg || 'presencial ___ o virtual ___';
        xml = setCellValue(xml, 'A39',
          `Ciudad: ${ciudad} y fecha de diligenciamiento: ${fecha} de forma ${forma}`);
      }
      zip.file('xl/worksheets/sheet3.xml', xml);
    }

    // Hoja 4 - Evaluación
    const f4 = zip.file('xl/worksheets/sheet4.xml');
    if (f4) {
      let xml = f4.asText();
      if (d.inicioEPEval || d.inicioEP) xml = setCellValue(xml, 'C8', d.inicioEPEval || d.inicioEP);
      if (d.finEPEval || d.finEP) xml = setCellValue(xml, 'I8', d.finEPEval || d.finEP);
      if (d.ciudadEval || d.fechaDiligEval || d.formaEval) {
        const ciudad = d.ciudadEval || d.ciudadPlan || '____________';
        const fecha = d.fechaDiligEval || '___/___/____';
        const forma = d.formaEval || 'presencial ___ o virtual ___';
        xml = setCellValue(xml, 'A61',
          `Ciudad: ${ciudad} y fecha de diligenciamiento: ${fecha} de forma ${forma}`);
      }
      zip.file('xl/worksheets/sheet4.xml', xml);
    }
  }

  private fillSignatures(zip: any, sheetPath: string, nameCells: { cell: string; idx: number }[], students: Estudiante[]) {
    const file = zip.file(sheetPath);
    if (!file) return;
    let xml = file.asText();

    for (const nc of nameCells) {
      if (nc.idx < students.length) {
        xml = setCellValue(xml, nc.cell, students[nc.idx].nombre);
      }
    }

    zip.file(sheetPath, xml);
  }

  private fillEvaluacion(zip: any, students: Estudiante[]) {
    const file = zip.file('xl/worksheets/sheet4.xml');
    if (!file) return;
    let xml = file.asText();

    for (const nc of EVAL_NAMES) {
      if (nc.idx < students.length) {
        xml = setCellValue(xml, nc.cell, students[nc.idx].nombre);
      }
    }

    for (const jc of EVAL_JUICIO) {
      if (jc.idx < students.length) {
        xml = setCellValue(xml, jc.cell, `${jc.idx + 1}. ${students[jc.idx].nombre}`);
      }
    }

    zip.file('xl/worksheets/sheet4.xml', xml);
  }

  private insertFirmaImages(zip: any, students: Estudiante[]) {
    const hasFirmas = students.some(s => s.firmaBuffer);
    if (!hasFirmas) return;

    const sheetConfigs = [
      { sheetPath: 'xl/worksheets/sheet2.xml', relsPath: 'xl/worksheets/_rels/sheet2.xml.rels', drawingPath: 'xl/drawings/drawing2.xml', drawingRelsPath: 'xl/drawings/_rels/drawing2.xml.rels',
        positions: [{ col: 0, toCol: 1, row: 43 }, { col: 2, toCol: 3, row: 43 }, { col: 4, toCol: 5, row: 43 }, { col: 6, toCol: 7, row: 43 }, { col: 8, toCol: 9, row: 43 }] },
      { sheetPath: 'xl/worksheets/sheet3.xml', relsPath: 'xl/worksheets/_rels/sheet3.xml.rels', drawingPath: 'xl/drawings/drawing3.xml', drawingRelsPath: 'xl/drawings/_rels/drawing3.xml.rels',
        positions: [{ col: 0, toCol: 1, row: 33 }, { col: 3, toCol: 5, row: 33 }, { col: 7, toCol: 9, row: 33 }, { col: 11, toCol: 13, row: 33 }, { col: 14, toCol: 16, row: 33 }] },
      { sheetPath: 'xl/worksheets/sheet4.xml', relsPath: 'xl/worksheets/_rels/sheet4.xml.rels', drawingPath: 'xl/drawings/drawing4.xml', drawingRelsPath: 'xl/drawings/_rels/drawing4.xml.rels',
        positions: [{ col: 0, toCol: 1, row: 55 }, { col: 3, toCol: 5, row: 55 }, { col: 7, toCol: 9, row: 55 }, { col: 11, toCol: 13, row: 55 }, { col: 14, toCol: 16, row: 55 }] },
    ];

    let imgCounter = 10;
    const contentTypes = zip.file('[Content_Types].xml')?.asText() || '';
    let ctUpdated = contentTypes;

    for (const cfg of sheetConfigs) {
      const drawingFile = zip.file(cfg.drawingPath);
      if (!drawingFile) continue;
      let drawingXml = drawingFile.asText();

      let drawRelsFile = zip.file(cfg.drawingRelsPath);
      let drawRelsXml = drawRelsFile ? drawRelsFile.asText() : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';

      for (let i = 0; i < students.length; i++) {
        const est = students[i];
        if (!est.firmaBuffer) continue;

        const ext = est.firmaExtension || 'png';
        imgCounter++;
        const mediaPath = `xl/media/firma${imgCounter}.${ext}`;
        zip.file(mediaPath, est.firmaBuffer);

        if (!ctUpdated.includes(`/${ext}"`)) {
          const mime = getImageMime(ext);
          ctUpdated = ctUpdated.replace('</Types>', `<Default Extension="${ext}" ContentType="${mime}"/></Types>`);
        }

        const rId = `rIdFirma${imgCounter}`;
        drawRelsXml = drawRelsXml.replace('</Relationships>',
          `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/firma${imgCounter}.${ext}"/></Relationships>`);

        const pos = cfg.positions[i];
        const anchor = `<xdr:twoCellAnchor>` +
          `<xdr:from><xdr:col>${pos.col}</xdr:col><xdr:colOff>38100</xdr:colOff><xdr:row>${pos.row}</xdr:row><xdr:rowOff>19050</xdr:rowOff></xdr:from>` +
          `<xdr:to><xdr:col>${pos.toCol}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${pos.row + 1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>` +
          `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${100 + imgCounter}" name="Firma${imgCounter}"/><xdr:cNvPicPr><a:picLocks/></xdr:cNvPicPr></xdr:nvPicPr>` +
          `<xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
          `<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>` +
          `<xdr:clientData/></xdr:twoCellAnchor>`;

        drawingXml = drawingXml.replace('</xdr:wsDr>', anchor + '</xdr:wsDr>');
      }

      zip.file(cfg.drawingPath, drawingXml);
      zip.file(cfg.drawingRelsPath, drawRelsXml);
    }

    zip.file('[Content_Types].xml', ctUpdated);
  }
}
