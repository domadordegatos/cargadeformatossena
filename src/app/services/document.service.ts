import { Injectable } from '@angular/core';
import { Estudiante } from '../models/estudiante.model';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import { saveAs } from 'file-saver';

@Injectable({ providedIn: 'root' })
export class DocumentService {

  private compromisoTemplate: ArrayBuffer | null = null;
  private tratamientoTemplate: ArrayBuffer | null = null;

  async loadTemplates(): Promise<void> {
    const [compromiso, tratamiento] = await Promise.all([
      fetch('assets/templates/compromiso.docx').then(r => r.arrayBuffer()),
      fetch('assets/templates/tratamiento.docx').then(r => r.arrayBuffer()),
    ]);
    this.compromisoTemplate = compromiso;
    this.tratamientoTemplate = tratamiento;
  }

  private pad2(val: string): string {
    const n = val.replace(/\D/g, '');
    if (!n) return '';
    return n.padStart(2, '0');
  }

  private upper(val: string): string {
    return val.toUpperCase();
  }

  private resolveDocType(tipo: string, otro: string) {
    const t = tipo.toUpperCase().trim();
    return {
      o1: t === 'TI' ? 'X' : '',
      o2: t === 'CC' ? 'X' : '',
      o3: t === 'CE' ? 'X' : '',
      o4: (!t && otro) ? 'X' : '',
      o5: (!t && otro) ? otro : '',
    };
  }

  private resolvePadreDocType(tipopadre: string) {
    const t = tipopadre.toUpperCase().trim();
    return {
      padreCC: t === 'CC' ? 'X' : '',
      padreCE: t === 'CE' ? 'X' : '',
    };
  }

  private buildCompromisoData(est: Estudiante) {
    const docType = this.resolveDocType(est.tipo, est.otro);
    return {
      nombre: this.upper(est.nombre),
      documento: est.documento,
      o1: docType.o1,
      o2: docType.o2,
      o3: docType.o3,
      o4: docType.o4,
      o5: docType.o5,
      programa: ' ' + this.upper(est.programa),
      ficha: est.ficha,
      tipopadre: est.tipopadre || 'PPT',
      docacudiente: est.docacudiente,
      dia: est.dia ? ' ' + this.pad2(est.dia) : '',
      mes: est.mes ? ' ' + this.pad2(est.mes) : '',
    };
  }

  private buildTratamientoData(est: Estudiante) {
    const padreType = this.resolvePadreDocType(est.tipopadre);
    const tipoEstudiante = est.tipo || est.otro || 'PPT';
    return {
      dia: est.dia ? this.pad2(est.dia) : '',
      mes: est.mes ? this.pad2(est.mes) : '',
      ciudadmunicipio: this.upper(est.ciudadmunicipio),
      regional: this.upper(est.regional),
      programa: this.upper(est.programa),
      ficha: est.ficha,
      acudiente: this.upper(est.acudiente),
      o1: padreType.padreCC,
      o2: padreType.padreCE,
      docacudiente: est.docacudiente,
      lugardoc: est.lugardoc,
      nombre: this.upper(est.nombre),
      documento: est.documento,
      tipo: tipoEstudiante,
      tipopadre: est.tipopadre || 'PPT',
      correo: est.correo,
      direccion: est.direccion,
      telefono1: est.telefono1,
    };
  }

  private fillTemplate(template: ArrayBuffer, data: Record<string, string>): Blob {
    const zip = new PizZip(template);
    const doc = new Docxtemplater(zip, {
      delimiters: { start: '[', end: ']' },
      paragraphLoop: true,
      linebreaks: true,
    });
    doc.render(data);
    return doc.getZip().generate({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
  }

  generateCompromiso(est: Estudiante): Blob {
    if (!this.compromisoTemplate) throw new Error('Templates not loaded');
    return this.fillTemplate(this.compromisoTemplate, this.buildCompromisoData(est));
  }

  generateTratamiento(est: Estudiante): Blob {
    if (!this.tratamientoTemplate) throw new Error('Templates not loaded');
    return this.fillTemplate(this.tratamientoTemplate, this.buildTratamientoData(est));
  }

  downloadCompromiso(est: Estudiante): void {
    const blob = this.generateCompromiso(est);
    saveAs(blob, `compromiso_${est.nombre.replace(/\s+/g, '_')}.docx`);
  }

  downloadTratamiento(est: Estudiante): void {
    const blob = this.generateTratamiento(est);
    saveAs(blob, `tratamiento_${est.nombre.replace(/\s+/g, '_')}.docx`);
  }

  async downloadAllForStudent(est: Estudiante): Promise<void> {
    this.downloadCompromiso(est);
    await new Promise(r => setTimeout(r, 300));
    this.downloadTratamiento(est);
  }

  async downloadAllSelected(
    estudiantes: Estudiante[],
    onProgress?: (current: number, total: number) => void
  ): Promise<void> {
    if (estudiantes.length === 1) {
      await this.downloadAllForStudent(estudiantes[0]);
      return;
    }

    const zipAll = new PizZip();
    const total = estudiantes.length;

    for (let i = 0; i < total; i++) {
      const est = estudiantes[i];
      const compBlob = this.generateCompromiso(est);
      const tratBlob = this.generateTratamiento(est);
      const safeName = est.nombre.toUpperCase().replace(/\s+/g, '_');
      zipAll.file(`compromiso_${safeName}.docx`, await compBlob.arrayBuffer());
      zipAll.file(`tratamiento_${safeName}.docx`, await tratBlob.arrayBuffer());

      if (onProgress) onProgress(i + 1, total);
      // yield to UI thread every 5 students
      if (i % 5 === 4) await new Promise(r => setTimeout(r, 0));
    }

    const fichas = [...new Set(estudiantes.map(e => e.ficha))];
    const fichaLabel = fichas.length === 1 ? fichas[0] : 'varias_fichas';
    const zipName = `ficha_${fichaLabel}_${total}_estudiantes.zip`;

    const zipBlob = zipAll.generate({ type: 'blob', mimeType: 'application/zip' });
    saveAs(zipBlob, zipName);
  }
}
