export interface SeguimientoData {
  inicioEP: string;
  finEP: string;
  fechaARL: string;
  ciudadPlan: string;
  fechaPlan: string;
  formaPlan: string;
  inicioEPSeg: string;
  fechaMomentoSeg: string;
  ciudadSeg: string;
  fechaDiligSeg: string;
  formaSeg: string;
  inicioEPEval: string;
  finEPEval: string;
  ciudadEval: string;
  fechaDiligEval: string;
  formaEval: string;
}

export interface Estudiante {
  grupo: string;
  nombre: string;
  firmaBuffer?: ArrayBuffer;
  firmaExtension?: string;
  documento: string;
  tipo: string;
  otro: string;
  acudiente: string;
  tipopadre: string;
  docacudiente: string;
  lugardoc: string;
  correo: string;
  telefono1: string;
  telefono2: string;
  direccion: string;
  ficha: string;
  dia: string;
  mes: string;
  programa: string;
  ciudadmunicipio: string;
  regional: string;
  selected?: boolean;
}
