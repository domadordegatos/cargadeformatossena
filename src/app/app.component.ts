import { Component, OnInit } from '@angular/core';
import { driver } from 'driver.js';
import { Estudiante } from './models/estudiante.model';
import { ExcelService, ExcelValidationError } from './services/excel.service';
import { DocumentService } from './services/document.service';
import { SeguimientoService } from './services/seguimiento.service';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent implements OnInit {
  activeMode: 'select' | 'compromiso' | 'seguimiento' = 'select';
  estudiantes: Estudiante[] = [];
  loading = false;
  templatesLoaded = false;
  generating = false;
  progressCurrent = 0;
  progressTotal = 0;
  fileName = '';
  searchTerm = '';
  allSelected = false;
  editMode = false;
  showPrivacyModal = false;
  privacyAccepted = false;
  errorMessage = '';
  private readonly TOUR_KEY = 'sena_tour_last';
  private readonly TOUR_INTERVAL = 3600000;

  generalDia = '';
  generalMes = '';
  generalAnio = '2026';
  generalFicha = '';
  generalPrograma = '';
  generalRegional = '';
  generalCiudad = '';
  sinFecha = false;
  sinFicha = false;
  fichaOriginalUnica = true;

  private originalDia = '';
  private originalMes = '';
  private originalAnio = '2026';
  private originalFicha = '';

  docTypes = ['TI', 'CC', 'CE', 'PPT'];
  grupos: { numero: string; estudiantes: Estudiante[]; selected: boolean }[] = [];
  gruposSinAsignar: Estudiante[] = [];
  customTemplateName = '';

  constructor(
    private excelService: ExcelService,
    private documentService: DocumentService,
    private seguimientoService: SeguimientoService
  ) {}

  async ngOnInit() {
    await this.documentService.loadTemplates();
    this.templatesLoaded = true;
  }

  get filteredEstudiantes(): Estudiante[] {
    if (!this.searchTerm.trim()) return this.estudiantes;
    const term = this.searchTerm.toLowerCase();
    return this.estudiantes.filter(e =>
      e.nombre.toLowerCase().includes(term) ||
      e.documento.includes(term) ||
      e.ficha.includes(term)
    );
  }

  get selectedCount(): number {
    return this.estudiantes.filter(e => e.selected).length;
  }

  get fichaUnica(): boolean {
    if (!this.estudiantes.length) return true;
    const first = this.estudiantes[0].ficha;
    return this.estudiantes.every(e => e.ficha === first);
  }

  get programaUnico(): boolean {
    if (!this.estudiantes.length) return true;
    const first = this.estudiantes[0].programa;
    return this.estudiantes.every(e => e.programa === first);
  }

  get regionalUnica(): boolean {
    if (!this.estudiantes.length) return true;
    const first = this.estudiantes[0].regional;
    return this.estudiantes.every(e => e.regional === first);
  }

  get ciudadUnica(): boolean {
    if (!this.estudiantes.length) return true;
    const first = this.estudiantes[0].ciudadmunicipio;
    return this.estudiantes.every(e => e.ciudadmunicipio === first);
  }

  async onFileChange(event: Event) {
    const input = event.target as HTMLInputElement;
    if (!input.files?.length) return;

    this.loading = true;
    this.errorMessage = '';
    this.fileName = input.files[0].name;
    try {
      this.estudiantes = await this.excelService.parseExcel(input.files[0]);
      this.extractGeneralFields();
      if (this.activeMode === 'seguimiento') {
        this.buildGrupos();
        const segData = await this.excelService.parseSeguimientoData(input.files[0]);
        this.seguimientoService.setSeguimientoData(segData);
      }
      if (this.shouldShowTour()) {
        setTimeout(() => this.startTour(), 500);
      }
    } catch (err: any) {
      if (err instanceof ExcelValidationError) {
        this.errorMessage = err.message;
      } else {
        this.errorMessage = 'Error al leer el archivo. Verifique que sea un archivo Excel valido (.xlsx o .xls).';
      }
      this.fileName = '';
      console.error(err);
    }
    this.loading = false;
    input.value = '';
  }

  buildGrupos() {
    const map = new Map<string, Estudiante[]>();
    this.gruposSinAsignar = [];
    for (const est of this.estudiantes) {
      const g = est.grupo.trim();
      if (!g) {
        this.gruposSinAsignar.push(est);
      } else {
        if (!map.has(g)) map.set(g, []);
        map.get(g)!.push(est);
      }
    }
    this.grupos = Array.from(map.entries())
      .sort((a, b) => Number(a[0]) - Number(b[0]))
      .map(([numero, estudiantes]) => ({ numero, estudiantes, selected: false }));
  }

  extractGeneralFields() {
    if (!this.estudiantes.length) return;
    const first = this.estudiantes[0];
    this.generalDia = first.dia;
    this.generalMes = first.mes;
    this.fichaOriginalUnica = this.fichaUnica;
    if (this.fichaUnica) this.generalFicha = first.ficha;
    if (this.programaUnico) this.generalPrograma = first.programa;
    if (this.regionalUnica) this.generalRegional = first.regional;
    if (this.ciudadUnica) this.generalCiudad = first.ciudadmunicipio;

    this.originalDia = this.generalDia;
    this.originalMes = this.generalMes;
    this.originalAnio = this.generalAnio;
    this.originalFicha = this.generalFicha;

    this.sinFecha = !first.dia && !first.mes;
    this.sinFicha = this.fichaUnica && !first.ficha;
  }

  toggleSinFecha() {
    this.sinFecha = !this.sinFecha;
    if (this.sinFecha) {
      this.generalDia = '';
      this.generalMes = '';
      this.generalAnio = '';
    } else {
      this.generalDia = this.originalDia;
      this.generalMes = this.originalMes;
      this.generalAnio = this.originalAnio;
    }
  }

  toggleSinFicha() {
    this.sinFicha = !this.sinFicha;
    if (this.sinFicha) {
      this.generalFicha = '';
    } else {
      this.generalFicha = this.originalFicha;
    }
  }

  applyGeneralFields() {
    this.estudiantes.forEach(e => {
      if (this.sinFecha) {
        e.dia = '';
        e.mes = '';
      } else {
        e.dia = this.generalDia;
        e.mes = this.generalMes;
      }
      if (this.fichaOriginalUnica) {
        e.ficha = this.sinFicha ? '' : this.generalFicha;
      }
      if (this.programaUnico) e.programa = this.generalPrograma;
      if (this.regionalUnica) e.regional = this.generalRegional;
      if (this.ciudadUnica) e.ciudadmunicipio = this.generalCiudad;
    });
  }

  toggleAll() {
    this.allSelected = !this.allSelected;
    this.filteredEstudiantes.forEach(e => e.selected = this.allSelected);
  }

  toggleEditMode() {
    if (this.editMode) {
      this.applyGeneralFields();
      this.editMode = false;
    } else {
      this.editMode = true;
    }
  }

  toggleEditModeSeg() {
    if (this.editMode) {
      this.buildGrupos();
      this.editMode = false;
    } else {
      this.editMode = true;
    }
  }

  setDocType(est: Estudiante, docType: string) {
    if (docType === 'PPT') {
      est.tipo = '';
      est.otro = 'PPT';
    } else {
      est.tipo = docType;
      est.otro = '';
    }
  }

  getActiveDocType(est: Estudiante): string {
    if (est.tipo) return est.tipo;
    if (est.otro) return est.otro;
    return '';
  }

  setPadreDocType(est: Estudiante, docType: string) {
    est.tipopadre = docType;
  }

  async downloadOne(est: Estudiante, type: 'compromiso' | 'tratamiento' | 'ambos') {
    if (!this.templatesLoaded) return;
    this.generating = true;
    try {
      if (type === 'compromiso') this.documentService.downloadCompromiso(est);
      else if (type === 'tratamiento') this.documentService.downloadTratamiento(est);
      else await this.documentService.downloadAllForStudent(est);
    } catch (err) {
      alert('Error al generar el documento.');
      console.error(err);
    }
    this.generating = false;
  }

  get progressPercent(): number {
    return this.progressTotal > 0 ? Math.round((this.progressCurrent / this.progressTotal) * 100) : 0;
  }

  async downloadSelected() {
    const selected = this.estudiantes.filter(e => e.selected);
    if (!selected.length) return;
    this.generating = true;
    this.progressCurrent = 0;
    this.progressTotal = selected.length;
    try {
      await this.documentService.downloadAllSelected(selected, (current, total) => {
        this.progressCurrent = current;
        this.progressTotal = total;
      });
    } catch (err) {
      alert('Error al generar los documentos.');
      console.error(err);
    }
    this.generating = false;
    this.progressCurrent = 0;
    this.progressTotal = 0;
  }

  getDocTypeLabel(est: Estudiante): string {
    if (est.tipo) return est.tipo;
    if (est.otro) return est.otro;
    return 'N/A';
  }

  private shouldShowTour(): boolean {
    const last = localStorage.getItem(this.TOUR_KEY);
    if (!last) return true;
    return Date.now() - Number(last) > this.TOUR_INTERVAL;
  }

  private markTourSeen() {
    localStorage.setItem(this.TOUR_KEY, String(Date.now()));
  }

  startTour() {
    const tour = driver({
      showProgress: true,
      animate: true,
      overlayColor: 'rgba(0,0,0,.55)',
      stagePadding: 8,
      stageRadius: 8,
      popoverClass: 'sena-tour',
      onDestroyStarted: () => { this.markTourSeen(); tour.destroy(); },
      nextBtnText: 'Siguiente',
      prevBtnText: 'Anterior',
      doneBtnText: 'Entendido',
      progressText: 'Paso {{current}} de {{total}}',
      steps: [
        {
          element: '#tour-toolbar',
          popover: {
            title: 'Barra de herramientas',
            description: 'Aqui encuentra el nombre del archivo cargado, la cantidad de estudiantes, el buscador y los botones principales de la aplicacion.',
            side: 'bottom', align: 'center'
          }
        },
        {
          element: '#tour-general',
          popover: {
            title: 'Campos generales',
            description: 'Estos campos aplican a todos los estudiantes: fecha, ficha, programa, regional y ciudad. Se detectan automaticamente del Excel. Para editarlos, active el modo de edicion.',
            side: 'bottom', align: 'center'
          }
        },
        {
          element: '#tour-select-all',
          popover: {
            title: 'Seleccionar estudiantes',
            description: 'Use esta casilla para seleccionar o deseleccionar todos los estudiantes a la vez. Tambien puede seleccionar estudiantes individualmente con la casilla de cada fila.',
            side: 'right', align: 'start'
          }
        },
        {
          element: '#tour-actions',
          popover: {
            title: 'Generar documentos individuales',
            description: '"C" genera el Compromiso del Aprendiz. "T" genera el Tratamiento de Datos. "2" genera ambos documentos a la vez para este estudiante.',
            side: 'left', align: 'start'
          }
        },
        {
          element: '#tour-edit-btn',
          popover: {
            title: 'Editar datos',
            description: 'Active el modo de edicion para modificar cualquier dato directamente en la tabla: nombres, documentos, tipo de documento, telefonos, etc. Al terminar, presione "Confirmar cambios".',
            side: 'bottom', align: 'center'
          }
        },
        {
          element: '#tour-generate-btn',
          popover: {
            title: 'Generar seleccionados',
            description: 'Seleccione varios estudiantes con las casillas y presione este boton. Si es uno solo, descarga los archivos directamente. Si son varios, genera un archivo ZIP con todos los documentos.',
            side: 'bottom', align: 'center'
          }
        },
        {
          element: '#tour-search',
          popover: {
            title: 'Buscar estudiantes',
            description: 'Filtre estudiantes por nombre, numero de documento o numero de ficha. La busqueda es instantanea y se actualiza mientras escribe.',
            side: 'bottom', align: 'end'
          }
        }
      ]
    });
    tour.drive();
  }

  goHome() {
    if (this.estudiantes.length) {
      this.clearData();
    } else {
      this.activeMode = 'select';
    }
  }

  selectMode(mode: 'compromiso' | 'seguimiento') {
    this.activeMode = mode;
  }

  get selectedGruposCount(): number {
    return this.grupos.filter(g => g.selected).length;
  }

  toggleAllGrupos() {
    const allSel = this.grupos.every(g => g.selected);
    this.grupos.forEach(g => g.selected = !allSel);
  }

  async generateSeguimiento() {
    const selected = this.grupos.filter(g => g.selected);
    if (!selected.length) return;
    this.generating = true;
    this.progressCurrent = 0;
    this.progressTotal = selected.length;
    try {
      await this.seguimientoService.generateForGroups(selected, (current, total) => {
        this.progressCurrent = current;
        this.progressTotal = total;
      });
    } catch (err) {
      alert('Error al generar los formatos de seguimiento.');
      console.error(err);
    }
    this.generating = false;
    this.progressCurrent = 0;
    this.progressTotal = 0;
  }

  onCustomTemplateChange(event: Event) {
    const input = event.target as HTMLInputElement;
    if (!input.files?.length) return;
    const file = input.files[0];
    this.customTemplateName = file.name;
    const reader = new FileReader();
    reader.onload = (e) => {
      this.seguimientoService.setCustomTemplate(e.target?.result as ArrayBuffer);
    };
    reader.readAsArrayBuffer(file);
    input.value = '';
  }

  removeCustomTemplate() {
    this.customTemplateName = '';
    this.seguimientoService.setCustomTemplate(null);
  }

  async downloadTemplate() {
    await this.excelService.downloadTemplate();
  }

  clearData() {
    this.estudiantes = [];
    this.fileName = '';
    this.searchTerm = '';
    this.allSelected = false;
    this.editMode = false;
    this.activeMode = 'select';
    this.customTemplateName = '';
    this.seguimientoService.setCustomTemplate(null);
    this.seguimientoService.setSeguimientoData(null);
  }
}
