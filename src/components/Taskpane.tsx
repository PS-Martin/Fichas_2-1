import React, { useState, useRef } from 'react';
import {
  VirtualBlock,
  DocumentConfiguration,
  TrafficColor,
  calculateBlockSummaryColor,
  RatingCategory,
} from '../types';
import { TrafficLightGroup, SummaryStatusBadge } from './TrafficLightButton';
import {
  exportConfigAsFile,
  validateAndApplyImport,
  saveConfigToLocalStorage,
  ImportResult,
} from '../services/storageService';
import {
  FileText,
  Save,
  Upload,
  CheckCircle2,
  AlertTriangle,
  FolderOpen,
  ChevronLeft,
  ChevronRight,
  Database,
  Briefcase,
  ListTodo,
  Layers,
  Settings2,
  RefreshCw,
  Sliders,
  ExternalLink,
} from 'lucide-react';

interface TaskpaneProps {
  titleKeyword: string;
  onTitleKeywordChange: (val: string) => void;
  onSegmentDocument: () => void;
  blocks: VirtualBlock[];
  activeBlockIndex: number; // 1-based index
  onSelectBlock: (index: number) => void;
  onUpdateRating: (blockId: string, category: RatingCategory, color: TrafficColor) => void;
  filePath: string;
  onFilePathChange: (path: string) => void;
  isWordHost: boolean;
  onOpenManifestGuide?: () => void;
}

export const Taskpane: React.FC<TaskpaneProps> = ({
  titleKeyword,
  onTitleKeywordChange,
  onSegmentDocument,
  blocks,
  activeBlockIndex,
  onSelectBlock,
  onUpdateRating,
  filePath,
  onFilePathChange,
  isWordHost,
  onOpenManifestGuide,
}) => {
  const [autoSaveEnabled, setAutoSaveEnabled] = useState(true);
  const [showConfigSettings, setShowConfigSettings] = useState(false);
  const [notification, setNotification] = useState<{
    type: 'success' | 'error' | 'info';
    message: string;
  } | null>(null);

  const fileInputRef = useRef<HTMLInputElement>(null);

  const activeBlock = blocks.find((b) => b.index === activeBlockIndex) || blocks[0];

  const showToast = (type: 'success' | 'error' | 'info', message: string) => {
    setNotification({ type, message });
    setTimeout(() => {
      setNotification((prev) => (prev?.message === message ? null : prev));
    }, 4500);
  };

  // Requirement: Button can only be clicked if title is provided, otherwise does nothing and displays NO warnings
  const isSegmentButtonActive = titleKeyword.trim().length > 0;

  const handleSegmentClick = () => {
    if (!isSegmentButtonActive) {
      // Do nothing silently without alert as requested:
      // "Solo puede pulsarse (o solo se activa) si el usuario ha escrito un titulo, si no, no hace nada o no puede pulsarse, pero no da mensajes de aviso"
      return;
    }
    onSegmentDocument();
    showToast('info', `Segmentación iniciada con título "${titleKeyword}". Bloques detectados: ${blocks.length}`);
  };

  const handleSaveExport = () => {
    const config: DocumentConfiguration = {
      filePath: filePath.trim() || 'C:\\Documentos\\Documento_Word_2016.docx',
      blockTitleKeyword: titleKeyword,
      totalBlocks: blocks.length,
      lastSavedAt: new Date().toISOString(),
      blocks: blocks.map((b) => ({
        id: b.id,
        index: b.index,
        titleText: b.titleText,
        ratings: b.ratings,
      })),
    };

    // Save to LocalStorage
    saveConfigToLocalStorage(config);
    // Export file to disk
    exportConfigAsFile(config);

    showToast('success', `Configuración guardada para: ${filePath.split(/[\\/]/).pop()}`);
  };

  const handleFileInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      const content = event.target?.result as string;
      if (!content) return;

      const result: ImportResult = validateAndApplyImport(
        content,
        blocks.length,
        filePath
      );

      if (result.success && result.config) {
        // Apply imported ratings to matching block indexes
        result.config.blocks.forEach((importedBlock, idx) => {
          const current = blocks[idx];
          if (current) {
            onUpdateRating(current.id, 'datos', importedBlock.ratings.datos);
            onUpdateRating(current.id, 'experiencia', importedBlock.ratings.experiencia);
            onUpdateRating(current.id, 'tareas', importedBlock.ratings.tareas);
          }
        });
        showToast('success', result.message);
      } else {
        showToast('error', result.message);
      }
    };
    reader.readAsText(file);
    // Reset input
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  return (
    <aside className="w-full h-full bg-slate-50 flex flex-col border-l border-slate-200 select-none text-slate-800 font-sans shadow-lg overflow-hidden">
      {/* Taskpane Header: Word 2016 Office Style */}
      <div className="bg-[#2b579a] text-white px-3.5 py-2.5 flex items-center justify-between shadow-sm">
        <div className="flex items-center gap-2">
          <div className="w-5 h-5 bg-white text-[#2b579a] font-bold rounded-xs flex items-center justify-center text-xs shadow-xs">
            W
          </div>
          <span className="font-semibold text-xs tracking-wide uppercase">Panel Lateral Word</span>
        </div>
        <div className="flex items-center gap-1.5 text-xs">
          {isWordHost ? (
            <span className="bg-emerald-600/90 text-white px-2 py-0.5 rounded-full text-[10px] font-medium flex items-center gap-1">
              <span className="w-1.5 h-1.5 rounded-full bg-white animate-pulse"></span> Word Activo
            </span>
          ) : (
            <span className="bg-blue-800 text-blue-100 px-2 py-0.5 rounded-full text-[10px] font-medium">
              Modo Conectado
            </span>
          )}
          {onOpenManifestGuide && (
            <button
              onClick={onOpenManifestGuide}
              title="Guía de Instalación en Word 2016 y Manifiesto XML"
              className="p-1 hover:bg-white/20 rounded-md text-white transition-colors"
            >
              <ExternalLink className="w-3.5 h-3.5" />
            </button>
          )}
        </div>
      </div>

      {/* Main Content Area */}
      <div className="flex-1 overflow-y-auto p-3.5 space-y-4">
        {/* Toast / Notification Banner */}
        {notification && (
          <div
            className={`p-2.5 rounded-lg text-xs flex items-start gap-2 border animate-in fade-in slide-in-from-top-2 duration-200 ${
              notification.type === 'success'
                ? 'bg-emerald-50 border-emerald-200 text-emerald-800'
                : notification.type === 'error'
                ? 'bg-rose-50 border-rose-200 text-rose-800'
                : 'bg-blue-50 border-blue-200 text-blue-800'
            }`}
          >
            {notification.type === 'success' && <CheckCircle2 className="w-4 h-4 shrink-0 text-emerald-600 mt-0.5" />}
            {notification.type === 'error' && <AlertTriangle className="w-4 h-4 shrink-0 text-rose-600 mt-0.5" />}
            {notification.type === 'info' && <RefreshCw className="w-4 h-4 shrink-0 text-blue-600 mt-0.5" />}
            <span className="flex-1 leading-snug">{notification.message}</span>
          </div>
        )}

        {/* 1. SECCIÓN DE DEFINICIÓN DEL TÍTULO INICIAL Y BOTÓN */}
        <section className="bg-white p-3 rounded-xl border border-slate-200/80 shadow-xs">
          <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-1.5">
            1. Espacio para texto (Título inicial)
          </label>
          <div className="space-y-2">
            <input
              type="text"
              value={titleKeyword}
              onChange={(e) => onTitleKeywordChange(e.target.value)}
              placeholder="Ej: BLOQUE, CANDIDATO, SECCIÓN..."
              className="w-full px-3 py-2 text-xs bg-slate-50 border border-slate-300 rounded-lg focus:bg-white focus:outline-hidden focus:ring-2 focus:ring-[#2b579a] focus:border-transparent transition-all placeholder:text-slate-400 font-medium"
            />

            {/* BOTÓN INICIAL:
                "Solo puede pulsarse (o solo se activa) si el usuario ha escrito un titulo,
                si no, no hace nada o no puede pulsarse, pero no da mensajes de aviso.
                Si el usuario ha escrito un titulo identifica que cada vez que ese titulo
                aparece en el documento es un nuevo bloque" */}
            <button
              type="button"
              disabled={!isSegmentButtonActive}
              onClick={handleSegmentClick}
              className={`w-full py-2 px-3 text-xs font-semibold rounded-lg flex items-center justify-center gap-2 transition-all duration-200 ${
                isSegmentButtonActive
                  ? 'bg-[#2b579a] hover:bg-[#204377] text-white shadow-sm cursor-pointer active:scale-[0.99]'
                  : 'bg-slate-200 text-slate-400 cursor-not-allowed opacity-60'
              }`}
            >
              <Layers className="w-3.5 h-3.5" />
              <span>Identificar Bloques Virtuales</span>
            </button>
          </div>
          <p className="text-[11px] text-slate-500 mt-1.5 leading-tight">
            Segmenta el documento cada vez que aparece este título como separador virtual.
          </p>
        </section>

        {/* 2. CONTROLES DEL BLOQUE ACTIVO EN PANTALLA */}
        {blocks.length > 0 ? (
          <section className="bg-white p-3.5 rounded-xl border border-slate-200/80 shadow-xs space-y-3">
            <div className="flex items-center justify-between pb-2 border-b border-slate-100">
              <div className="flex items-center gap-2">
                <span className="text-xs font-bold text-slate-800 uppercase tracking-wider">
                  Bloque Activo
                </span>
                <span className="bg-[#2b579a]/10 text-[#2b579a] text-[11px] font-bold px-2 py-0.5 rounded-full">
                  #{activeBlock?.index || 1} de {blocks.length}
                </span>
              </div>

              {/* Botones de navegación rápida entre bloques */}
              <div className="flex items-center gap-1">
                <button
                  type="button"
                  disabled={activeBlockIndex <= 1}
                  onClick={() => onSelectBlock(Math.max(1, activeBlockIndex - 1))}
                  className="p-1 rounded-md text-slate-600 hover:bg-slate-100 disabled:opacity-30 disabled:hover:bg-transparent"
                  title="Bloque anterior"
                >
                  <ChevronLeft className="w-4 h-4" />
                </button>
                <button
                  type="button"
                  disabled={activeBlockIndex >= blocks.length}
                  onClick={() => onSelectBlock(Math.min(blocks.length, activeBlockIndex + 1))}
                  className="p-1 rounded-md text-slate-600 hover:bg-slate-100 disabled:opacity-30 disabled:hover:bg-transparent"
                  title="Siguiente bloque"
                >
                  <ChevronRight className="w-4 h-4" />
                </button>
              </div>
            </div>

            {/* Título o extracto del bloque actual */}
            <div className="bg-slate-50 p-2.5 rounded-lg border border-slate-200/70">
              <div className="text-xs font-semibold text-slate-800 truncate">
                {activeBlock?.titleText || `Bloque ${activeBlockIndex}`}
              </div>
              <div className="text-[11px] text-slate-500 line-clamp-1 mt-0.5">
                {activeBlock?.snippet}
              </div>
            </div>

            {/* TRES GRUPOS DE BOTONES CON SUS ETIQUETAS:
                Datos, Experiencia, Tareas
                Cada uno con semáforo excluyente: Verde, Amarillo, Rojo.
                Inicialmente todos en Rojo por defecto. */}
            <div className="space-y-2 pt-1">
              <div className="flex items-center justify-between text-[11px] text-slate-500 px-1 font-medium">
                <span>Criterio de Evaluación</span>
                <span className="flex items-center gap-3">
                  <span className="text-emerald-700 font-semibold">V</span>
                  <span className="text-amber-700 font-semibold">A</span>
                  <span className="text-rose-700 font-semibold">R</span>
                </span>
              </div>

              {/* GRUPO 1: DATOS */}
              <TrafficLightGroup
                label="Datos"
                icon={<Database className="w-3.5 h-3.5" />}
                value={activeBlock?.ratings.datos || 'red'}
                onChange={(color) => activeBlock && onUpdateRating(activeBlock.id, 'datos', color)}
              />

              {/* GRUPO 2: EXPERIENCIA */}
              <TrafficLightGroup
                label="Experiencia"
                icon={<Briefcase className="w-3.5 h-3.5" />}
                value={activeBlock?.ratings.experiencia || 'red'}
                onChange={(color) => activeBlock && onUpdateRating(activeBlock.id, 'experiencia', color)}
              />

              {/* GRUPO 3: TAREAS */}
              <TrafficLightGroup
                label="Tareas"
                icon={<ListTodo className="w-3.5 h-3.5" />}
                value={activeBlock?.ratings.tareas || 'red'}
                onChange={(color) => activeBlock && onUpdateRating(activeBlock.id, 'tareas', color)}
              />
            </div>
          </section>
        ) : (
          <div className="bg-white p-4 rounded-xl border border-dashed border-slate-300 text-center space-y-2">
            <Layers className="w-8 h-8 text-slate-300 mx-auto" />
            <div className="text-xs font-semibold text-slate-700">Ningún bloque identificado todavía</div>
            <p className="text-[11px] text-slate-500">
              Escribe un título en el campo superior y pulsa el botón para segmentar el documento en bloques virtuales.
            </p>
          </div>
        )}

        {/* 3. LISTA DE SECCIONES CON BOTÓN NO CLICABLE SEMÁFORO:
            "Adicionalmente el panel lateral debe verse una lista de numeros identificando cada sección,
            con un único 'botón' no clicable que estará iluminado en verde si todas las etiquetas correspondientes
            a su numero de bloque estan en verde, rojo si hay una sola etiqueta marcada en rojo, y amarillo en cualquier otro caso." */}
        <section className="bg-white p-3.5 rounded-xl border border-slate-200/80 shadow-xs space-y-2.5">
          <div className="flex items-center justify-between pb-1 border-b border-slate-100">
            <span className="text-xs font-bold text-slate-700 uppercase tracking-wider">
              Resumen de Secciones ({blocks.length})
            </span>
            <span className="text-[10px] text-slate-400">Estado global</span>
          </div>

          {blocks.length === 0 ? (
            <div className="text-[11px] text-slate-400 py-2 text-center italic">
              Lista vacía (requiere segmentación)
            </div>
          ) : (
            <div className="max-h-52 overflow-y-auto space-y-1.5 pr-1">
              {blocks.map((block) => {
                const summaryColor = calculateBlockSummaryColor(block.ratings);
                const isSelected = block.index === activeBlockIndex;

                return (
                  <div
                    key={block.id}
                    onClick={() => onSelectBlock(block.index)}
                    className={`flex items-center justify-between p-2 rounded-lg cursor-pointer transition-all border ${
                      isSelected
                        ? 'bg-blue-50/80 border-[#2b579a]/40 shadow-xs'
                        : 'bg-slate-50/70 border-slate-200/80 hover:bg-slate-100'
                    }`}
                  >
                    <div className="flex items-center gap-2.5 min-w-0">
                      {/* Número de sección */}
                      <span
                        className={`w-5 h-5 rounded-md flex items-center justify-center font-bold text-xs ${
                          isSelected ? 'bg-[#2b579a] text-white' : 'bg-slate-200 text-slate-700'
                        }`}
                      >
                        {block.index}
                      </span>
                      <div className="min-w-0">
                        <div className="text-xs font-medium text-slate-800 truncate max-w-[130px]">
                          {block.titleText}
                        </div>
                        <div className="flex items-center gap-1.5 text-[9px] text-slate-400">
                          <span>D:{block.ratings.datos[0].toUpperCase()}</span>
                          <span>E:{block.ratings.experiencia[0].toUpperCase()}</span>
                          <span>T:{block.ratings.tareas[0].toUpperCase()}</span>
                        </div>
                      </div>
                    </div>

                    {/* Único botón NO CLICABLE iluminado según la regla */}
                    <div className="shrink-0 flex items-center gap-1">
                      <SummaryStatusBadge color={summaryColor} size="sm" />
                    </div>
                  </div>
                );
              })}
            </div>
          )}

          {/* Leyenda de la regla */}
          <div className="bg-slate-50 p-2 rounded-lg text-[10px] text-slate-500 space-y-1 border border-slate-100">
            <div className="flex items-center gap-1.5">
              <span className="w-2 h-2 rounded-full bg-emerald-500"></span>
              <span><strong>Verde:</strong> Todas las etiquetas (D, E, T) en verde</span>
            </div>
            <div className="flex items-center gap-1.5">
              <span className="w-2 h-2 rounded-full bg-rose-600"></span>
              <span><strong>Rojo:</strong> Si hay al menos una etiqueta en rojo</span>
            </div>
            <div className="flex items-center gap-1.5">
              <span className="w-2 h-2 rounded-full bg-amber-400"></span>
              <span><strong>Amarillo:</strong> Cualquier otro caso</span>
            </div>
          </div>
        </section>

        {/* 4. GUARDADO, AUTOGUARDADO Y CARGA CON REFERENCIA DE RUTA COMPLETA */}
        <section className="bg-white p-3 rounded-xl border border-slate-200/80 shadow-xs space-y-2.5">
          <div className="flex items-center justify-between">
            <span className="text-xs font-bold text-slate-700 uppercase tracking-wider flex items-center gap-1">
              <FolderOpen className="w-3.5 h-3.5 text-[#2b579a]" /> Almacenamiento & Carga
            </span>
            <button
              type="button"
              onClick={() => setShowConfigSettings(!showConfigSettings)}
              className="text-slate-400 hover:text-slate-600 p-1 rounded-md"
              title="Ajustes de ruta"
            >
              <Settings2 className="w-3.5 h-3.5" />
            </button>
          </div>

          {/* Referencia de ruta completa del archivo */}
          <div className="space-y-1">
            <label className="text-[11px] font-semibold text-slate-600 flex items-center justify-between">
              <span>Ruta completa del documento referenciado:</span>
            </label>
            <input
              type="text"
              value={filePath}
              onChange={(e) => onFilePathChange(e.target.value)}
              placeholder="C:\Ruta\Completa\Documento.docx"
              className="w-full px-2.5 py-1.5 text-[11px] bg-slate-50 border border-slate-300 rounded-md font-mono text-slate-700 focus:outline-hidden focus:ring-1 focus:ring-[#2b579a]"
            />
          </div>

          {/* Botones de Guardar y Cargar */}
          <div className="grid grid-cols-2 gap-2 pt-1">
            {/* Botón de Guardado */}
            <button
              type="button"
              disabled={blocks.length === 0}
              onClick={handleSaveExport}
              className="py-1.5 px-2 bg-emerald-600 hover:bg-emerald-700 disabled:opacity-40 text-white rounded-lg text-xs font-medium flex items-center justify-center gap-1.5 shadow-xs transition-colors cursor-pointer"
              title="Guardar archivo JSON de configuración"
            >
              <Save className="w-3.5 h-3.5" />
              <span>Guardar Archivo</span>
            </button>

            {/* Botón de Carga */}
            <button
              type="button"
              disabled={blocks.length === 0}
              onClick={() => fileInputRef.current?.click()}
              className="py-1.5 px-2 bg-slate-800 hover:bg-slate-900 disabled:opacity-40 text-white rounded-lg text-xs font-medium flex items-center justify-center gap-1.5 shadow-xs transition-colors cursor-pointer"
              title="Cargar y asignar archivo JSON si coincide el número de bloques"
            >
              <Upload className="w-3.5 h-3.5" />
              <span>Cargar Archivo</span>
            </button>

            {/* Input oculto para carga de archivo */}
            <input
              ref={fileInputRef}
              type="file"
              accept=".json"
              onChange={handleFileInputChange}
              className="hidden"
            />
          </div>

          {/* Autoguardado toggle */}
          <div className="flex items-center justify-between pt-1 border-t border-slate-100 text-[11px]">
            <span className="text-slate-600">Autoguardado en documento / local:</span>
            <label className="relative inline-flex items-center cursor-pointer">
              <input
                type="checkbox"
                checked={autoSaveEnabled}
                onChange={(e) => {
                  setAutoSaveEnabled(e.target.checked);
                  showToast('info', e.target.checked ? 'Autoguardado activado' : 'Autoguardado pausado');
                }}
                className="sr-only peer"
              />
              <div className="w-7 h-4 bg-slate-300 peer-focus:outline-hidden rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-slate-300 after:border after:rounded-full after:h-3 after:w-3 after:transition-all peer-checked:bg-[#2b579a]"></div>
            </label>
          </div>
        </section>
      </div>

      {/* Taskpane Footer */}
      <div className="bg-slate-100 border-t border-slate-200 px-3.5 py-2 text-[10px] text-slate-500 flex items-center justify-between">
        <span>Microsoft Word 2016 Add-in</span>
        <span>Office.js v1.1+</span>
      </div>
    </aside>
  );
};
