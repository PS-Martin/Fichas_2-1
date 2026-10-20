import React, { useRef, useEffect } from 'react';
import { VirtualBlock } from '../types';
import {
  FileText,
  PlusCircle,
  RotateCcw,
  MousePointerClick,
  Maximize2,
  FileCheck,
  Bold,
  Italic,
  Underline,
  AlignLeft,
  AlignCenter,
  AlignRight,
  List,
  Save,
  Printer,
  ChevronDown,
  Sparkles,
} from 'lucide-react';

interface WordSimulatorProps {
  documentText: string;
  onDocumentTextChange: (newText: string) => void;
  titleKeyword: string;
  blocks: VirtualBlock[];
  activeBlockIndex: number;
  onBlockSelect: (index: number) => void;
  fileName: string;
  onInsertBlockBetween: () => void;
  onInsertBlockEnd: () => void;
  onResetSampleDocument: () => void;
}

export const WordSimulator: React.FC<WordSimulatorProps> = ({
  documentText,
  onDocumentTextChange,
  titleKeyword,
  blocks,
  activeBlockIndex,
  onBlockSelect,
  fileName,
  onInsertBlockBetween,
  onInsertBlockEnd,
  onResetSampleDocument,
}) => {
  const editorRef = useRef<HTMLTextAreaElement>(null);
  const containerRef = useRef<HTMLDivElement>(null);

  // When activeBlock changes from sidebar, scroll the relevant section into view if in block preview mode
  const activeBlockRef = useRef<HTMLDivElement>(null);
  useEffect(() => {
    if (activeBlockRef.current) {
      activeBlockRef.current.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }
  }, [activeBlockIndex]);

  // Splits document by blocks for visual highlighting in the Word simulator
  const renderDocumentBlocks = () => {
    if (!titleKeyword.trim() || blocks.length === 0) {
      return (
        <div className="prose max-w-none text-slate-800 text-sm whitespace-pre-wrap font-serif leading-relaxed">
          {documentText}
        </div>
      );
    }

    // Split document into segments based on title keyword
    const lines = documentText.split('\n');
    const sections: { title: string; lines: string[]; blockIndex: number }[] = [];
    let currentTitle = '';
    let currentLines: string[] = [];
    let currentIndex = 0;

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      if (line.toLowerCase().includes(titleKeyword.toLowerCase().trim())) {
        if (currentLines.length > 0 || currentTitle) {
          sections.push({
            title: currentTitle || 'Sección Inicial',
            lines: [...currentLines],
            blockIndex: currentIndex,
          });
        }
        currentTitle = line;
        currentLines = [];
        currentIndex++;
      } else {
        currentLines.push(line);
      }
    }

    if (currentTitle || currentLines.length > 0) {
      sections.push({
        title: currentTitle || 'Sección',
        lines: [...currentLines],
        blockIndex: currentIndex,
      });
    }

    return (
      <div className="space-y-4">
        {sections.map((sec, idx) => {
          const isCurrentActive = sec.blockIndex === activeBlockIndex;
          const matchedBlock = blocks.find((b) => b.index === sec.blockIndex);

          return (
            <div
              key={idx}
              ref={isCurrentActive ? activeBlockRef : null}
              onClick={() => {
                if (sec.blockIndex > 0) {
                  onBlockSelect(sec.blockIndex);
                }
              }}
              className={`p-4 rounded-xl transition-all duration-200 cursor-pointer border ${
                isCurrentActive
                  ? 'bg-blue-50/70 border-blue-400 ring-2 ring-blue-400/30 shadow-md'
                  : 'bg-white hover:bg-slate-50 border-slate-200/90 hover:border-slate-300'
              }`}
            >
              <div className="flex items-center justify-between mb-2">
                <div className="flex items-center gap-2">
                  <span
                    className={`text-xs font-bold px-2 py-0.5 rounded-full ${
                      isCurrentActive
                        ? 'bg-[#2b579a] text-white'
                        : 'bg-slate-200 text-slate-700'
                    }`}
                  >
                    Bloque #{sec.blockIndex}
                  </span>
                  <span className="font-bold text-sm text-slate-900 font-sans tracking-tight">
                    {sec.title}
                  </span>
                </div>

                {matchedBlock && (
                  <div className="flex items-center gap-1 text-[11px] font-sans">
                    <span className="text-slate-400 mr-1">Evaluación:</span>
                    <span
                      className={`px-1.5 py-0.5 rounded text-[10px] font-bold ${
                        matchedBlock.ratings.datos === 'green'
                          ? 'bg-emerald-100 text-emerald-800'
                          : matchedBlock.ratings.datos === 'yellow'
                          ? 'bg-amber-100 text-amber-800'
                          : 'bg-rose-100 text-rose-800'
                      }`}
                    >
                      D:{matchedBlock.ratings.datos}
                    </span>
                    <span
                      className={`px-1.5 py-0.5 rounded text-[10px] font-bold ${
                        matchedBlock.ratings.experiencia === 'green'
                          ? 'bg-emerald-100 text-emerald-800'
                          : matchedBlock.ratings.experiencia === 'yellow'
                          ? 'bg-amber-100 text-amber-800'
                          : 'bg-rose-100 text-rose-800'
                      }`}
                    >
                      E:{matchedBlock.ratings.experiencia}
                    </span>
                    <span
                      className={`px-1.5 py-0.5 rounded text-[10px] font-bold ${
                        matchedBlock.ratings.tareas === 'green'
                          ? 'bg-emerald-100 text-emerald-800'
                          : matchedBlock.ratings.tareas === 'yellow'
                          ? 'bg-amber-100 text-amber-800'
                          : 'bg-rose-100 text-rose-800'
                      }`}
                    >
                      T:{matchedBlock.ratings.tareas}
                    </span>
                  </div>
                )}
              </div>

              <div className="font-serif text-slate-700 text-sm leading-relaxed whitespace-pre-line">
                {sec.lines.join('\n')}
              </div>
            </div>
          );
        })}
      </div>
    );
  };

  return (
    <div className="flex-1 flex flex-col bg-[#e6e9ee] h-full overflow-hidden select-none">
      {/* 1. Word 2016 Titlebar */}
      <div className="bg-[#2b579a] text-white px-3 py-1 flex items-center justify-between text-xs border-b border-[#204377] select-none">
        <div className="flex items-center gap-2">
          <div className="bg-white text-[#2b579a] font-black px-1.5 py-0.5 rounded-xs text-[11px] shadow-xs">
            W
          </div>
          <div className="flex items-center gap-1.5 text-slate-200">
            <Save className="w-3.5 h-3.5 opacity-80" />
            <Printer className="w-3.5 h-3.5 opacity-80" />
          </div>
          <span className="font-medium text-white ml-2 tracking-wide truncate max-w-sm">
            {fileName} - Word 2016 (Simulador de Entorno)
          </span>
        </div>
        <div className="flex items-center gap-2 text-slate-200 text-[11px]">
          <span>Complementos activos: 1</span>
        </div>
      </div>

      {/* 2. Word 2016 Ribbon Tabs */}
      <div className="bg-[#f3f4f7] border-b border-slate-300 px-3 flex items-center gap-1 text-xs text-slate-700 pt-1">
        {['Archivo', 'Inicio', 'Insertar', 'Diseño', 'Formato', 'Referencias', 'Revisar', 'Vista'].map((tab, i) => (
          <span
            key={tab}
            className={`px-3 py-1.5 cursor-pointer font-medium ${
              i === 1 ? 'bg-white border-t-2 border-t-[#2b579a] text-[#2b579a] shadow-xs' : 'hover:bg-slate-200/70'
            }`}
          >
            {tab}
          </span>
        ))}
        <span className="px-3 py-1.5 bg-[#2b579a]/10 text-[#2b579a] font-bold border-t-2 border-t-[#2b579a]">
          Complementos (Word Add-in)
        </span>
      </div>

      {/* 3. Word 2016 Toolbar */}
      <div className="bg-white border-b border-slate-200 px-4 py-2 flex items-center justify-between text-xs shadow-xs">
        <div className="flex items-center gap-3">
          <div className="flex items-center gap-1 text-slate-600 bg-slate-100 p-1 rounded-md">
            <button className="p-1 hover:bg-slate-200 rounded" title="Negrita"><Bold className="w-3.5 h-3.5" /></button>
            <button className="p-1 hover:bg-slate-200 rounded" title="Cursiva"><Italic className="w-3.5 h-3.5" /></button>
            <button className="p-1 hover:bg-slate-200 rounded" title="Subrayado"><Underline className="w-3.5 h-3.5" /></button>
            <span className="w-px h-3.5 bg-slate-300 mx-0.5"></span>
            <button className="p-1 hover:bg-slate-200 rounded" title="Alinear a la izquierda"><AlignLeft className="w-3.5 h-3.5" /></button>
            <button className="p-1 hover:bg-slate-200 rounded" title="Centrar"><AlignCenter className="w-3.5 h-3.5" /></button>
            <button className="p-1 hover:bg-slate-200 rounded" title="Lista"><List className="w-3.5 h-3.5" /></button>
          </div>

          <div className="h-4 w-px bg-slate-200"></div>

          {/* Quick test actions */}
          <div className="flex items-center gap-1.5">
            <button
              onClick={onInsertBlockBetween}
              className="px-2.5 py-1 bg-blue-50 text-[#2b579a] hover:bg-blue-100 rounded-md font-semibold text-[11px] flex items-center gap-1 border border-blue-200 transition-colors"
              title="Prueba de inserción intermedia: Inserta un nuevo bloque entre otros dos desplazando las demás secciones para verificar que no se pierden calificaciones"
            >
              <PlusCircle className="w-3.5 h-3.5" />
              <span>+ Insertar Bloque en medio</span>
            </button>

            <button
              onClick={onInsertBlockEnd}
              className="px-2.5 py-1 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-md font-medium text-[11px] flex items-center gap-1 border border-slate-300 transition-colors"
              title="Añadir nuevo bloque al final del documento"
            >
              <PlusCircle className="w-3.5 h-3.5" />
              <span>+ Bloque al final</span>
            </button>

            <button
              onClick={onResetSampleDocument}
              className="px-2 py-1 text-slate-500 hover:text-slate-800 hover:bg-slate-100 rounded-md text-[11px] flex items-center gap-1"
              title="Restablecer documento de ejemplo"
            >
              <RotateCcw className="w-3 h-3" />
              <span>Restablecer</span>
            </button>
          </div>
        </div>

        <div className="text-[11px] text-slate-500 flex items-center gap-2">
          <span>Haz clic en un bloque o edita el texto directamente</span>
        </div>
      </div>

      {/* 4. Document Page Viewport (A4 Sheet Simulation) */}
      <div ref={containerRef} className="flex-1 overflow-y-auto p-6 flex justify-center">
        <div className="w-full max-w-3xl bg-white rounded-lg shadow-xl border border-slate-300 p-8 sm:p-12 min-h-[900px] flex flex-col justify-between">
          <div>
            {/* Document Header */}
            <div className="border-b border-slate-200 pb-4 mb-6 flex justify-between items-start">
              <div>
                <h1 className="text-xl font-bold font-serif text-slate-900">
                  Documento de Evaluación y Gestión
                </h1>
                <p className="text-xs text-slate-500 mt-1">
                  Documento estructurado por bloques virtuales identificados mediante el título:{' '}
                  <span className="font-bold text-[#2b579a]">"{titleKeyword}"</span>
                </p>
              </div>
              <div className="text-right text-[11px] text-slate-400">
                <div>Página 1 de 1</div>
                <div>Formato A4</div>
              </div>
            </div>

            {/* Interactive Document Sections */}
            {renderDocumentBlocks()}
          </div>

          {/* Quick inline text editor drawer */}
          <div className="mt-8 pt-4 border-t border-slate-200">
            <details className="text-xs text-slate-600">
              <summary className="cursor-pointer font-semibold hover:text-[#2b579a] flex items-center gap-1 select-none">
                <FileText className="w-3.5 h-3.5" />
                <span>Editar texto plano del documento directamente</span>
              </summary>
              <div className="mt-2 space-y-2">
                <textarea
                  ref={editorRef}
                  value={documentText}
                  onChange={(e) => onDocumentTextChange(e.target.value)}
                  rows={8}
                  className="w-full p-3 font-mono text-xs bg-slate-50 border border-slate-300 rounded-lg focus:outline-hidden focus:ring-1 focus:ring-[#2b579a]"
                  placeholder="Escribe aquí el contenido del documento..."
                />
                <p className="text-[10px] text-slate-400">
                  Tip: Cada vez que escribas la palabra configurada como título (ej: "{titleKeyword}"), se delimitará un nuevo bloque virtual.
                </p>
              </div>
            </details>
          </div>
        </div>
      </div>

      {/* 5. Word 2016 Statusbar */}
      <div className="bg-[#2b579a] text-white px-4 py-1 flex items-center justify-between text-[11px] select-none border-t border-[#204377]">
        <div className="flex items-center gap-4">
          <span>Página 1 de 1</span>
          <span>{documentText.split(/\s+/).filter(Boolean).length} palabras</span>
          <span>Español (España)</span>
        </div>
        <div className="flex items-center gap-3">
          <span>Bloques detectados: {blocks.length}</span>
          <span className="bg-white/20 px-2 py-0.5 rounded text-[10px]">100% Zoom</span>
        </div>
      </div>
    </div>
  );
};
