import React, { useState, useEffect, useMemo, useCallback } from 'react';
import {
  VirtualBlock,
  TrafficColor,
  RatingCategory,
  DocumentConfiguration,
} from './types';
import { Taskpane } from './components/Taskpane';
import { WordSimulator } from './components/WordSimulator';
import { ManifestModal } from './components/ManifestModal';
import {
  saveConfigToLocalStorage,
  loadConfigFromLocalStorage,
  mergeBlocksPreservingRatings,
} from './services/storageService';
import { wordService } from './services/wordService';
import {
  Layers,
  Monitor,
  PanelRight,
  FileCode,
  Info,
  CheckCircle,
} from 'lucide-react';

const INITIAL_SAMPLE_TEXT = `DOCUMENTO DE EVALUACIÓN TÉCNICA Y AUDITORÍA DE PROCESOS
Versión 2.4 - Octubre 2026
Confidencialidad: Uso Interno Exclusivo

BLOQUE 1: Carlos Mendoza - Especialista en Infraestructura Cloud
El candidato cuenta con más de 8 años de experiencia en arquitecturas escalables, diseño de microservicios y despliegues en contenedores Kubernetes. Ha liderado migraciones corporativas críticas con cero tiempo de inactividad y certificado en soluciones multinube.
- Formación académica: Grado en Ingeniería Informática.
- Certificaciones vigentes: Cloud Architect Professional, CKA.
- Disponibilidad: Inmediata para incorporación presencial o híbrida.

BLOQUE 2: Lucía Torres - Directora de Producto y Experiencia UX
Responsable de la dirección estratégica de producto digital en empresas de comercio electrónico y fintech. Gran capacidad para coordinar equipos multidisciplinares de diseño, analítica de datos e ingeniería de desarrollo.
- Metodologías: Scrum, Kanban, Lean Startup y Design Sprints.
- Proyectos clave: Lanzamiento de app móvil con más de 2 millones de descargas activas.
- Valoración de liderazgo: Excelente según informes de desempeño previos.

BLOQUE 3: Marcos Silva - Analista de Seguridad y Auditor de Datos
Especialista enfocado en cumplimiento normativo RGPD, auditorías de seguridad perimetral y análisis forense digital. Manejo avanzado de herramientas de monitorización SIEM y análisis de vulnerabilidades automatizado.
- Experiencia laboral: 6 años en auditoría bancaria internacional.
- Tareas habituales: Detección de intrusiones, revisión de código seguro y pentesting.
- Criterios cumplidos: Documentación técnica completa y referencias contrastadas.

BLOQUE 4: Elena Rivas - Gestora de Operaciones y Calidad de Procesos
Supervisa la ejecución operativa, trazabilidad y optimización continua de procesos empresariales. Especializada en automatización mediante flujos robóticos y análisis de KPIs operativos para toma de decisiones ágiles.
- Logros destacados: Reducción del 34% en tiempos de ciclo operacional.
- Idiomas: Español nativo, Inglés C1 certificado.
- Estado de revisión: Documentación pendiente de contrastar por el equipo de recursos humanos.`;

export default function App() {
  // Title / Keyword used to delimit blocks
  const [titleKeyword, setTitleKeyword] = useState<string>('BLOQUE');
  
  // Full path of the referenced document
  const [filePath, setFilePath] = useState<string>(
    'C:\\Documentos\\Evaluacion_Procesos_Word2016.docx'
  );

  // Document text content
  const [documentText, setDocumentText] = useState<string>(INITIAL_SAMPLE_TEXT);

  // Parsed blocks with persistent identifiers
  const [blocks, setBlocks] = useState<VirtualBlock[]>([]);

  // Currently active block in view (1-based index)
  const [activeBlockIndex, setActiveBlockIndex] = useState<number>(1);

  // View mode switcher: side-by-side simulator vs pure taskpane view
  const [viewMode, setViewMode] = useState<'both' | 'taskpane_only'>('both');

  // Manifest modal state
  const [isManifestModalOpen, setIsManifestModalOpen] = useState<boolean>(false);

  // Word environment info
  const [isWordHost, setIsWordHost] = useState<boolean>(false);

  // Initialize Word environment on mount
  useEffect(() => {
    wordService.initialize().then((status) => {
      setIsWordHost(status.isOfficeHost);
      if (status.documentUrl) {
        setFilePath(status.documentUrl);
      }
    });
  }, []);

  /**
   * Scans text to find all virtual block occurrences based on titleKeyword
   */
  const scanDocumentTextIntoBlocks = useCallback(
    (text: string, keyword: string): Omit<VirtualBlock, 'ratings'>[] => {
      if (!keyword.trim()) return [];

      const lines = text.split('\n');
      const results: Omit<VirtualBlock, 'ratings'>[] = [];
      let currentBlockTitle = '';
      let currentSnippet = '';
      let blockCount = 0;

      for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        if (line.toLowerCase().includes(keyword.toLowerCase().trim())) {
          // If we had a previous block pending, push it
          if (blockCount > 0) {
            results.push({
              id: `blk_hash_${blockCount}_${currentBlockTitle.replace(/[^a-zA-Z0-9]/g, '_').slice(0, 20)}`,
              index: blockCount,
              titleText: currentBlockTitle,
              snippet: currentSnippet.trim().slice(0, 140) || 'Contenido del bloque...',
            });
          }
          blockCount++;
          currentBlockTitle = line.trim();
          currentSnippet = '';
        } else if (blockCount > 0) {
          if (!currentSnippet && line.trim()) {
            currentSnippet = line.trim();
          }
        }
      }

      // Push final block
      if (blockCount > 0) {
        results.push({
          id: `blk_hash_${blockCount}_${currentBlockTitle.replace(/[^a-zA-Z0-9]/g, '_').slice(0, 20)}`,
          index: blockCount,
          titleText: currentBlockTitle,
          snippet: currentSnippet.trim().slice(0, 140) || 'Contenido del bloque...',
        });
      }

      return results;
    },
    []
  );

  /**
   * Action triggered when user clicks the initial button:
   * "Si el usuario ha escrito un titulo identifica que cada vez que ese titulo
   * aparece en el documento es un nuevo bloque, como si fueran secciones virtuales"
   */
  const handleSegmentDocument = useCallback(() => {
    if (!titleKeyword.trim()) return;

    const scanned = scanDocumentTextIntoBlocks(documentText, titleKeyword);

    setBlocks((prevBlocks) => {
      // Merge scanned blocks with existing ratings by persistent ID / position
      // Any newly added blocks receive default RED ratings, existing ratings are preserved!
      const merged = mergeBlocksPreservingRatings(scanned, prevBlocks);
      return merged;
    });

    setActiveBlockIndex((prev) => (prev > scanned.length ? 1 : prev));
  }, [titleKeyword, documentText, scanDocumentTextIntoBlocks]);

  // Initial segmentation on first render
  useEffect(() => {
    const scanned = scanDocumentTextIntoBlocks(INITIAL_SAMPLE_TEXT, 'BLOQUE');
    const initialBlocks: VirtualBlock[] = scanned.map((b) => ({
      ...b,
      ratings: {
        datos: 'red',
        experiencia: 'red',
        tareas: 'red',
      },
    }));
    setBlocks(initialBlocks);
  }, [scanDocumentTextIntoBlocks]);

  /**
   * Updates a rating (datos, experiencia, tareas) for a specific block.
   * Ratings are mutually exclusive per label: Green, Yellow, Red.
   */
  const handleUpdateRating = (
    blockId: string,
    category: RatingCategory,
    color: TrafficColor
  ) => {
    setBlocks((prev) =>
      prev.map((b) => {
        if (b.id === blockId) {
          const updatedRatings = {
            ...b.ratings,
            [category]: color,
          };
          return {
            ...b,
            ratings: updatedRatings,
          };
        }
        return b;
      })
    );
  };

  /**
   * Automatically persists settings to document/localStorage when ratings change
   */
  useEffect(() => {
    if (blocks.length === 0) return;

    const config: DocumentConfiguration = {
      filePath,
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

    saveConfigToLocalStorage(config);

    if (isWordHost) {
      wordService.saveToDocumentSettings(config);
    }
  }, [blocks, filePath, titleKeyword, isWordHost]);

  /**
   * Test scenario: Insert a new block between Block 1 and Block 2
   * Verifies requirement: "preparado para la posibilidad de añadir mas bloques sin afectar
   * a las configuraciones realizadas. Por ejemplo si se añade un bloque nuevo entre otros dos, desplazando"
   */
  const handleInsertBlockBetween = () => {
    const lines = documentText.split('\n');
    const insertPosition = lines.findIndex((l) =>
      l.toLowerCase().includes(`${titleKeyword.toLowerCase()} 2`)
    );

    const newBlockText = `\n${titleKeyword} 1.5 (Nuevo en medio): Javier Ortiz - Consultor de Ciberseguridad\nNuevo bloque insertado entre el bloque 1 y el bloque 2 para verificar que los bloques existentes conservan sus calificaciones intactas sin alteración.\n`;

    let newDoc = '';
    if (insertPosition !== -1) {
      lines.splice(insertPosition, 0, newBlockText);
      newDoc = lines.join('\n');
    } else {
      newDoc = documentText + '\n' + newBlockText;
    }

    setDocumentText(newDoc);

    // Re-segment immediately
    const scanned = scanDocumentTextIntoBlocks(newDoc, titleKeyword);
    setBlocks((prev) => mergeBlocksPreservingRatings(scanned, prev));
  };

  /**
   * Test scenario: Append a new block at the end
   */
  const handleInsertBlockEnd = () => {
    const nextNum = blocks.length + 1;
    const appendText = `\n\n${titleKeyword} ${nextNum}: Sofía Castro - Administradora de Base de Datos\nEspecialista en replicación de datos, alta disponibilidad y optimización de consultas SQL complejas para informes de analítica en tiempo real.\n`;
    const newDoc = documentText + appendText;
    setDocumentText(newDoc);

    const scanned = scanDocumentTextIntoBlocks(newDoc, titleKeyword);
    setBlocks((prev) => mergeBlocksPreservingRatings(scanned, prev));
  };

  /**
   * Reset document to original sample state
   */
  const handleResetSampleDocument = () => {
    setDocumentText(INITIAL_SAMPLE_TEXT);
    setTitleKeyword('BLOQUE');
    const scanned = scanDocumentTextIntoBlocks(INITIAL_SAMPLE_TEXT, 'BLOQUE');
    setBlocks(
      scanned.map((b) => ({
        ...b,
        ratings: {
          datos: 'red',
          experiencia: 'red',
          tareas: 'red',
        },
      }))
    );
    setActiveBlockIndex(1);
  };

  return (
    <div className="flex flex-col h-screen w-screen overflow-hidden bg-slate-900 font-sans">
      {/* Top Application Bar */}
      <header className="bg-slate-900 text-white px-4 py-2 flex items-center justify-between border-b border-slate-800 shrink-0">
        <div className="flex items-center gap-3">
          <div className="flex items-center gap-2">
            <span className="w-6 h-6 rounded-md bg-[#2b579a] flex items-center justify-center font-bold text-white text-xs shadow-xs">
              W
            </span>
            <span className="font-bold text-sm tracking-tight text-slate-100">
              Word 2016 Add-in <span className="font-normal text-slate-400">| Segmentador y Semáforo de Bloques</span>
            </span>
          </div>

          <span className="hidden sm:inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[11px] bg-slate-800 text-slate-300 border border-slate-700">
            <CheckCircle className="w-3 h-3 text-emerald-400" />
            WordApi 1.1 / Word 2016 Compliant
          </span>
        </div>

        {/* View mode switcher */}
        <div className="flex items-center gap-2">
          <div className="flex items-center bg-slate-800 p-1 rounded-lg border border-slate-700 text-xs">
            <button
              onClick={() => setViewMode('both')}
              className={`px-3 py-1 rounded-md font-medium flex items-center gap-1.5 transition-all ${
                viewMode === 'both'
                  ? 'bg-[#2b579a] text-white shadow-xs'
                  : 'text-slate-300 hover:text-white hover:bg-slate-700'
              }`}
              title="Ver simulador de Word 2016 y panel lateral juntos"
            >
              <Monitor className="w-3.5 h-3.5" />
              <span className="hidden md:inline">Word + Panel Lateral</span>
            </button>

            <button
              onClick={() => setViewMode('taskpane_only')}
              className={`px-3 py-1 rounded-md font-medium flex items-center gap-1.5 transition-all ${
                viewMode === 'taskpane_only'
                  ? 'bg-[#2b579a] text-white shadow-xs'
                  : 'text-slate-300 hover:text-white hover:bg-slate-700'
              }`}
              title="Ver exclusivamente el panel lateral (como se verá acoplado en Word)"
            >
              <PanelRight className="w-3.5 h-3.5" />
              <span className="hidden md:inline">Solo Panel Lateral</span>
            </button>
          </div>

          {/* Sideloading & Manifest XML button */}
          <button
            onClick={() => setIsManifestModalOpen(true)}
            className="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg font-semibold text-xs flex items-center gap-1.5 shadow-sm transition-all"
            title="Descargar archivo manifest.xml para Word 2016 de escritorio"
          >
            <FileCode className="w-3.5 h-3.5" />
            <span>Instalar en Word 2016 (XML)</span>
          </button>
        </div>
      </header>

      {/* Main Workspace */}
      <main className="flex-1 flex overflow-hidden">
        {/* Word 2016 Document Simulator (visible in 'both' view mode) */}
        {viewMode === 'both' && (
          <div className="flex-1 min-w-0 h-full overflow-hidden">
            <WordSimulator
              documentText={documentText}
              onDocumentTextChange={(val) => {
                setDocumentText(val);
                // Dynamically update blocks
                const scanned = scanDocumentTextIntoBlocks(val, titleKeyword);
                setBlocks((prev) => mergeBlocksPreservingRatings(scanned, prev));
              }}
              titleKeyword={titleKeyword}
              blocks={blocks}
              activeBlockIndex={activeBlockIndex}
              onBlockSelect={(idx) => setActiveBlockIndex(idx)}
              fileName={filePath.split(/[\\/]/).pop() || 'Documento1.docx'}
              onInsertBlockBetween={handleInsertBlockBetween}
              onInsertBlockEnd={handleInsertBlockEnd}
              onResetSampleDocument={handleResetSampleDocument}
            />
          </div>
        )}

        {/* Taskpane (The Web Add-in Lateral Panel) */}
        <div
          className={`${
            viewMode === 'taskpane_only'
              ? 'w-full max-w-md mx-auto my-0 h-full border-x border-slate-700 shadow-2xl'
              : 'w-[360px] lg:w-[380px] shrink-0 h-full'
          }`}
        >
          <Taskpane
            titleKeyword={titleKeyword}
            onTitleKeywordChange={setTitleKeyword}
            onSegmentDocument={handleSegmentDocument}
            blocks={blocks}
            activeBlockIndex={activeBlockIndex}
            onSelectBlock={(idx) => setActiveBlockIndex(idx)}
            onUpdateRating={handleUpdateRating}
            filePath={filePath}
            onFilePathChange={setFilePath}
            isWordHost={isWordHost}
            onOpenManifestGuide={() => setIsManifestModalOpen(true)}
          />
        </div>
      </main>

      {/* Manifest & Installation Modal */}
      <ManifestModal
        isOpen={isManifestModalOpen}
        onClose={() => setIsManifestModalOpen(false)}
        appUrl={window.location.origin}
      />
    </div>
  );
}
