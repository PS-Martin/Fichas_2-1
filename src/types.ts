export type TrafficColor = 'green' | 'yellow' | 'red';

export type RatingCategory = 'datos' | 'experiencia' | 'tareas';

export interface BlockRatings {
  datos: TrafficColor;
  experiencia: TrafficColor;
  tareas: TrafficColor;
}

export interface VirtualBlock {
  id: string; // Unique persistent identifier (e.g. blk_abc123)
  index: number; // 1-based sequential position in document
  titleText: string; // The heading or title found
  snippet: string; // Brief preview text of the block contents
  startOffset?: number; // Text offset or paragraph index in document
  ratings: BlockRatings;
}

export interface DocumentConfiguration {
  filePath: string; // Full path including file name, e.g. "C:\Documentos\Evaluacion_Candidatos_2026.docx"
  blockTitleKeyword: string; // The keyword or title used to segment (e.g. "BLOQUE")
  totalBlocks: number;
  lastSavedAt: string; // ISO string
  blocks: {
    id: string;
    index: number;
    titleText: string;
    ratings: BlockRatings;
  }[];
}

export type ViewMode = 'both' | 'taskpane_only' | 'manifest_guide';

/**
 * Calculates the summary color of a block:
 * - Green: if ALL tags (datos, experiencia, tareas) are green
 * - Red: if AT LEAST ONE tag is red
 * - Yellow: in any other case (e.g., green+yellow, yellow+yellow, etc.)
 */
export function calculateBlockSummaryColor(ratings: BlockRatings): TrafficColor {
  if (ratings.datos === 'green' && ratings.experiencia === 'green' && ratings.tareas === 'green') {
    return 'green';
  }
  if (ratings.datos === 'red' || ratings.experiencia === 'red' || ratings.tareas === 'red') {
    return 'red';
  }
  return 'yellow';
}
