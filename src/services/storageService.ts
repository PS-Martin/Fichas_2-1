import { DocumentConfiguration, VirtualBlock, BlockRatings, calculateBlockSummaryColor } from '../types';

const STORAGE_PREFIX = 'word_addin_blocks_';

/**
 * Normalizes a document file path for storage keying
 */
export function getStorageKeyForPath(filePath: string): string {
  const clean = filePath.trim().toLowerCase().replace(/[\/\\]+/g, '_');
  return `${STORAGE_PREFIX}${clean || 'default_document'}`;
}

/**
 * Saves configuration to browser's LocalStorage under the full file path key
 */
export function saveConfigToLocalStorage(config: DocumentConfiguration): void {
  try {
    const key = getStorageKeyForPath(config.filePath);
    localStorage.setItem(key, JSON.stringify(config));
    // Also save as last active configuration
    localStorage.setItem(`${STORAGE_PREFIX}last_active`, JSON.stringify(config));
  } catch (error) {
    console.error('Error saving configuration to LocalStorage:', error);
  }
}

/**
 * Retrieves configuration from LocalStorage by file path
 */
export function loadConfigFromLocalStorage(filePath: string): DocumentConfiguration | null {
  try {
    const key = getStorageKeyForPath(filePath);
    const item = localStorage.getItem(key);
    if (item) {
      return JSON.parse(item) as DocumentConfiguration;
    }
  } catch (error) {
    console.error('Error reading configuration from LocalStorage:', error);
  }
  return null;
}

/**
 * Exports the document configuration as a downloadable JSON file.
 * Formatted with full document path, block details, timestamp, and ratings.
 */
export function exportConfigAsFile(config: DocumentConfiguration): void {
  const jsonContent = JSON.stringify(config, null, 2);
  const blob = new Blob([jsonContent], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  
  // Extract file base name for friendly download naming
  const parts = config.filePath.split(/[\\/]/);
  const docName = parts[parts.length - 1] || 'documento';
  const cleanName = docName.replace(/\.[^/.]+$/, '');
  const fileName = `${cleanName}_semaforos.json`;

  const link = document.createElement('a');
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

export interface ImportResult {
  success: boolean;
  message: string;
  config?: DocumentConfiguration;
  blockCountMismatch?: {
    expected: number;
    found: number;
  };
}

/**
 * Validates and imports a JSON configuration file into the current document.
 * Must match the exact number of blocks according to requirements:
 * "siempre y cuando detecte el mismo número de bloques"
 */
export function validateAndApplyImport(
  rawJson: string,
  currentBlockCount: number,
  targetFilePath: string
): ImportResult {
  try {
    const parsed = JSON.parse(rawJson);

    // Validate JSON structure
    if (!parsed || !Array.isArray(parsed.blocks)) {
      return {
        success: false,
        message: 'El archivo JSON no tiene un formato válido de configuración de bloques.',
      };
    }

    const importedBlocksCount = parsed.blocks.length;

    // Strict block count validation as specified by user
    if (importedBlocksCount !== currentBlockCount) {
      return {
        success: false,
        message: `Discrepancia en el número de bloques: El archivo contiene ${importedBlocksCount} bloques, pero el documento actual tiene ${currentBlockCount} bloques. Solo se puede asignar a documentos con el mismo número de bloques.`,
        blockCountMismatch: {
          expected: currentBlockCount,
          found: importedBlocksCount,
        },
      };
    }

    // Build verified configuration adapted to current document path
    const validConfig: DocumentConfiguration = {
      filePath: targetFilePath,
      blockTitleKeyword: parsed.blockTitleKeyword || '',
      totalBlocks: importedBlocksCount,
      lastSavedAt: new Date().toISOString(),
      blocks: parsed.blocks.map((b: any, i: number) => ({
        id: b.id || `blk_import_${i + 1}`,
        index: i + 1,
        titleText: b.titleText || `Bloque ${i + 1}`,
        ratings: {
          datos: b.ratings?.datos === 'green' || b.ratings?.datos === 'yellow' ? b.ratings.datos : 'red',
          experiencia: b.ratings?.experiencia === 'green' || b.ratings?.experiencia === 'yellow' ? b.ratings.experiencia : 'red',
          tareas: b.ratings?.tareas === 'green' || b.ratings?.tareas === 'yellow' ? b.ratings.tareas : 'red',
        },
      })),
    };

    return {
      success: true,
      message: `Configuración importada con éxito: ${importedBlocksCount} bloques asignados correctamente.`,
      config: validConfig,
    };
  } catch (err: any) {
    return {
      success: false,
      message: `Error al procesar el archivo: ${err.message || 'JSON inválido'}`,
    };
  }
}

/**
 * Merges new document blocks with existing ratings by persistent ID or matching title,
 * guaranteeing that inserting a new block between existing blocks or adding at the end
 * preserves all existing ratings without altering them.
 */
export function mergeBlocksPreservingRatings(
  scannedBlocks: Omit<VirtualBlock, 'ratings'>[],
  existingBlocks: VirtualBlock[]
): VirtualBlock[] {
  const existingMap = new Map<string, BlockRatings>();
  
  // Index by persistent ID
  for (const block of existingBlocks) {
    existingMap.set(block.id, block.ratings);
  }

  return scannedBlocks.map((scanned, index) => {
    // If we have an existing block with this ID, keep its exact ratings
    const existingRatings = existingMap.get(scanned.id);
    
    // Otherwise fallback to default: all in RED initially as required:
    // "inicialmente para todos los bloques del documento y por defecto los botones están en rojo"
    const ratings: BlockRatings = existingRatings ? { ...existingRatings } : {
      datos: 'red',
      experiencia: 'red',
      tareas: 'red',
    };

    return {
      ...scanned,
      index: index + 1,
      ratings,
    };
  });
}
