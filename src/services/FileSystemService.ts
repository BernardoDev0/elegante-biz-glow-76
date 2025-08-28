/**
 * Serviço para interagir com o sistema de arquivos local
 * Responsável por ler a pasta "registros monitorar" e seus arquivos Excel
 */
export class FileSystemService {
  
  /**
   * Lista todos os arquivos Excel na pasta "registros monitorar"
   */
  static async listExcelFiles(): Promise<Array<{ name: string; path: string; folder: string }>> {
    const files: Array<{ name: string; path: string; folder: string }> = [];
    
    try {
      // Listar subpastas conhecidas
      const folders = ['mes 4', 'mes 5', 'mes 6', 'mes 7'];
      
      for (const folder of folders) {
        try {
          // Tentar listar arquivos da pasta
          const folderFiles = await this.listFilesInFolder(`registros monitorar/${folder}`);
          
          folderFiles.forEach(file => {
            if (file.name.match(/\.(xlsx|xls)$/i)) {
              files.push({
                name: file.name,
                path: file.path,
                folder: folder
              });
            }
          });
          
        } catch (error) {
          console.warn(`Pasta ${folder} não encontrada ou inacessível:`, error);
        }
      }
      
      return files;
      
    } catch (error) {
      console.error('Erro ao listar arquivos Excel:', error);
      return [];
    }
  }

  /**
   * Lista arquivos em uma pasta específica
   */
  private static async listFilesInFolder(folderPath: string): Promise<Array<{ name: string; path: string }>> {
    // Esta implementação depende do ambiente
    // No WebContainer, você pode usar APIs específicas ou fetch
    
    // Por enquanto, retornar lista conhecida baseada na estrutura do projeto
    const knownFiles: Record<string, string[]> = {
      'registros monitorar/mes 4': [
        'Matheus Abril.xlsx',
        'Maurício Abril.xlsx', 
        'Rodrigo Abril.xlsx'
      ],
      'registros monitorar/mes 5': [
        'Matheus Maio.xlsx',
        'Maurício Maio.xlsx',
        'Wesley Maio.xlsx'
      ],
      'registros monitorar/mes 6': [
        'Matheus Junho.xlsx',
        'Maurício Junho.xlsx',
        'Wesley Junho.xlsx'
      ],
      'registros monitorar/mes 7': [
        'Matheus Julho.xlsx',
        'Maurício Julho.xlsx',
        'Wesley Julho.xlsx'
      ]
    };

    const files = knownFiles[folderPath] || [];
    return files.map(fileName => ({
      name: fileName,
      path: `${folderPath}/${fileName}`
    }));
  }

  /**
   * Lê conteúdo de um arquivo Excel
   */
  static async readExcelFile(filePath: string): Promise<ArrayBuffer> {
    try {
      const response = await fetch(`/${filePath}`);
      if (!response.ok) {
        throw new Error(`Arquivo não encontrado: ${filePath}`);
      }
      return await response.arrayBuffer();
    } catch (error) {
      console.error(`Erro ao ler arquivo ${filePath}:`, error);
      throw error;
    }
  }

  /**
   * Verifica se um arquivo existe
   */
  static async fileExists(filePath: string): Promise<boolean> {
    try {
      const response = await fetch(`/${filePath}`, { method: 'HEAD' });
      return response.ok;
    } catch {
      return false;
    }
  }

  /**
   * Obtém informações de um arquivo
   */
  static async getFileInfo(filePath: string): Promise<{ size: number; lastModified: Date } | null> {
    try {
      const response = await fetch(`/${filePath}`, { method: 'HEAD' });
      if (!response.ok) return null;
      
      return {
        size: parseInt(response.headers.get('content-length') || '0'),
        lastModified: new Date(response.headers.get('last-modified') || Date.now())
      };
    } catch {
      return null;
    }
  }
}