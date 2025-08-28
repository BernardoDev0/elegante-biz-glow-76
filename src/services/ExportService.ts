import { supabase } from '@/integrations/supabase/client';
import { ExcelFolderService } from './ExcelFolderService';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import { format } from 'date-fns';
import { ptBR } from 'date-fns/locale';

interface ExportEntry {
  Data: string;
  Refinaria: string;
  Pontos: number;
  Observações: string;
}

export class ExportService {
  /**
   * Exporta dados dos funcionários em arquivos Excel separados dentro de um ZIP
   * Agora usa dados dos arquivos Excel em vez do Supabase
   */
  static async exportEmployeeDataToZip(): Promise<void> {
    try {
      // Buscar dados dos arquivos Excel
      const folderData = await ExcelFolderService.processRegistrosFolder();
      const employees = Object.keys(folderData.employees);

      // Criar um ZIP
      const zip = new JSZip();

      // Criar arquivo Excel para cada funcionário
      for (const employeeName of employees) {
        const employeeData = folderData.employees[employeeName];
        
        // Formatar dados para Excel (igual ao formato da imagem)
        let exportData: ExportEntry[];
        
        if (employeeData.records.length > 0) {
          exportData = employeeData.records.map(record => ({
            Data: format(new Date(record.date), 'dd/MM/yyyy HH:mm:ss', { locale: ptBR }),
            Refinaria: record.refinery || '',
            Pontos: record.points || 0,
            Observações: record.observations || ''
          }));
        } else {
          // Se não tem dados, criar arquivo com uma linha indicando
          exportData = [{
            Data: format(new Date(), 'dd/MM/yyyy HH:mm:ss', { locale: ptBR }),
            Refinaria: '',
            Pontos: 0,
            Observações: 'Nenhum registro encontrado'
          }];
        }

        // Criar workbook
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(exportData);

        // Ajustar largura das colunas
        const colWidths = [
          { wch: 20 }, // Data
          { wch: 12 }, // Refinaria  
          { wch: 8 },  // Pontos
          { wch: 50 }  // Observações
        ];
        ws['!cols'] = colWidths;

        // Adicionar ao workbook
        XLSX.utils.book_append_sheet(wb, ws, "Registros");

        // Converter para buffer
        const excelBuffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
        
        // Adicionar ao ZIP - nome igual à segunda imagem
        const currentMonth = format(new Date(), 'MMMM', { locale: ptBR });
        const fileName = `${employeeName} ${currentMonth}.xlsx`;
        zip.file(fileName, excelBuffer);
      }

      // Gerar ZIP e fazer download
      const zipBlob = await zip.generateAsync({ type: 'blob' });
      const url = window.URL.createObjectURL(zipBlob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `registros_funcionarios_${format(new Date(), 'yyyy-MM-dd')}.zip`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      window.URL.revokeObjectURL(url);

    } catch (error) {
      console.error('Erro ao exportar dados:', error);
      throw error;
    }
  }

  /**
   * Busca dados dos funcionários para os gráficos usando arquivos Excel
   */
  static async getEmployeesChartData() {
    try {
      // Usar dados dos arquivos Excel
      return await ExcelFolderService.generateChartData();
      
    } catch (error) {
      console.error('Erro ao buscar dados dos gráficos dos arquivos Excel:', error);
      throw error;
    }
  }

  /**
   * Retorna cor específica para cada funcionário
   */
  static getEmployeeColor(employeeName: string): string {
    const colorMap: Record<string, string> = {
      'Rodrigo': '#8b5cf6',
      'Maurício': '#f59e0b', 
      'Matheus': '#10b981',
      'Wesley': '#ef4444'
    };
    return colorMap[employeeName] || '#6b7280';
  }

  /**
   * Calcula estatísticas gerais dos dados
   * Agora usa dados dos arquivos Excel
   */
  static async getGeneralStats() {
    try {
      // Usar dados dos arquivos Excel
      return await ExcelFolderService.getGeneralStats();
      
    } catch (error) {
      console.error('Erro ao calcular estatísticas dos arquivos Excel:', error);
      return null;
    }
  }
}