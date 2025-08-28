import * as XLSX from 'xlsx';
import { CalculationsService } from './CalculationsService';

export interface ExcelEmployeeData {
  name: string;
  totalPoints: number;
  totalRecords: number;
  monthlyData: { [month: string]: { points: number; records: number } };
  weeklyData: { [week: string]: { points: number; records: number } };
  records: Array<{
    date: string;
    refinery: string;
    points: number;
    observations: string;
    month: string;
    week: number;
  }>;
}

export interface FolderProcessingResult {
  employees: { [name: string]: ExcelEmployeeData };
  statistics: {
    totalFiles: number;
    totalEmployees: number;
    totalRecords: number;
    totalPoints: number;
    totalProfit: number;
  };
  chartData: {
    weeklyData: any[];
    monthlyData: any[];
    teamPerformance: any[];
  };
}

export class ExcelFolderService {
  private static readonly POINT_VALUE = 3.25; // R$ 3,25 por ponto
  
  /**
   * Processa toda a pasta "registros monitorar" e subpastas
   */
  static async processFolderStructure(): Promise<FolderProcessingResult> {
    const result: FolderProcessingResult = {
      employees: {},
      statistics: {
        totalFiles: 0,
        totalEmployees: 0,
        totalRecords: 0,
        totalPoints: 0,
        totalProfit: 0
      },
      chartData: {
        weeklyData: [],
        monthlyData: [],
        teamPerformance: []
      }
    };

    try {
      // Simular estrutura da pasta (você pode adaptar para ler arquivos reais)
      const folderStructure = await this.getFolderStructure();
      
      for (const folder of folderStructure) {
        for (const file of folder.files) {
          try {
            const employeeData = await this.processExcelFile(file);
            this.mergeEmployeeData(result, employeeData);
            result.statistics.totalFiles++;
          } catch (error) {
            console.error(`Erro ao processar ${file.name}:`, error);
          }
        }
      }

      // Calcular estatísticas finais
      this.calculateFinalStatistics(result);
      
      // Gerar dados dos gráficos
      this.generateChartData(result);

      return result;
      
    } catch (error) {
      console.error('Erro ao processar pasta:', error);
      throw error;
    }
  }

  /**
   * Simula a estrutura da pasta (adapte para ler arquivos reais do sistema)
   */
  private static async getFolderStructure() {
    // Esta função deve ser adaptada para ler a estrutura real da pasta
    // Por enquanto, vou simular baseado nos arquivos que vejo no projeto
    
    const folders = [
      {
        name: 'mes 4',
        files: [
          { name: 'Matheus Abril.xlsx', path: 'registros monitorar/mes 4/Matheus Abril.xlsx' },
          { name: 'Maurício Abril.xlsx', path: 'registros monitorar/mes 4/Maurício Abril.xlsx' },
          { name: 'Rodrigo Abril.xlsx', path: 'registros monitorar/mes 4/Rodrigo Abril.xlsx' }
        ]
      },
      {
        name: 'mes 5',
        files: [
          { name: 'Matheus Maio.xlsx', path: 'registros monitorar/mes 5/Matheus Maio.xlsx' },
          { name: 'Maurício Maio.xlsx', path: 'registros monitorar/mes 5/Maurício Maio.xlsx' },
          { name: 'Wesley Maio.xlsx', path: 'registros monitorar/mes 5/Wesley Maio.xlsx' }
        ]
      },
      {
        name: 'mes 6',
        files: [
          { name: 'Matheus Junho.xlsx', path: 'registros monitorar/mes 6/Matheus Junho.xlsx' },
          { name: 'Maurício Junho.xlsx', path: 'registros monitorar/mes 6/Maurício Junho.xlsx' },
          { name: 'Wesley Junho.xlsx', path: 'registros monitorar/mes 6/Wesley Junho.xlsx' }
        ]
      },
      {
        name: 'mes 7',
        files: [
          { name: 'Matheus Julho.xlsx', path: 'registros monitorar/mes 7/Matheus Julho.xlsx' },
          { name: 'Maurício Julho.xlsx', path: 'registros monitorar/mes 7/Maurício Julho.xlsx' },
          { name: 'Wesley Julho.xlsx', path: 'registros monitorar/mes 7/Wesley Julho.xlsx' }
        ]
      }
    ];

    return folders;
  }

  /**
   * Processa um arquivo Excel individual
   */
  private static async processExcelFile(file: { name: string; path: string }): Promise<ExcelEmployeeData> {
    try {
      // Extrair nome do funcionário do arquivo
      const employeeName = this.extractEmployeeNameFromFile(file.name);
      
      // Ler arquivo Excel (você precisa adaptar para ler do sistema de arquivos)
      const workbook = await this.readExcelFile(file.path);
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      const jsonData = XLSX.utils.sheet_to_json(worksheet);

      const employeeData: ExcelEmployeeData = {
        name: employeeName,
        totalPoints: 0,
        totalRecords: 0,
        monthlyData: {},
        weeklyData: {},
        records: []
      };

      // Processar cada linha do Excel
      jsonData.forEach((row: any) => {
        try {
          const record = this.parseExcelRow(row, employeeName);
          if (record && record.points > 0) {
            employeeData.records.push(record);
            employeeData.totalPoints += record.points;
            employeeData.totalRecords++;

            // Agrupar por mês
            if (!employeeData.monthlyData[record.month]) {
              employeeData.monthlyData[record.month] = { points: 0, records: 0 };
            }
            employeeData.monthlyData[record.month].points += record.points;
            employeeData.monthlyData[record.month].records++;

            // Agrupar por semana
            const weekKey = `Semana ${record.week}`;
            if (!employeeData.weeklyData[weekKey]) {
              employeeData.weeklyData[weekKey] = { points: 0, records: 0 };
            }
            employeeData.weeklyData[weekKey].points += record.points;
            employeeData.weeklyData[weekKey].records++;
          }
        } catch (error) {
          console.warn('Erro ao processar linha do Excel:', error);
        }
      });

      return employeeData;
      
    } catch (error) {
      console.error(`Erro ao processar arquivo ${file.name}:`, error);
      throw error;
    }
  }

  /**
   * Lê arquivo Excel do sistema de arquivos
   */
  private static async readExcelFile(filePath: string): Promise<XLSX.WorkBook> {
    try {
      // No ambiente WebContainer, você pode usar fetch para ler arquivos locais
      const response = await fetch(`/${filePath}`);
      if (!response.ok) {
        throw new Error(`Arquivo não encontrado: ${filePath}`);
      }
      
      const arrayBuffer = await response.arrayBuffer();
      return XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
      
    } catch (error) {
      console.error(`Erro ao ler arquivo ${filePath}:`, error);
      throw error;
    }
  }

  /**
   * Extrai nome do funcionário do nome do arquivo
   */
  private static extractEmployeeNameFromFile(fileName: string): string {
    // Remove extensão e mês
    const nameWithoutExt = fileName.replace(/\.(xlsx|xls)$/i, '');
    
    // Extrair apenas o nome (antes do mês)
    const parts = nameWithoutExt.split(' ');
    return parts[0]; // Primeiro nome
  }

  /**
   * Converte linha do Excel para registro estruturado
   */
  private static parseExcelRow(row: any, employeeName: string) {
    try {
      // Tentar diferentes nomes de colunas
      const dateValue = row.Data || row.data || row.DATE || row.Date;
      const pointsValue = parseFloat(row.Pontos || row.pontos || row.PONTOS || row.Points || 0);
      const refineryValue = String(row.Refinaria || row.refinaria || row.REFINARIA || row.Refinery || '').trim();
      const observationsValue = String(row.Observações || row.observacoes || row.OBSERVACOES || row.Observations || '').trim();

      if (!dateValue || pointsValue <= 0) {
        return null;
      }

      const parsedDate = this.parseExcelDate(dateValue);
      const month = this.getMonthFromDate(parsedDate);
      const week = CalculationsService.getWeekFromDate(parsedDate.toISOString().split('T')[0]);

      return {
        date: parsedDate.toISOString(),
        refinery: refineryValue,
        points: pointsValue,
        observations: observationsValue,
        month,
        week
      };
      
    } catch (error) {
      console.warn('Erro ao converter linha do Excel:', error);
      return null;
    }
  }

  /**
   * Converte valor de data do Excel para Date
   */
  private static parseExcelDate(dateValue: any): Date {
    if (dateValue instanceof Date) {
      return dateValue;
    }
    
    if (typeof dateValue === 'string') {
      // Tentar formatos brasileiros: DD/MM/YYYY
      const parts = dateValue.split('/');
      if (parts.length === 3) {
        const day = parseInt(parts[0]);
        const month = parseInt(parts[1]) - 1; // Month is 0-based
        const year = parseInt(parts[2]);
        return new Date(year, month, day);
      }
      
      // Tentar parse direto
      const parsed = new Date(dateValue);
      if (!isNaN(parsed.getTime())) {
        return parsed;
      }
    }
    
    if (typeof dateValue === 'number') {
      // Excel serial date
      return new Date((dateValue - 25569) * 86400 * 1000);
    }
    
    return new Date();
  }

  /**
   * Determina mês baseado na lógica 26→25
   */
  private static getMonthFromDate(date: Date): string {
    const day = date.getDate();
    const month = date.getMonth() + 1;
    const year = date.getFullYear();

    const monthNames = [
      'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
      'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'
    ];

    let targetMonth = month;
    let targetYear = year;

    // Lógica 26→25: se dia >= 26, pertence ao próximo mês
    if (day >= 26) {
      targetMonth = month + 1;
      if (targetMonth > 12) {
        targetMonth = 1;
        targetYear = year + 1;
      }
    }

    return monthNames[targetMonth - 1];
  }

  /**
   * Mescla dados de um funcionário no resultado geral
   */
  private static mergeEmployeeData(result: FolderProcessingResult, employeeData: ExcelEmployeeData) {
    const employeeName = employeeData.name;
    
    if (!result.employees[employeeName]) {
      result.employees[employeeName] = employeeData;
    } else {
      // Mesclar dados existentes
      const existing = result.employees[employeeName];
      existing.totalPoints += employeeData.totalPoints;
      existing.totalRecords += employeeData.totalRecords;
      existing.records.push(...employeeData.records);

      // Mesclar dados mensais
      Object.entries(employeeData.monthlyData).forEach(([month, data]) => {
        if (!existing.monthlyData[month]) {
          existing.monthlyData[month] = { points: 0, records: 0 };
        }
        existing.monthlyData[month].points += data.points;
        existing.monthlyData[month].records += data.records;
      });

      // Mesclar dados semanais
      Object.entries(employeeData.weeklyData).forEach(([week, data]) => {
        if (!existing.weeklyData[week]) {
          existing.weeklyData[week] = { points: 0, records: 0 };
        }
        existing.weeklyData[week].points += data.points;
        existing.weeklyData[week].records += data.records;
      });
    }
  }

  /**
   * Calcula estatísticas finais
   */
  private static calculateFinalStatistics(result: FolderProcessingResult) {
    result.statistics.totalEmployees = Object.keys(result.employees).length;
    
    Object.values(result.employees).forEach(employee => {
      result.statistics.totalRecords += employee.totalRecords;
      result.statistics.totalPoints += employee.totalPoints;
    });
    
    result.statistics.totalProfit = result.statistics.totalPoints * this.POINT_VALUE;
  }

  /**
   * Gera dados para os gráficos
   */
  private static generateChartData(result: FolderProcessingResult) {
    const employees = Object.keys(result.employees);
    
    // Dados semanais
    const weeklyData = [];
    for (let week = 1; week <= 5; week++) {
      const weekData = { name: `Semana ${week}` };
      employees.forEach(employeeName => {
        const weekKey = `Semana ${week}`;
        weekData[employeeName] = result.employees[employeeName].weeklyData[weekKey]?.points || 0;
      });
      weeklyData.push(weekData);
    }
    result.chartData.weeklyData = weeklyData;

    // Dados mensais
    const allMonths = new Set<string>();
    Object.values(result.employees).forEach(employee => {
      Object.keys(employee.monthlyData).forEach(month => allMonths.add(month));
    });

    const monthlyData = Array.from(allMonths).sort().map(month => {
      const monthData = { name: month };
      employees.forEach(employeeName => {
        monthData[employeeName] = result.employees[employeeName].monthlyData[month]?.points || 0;
      });
      return monthData;
    });
    result.chartData.monthlyData = monthlyData;

    // Performance da equipe (pizza)
    const teamPerformance = employees.map(employeeName => ({
      name: employeeName,
      value: result.employees[employeeName].totalPoints,
      color: this.getEmployeeColor(employeeName)
    }));
    result.chartData.teamPerformance = teamPerformance;
  }

  /**
   * Retorna cor específica para cada funcionário
   */
  private static getEmployeeColor(employeeName: string): string {
    const colorMap: Record<string, string> = {
      'Rodrigo': '#8b5cf6',
      'Maurício': '#f59e0b', 
      'Matheus': '#10b981',
      'Wesley': '#ef4444'
    };
    return colorMap[employeeName] || '#6b7280';
  }

  /**
   * Busca dados de um funcionário específico
   */
  static async getEmployeeData(employeeName: string): Promise<ExcelEmployeeData | null> {
    try {
      const folderData = await this.processFolderStructure();
      return folderData.employees[employeeName] || null;
    } catch (error) {
      console.error(`Erro ao buscar dados de ${employeeName}:`, error);
      return null;
    }
  }

  /**
   * Exporta dados processados para Excel
   */
  static async exportProcessedData(data: FolderProcessingResult): Promise<void> {
    try {
      const wb = XLSX.utils.book_new();

      // Aba de resumo
      const summaryData = [
        ['Estatística', 'Valor'],
        ['Total de Arquivos', data.statistics.totalFiles],
        ['Total de Funcionários', data.statistics.totalEmployees],
        ['Total de Registros', data.statistics.totalRecords],
        ['Total de Pontos', data.statistics.totalPoints],
        ['Lucro Total', `R$ ${data.statistics.totalProfit.toFixed(2)}`]
      ];
      
      const summaryWs = XLSX.utils.aoa_to_sheet(summaryData);
      XLSX.utils.book_append_sheet(wb, summaryWs, 'Resumo');

      // Aba por funcionário
      Object.entries(data.employees).forEach(([name, employee]) => {
        const employeeSheet = employee.records.map(record => ({
          Data: new Date(record.date).toLocaleDateString('pt-BR'),
          Refinaria: record.refinery,
          Pontos: record.points,
          Observações: record.observations,
          Mês: record.month,
          Semana: record.week
        }));
        
        const ws = XLSX.utils.json_to_sheet(employeeSheet);
        XLSX.utils.book_append_sheet(wb, ws, name);
      });

      // Salvar arquivo
      const fileName = `dados_processados_${new Date().toISOString().split('T')[0]}.xlsx`;
      XLSX.writeFile(wb, fileName);
      
    } catch (error) {
      console.error('Erro ao exportar dados processados:', error);
      throw error;
    }
  }

  /**
   * Formata valor monetário
   */
  static formatCurrency(value: number): string {
    return new Intl.NumberFormat('pt-BR', {
      style: 'currency',
      currency: 'BRL'
    }).format(value);
  }
}