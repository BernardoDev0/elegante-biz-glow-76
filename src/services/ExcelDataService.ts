import * as XLSX from 'xlsx';
import { FileSystemService } from './FileSystemService';
import { CalculationsService } from './CalculationsService';

export interface ProcessedExcelData {
  employees: Record<string, EmployeeExcelData>;
  statistics: {
    totalFiles: number;
    totalEmployees: number;
    totalRecords: number;
    totalPoints: number;
    totalProfit: number;
  };
  lastProcessed: string;
}

export interface EmployeeExcelData {
  name: string;
  totalPoints: number;
  totalRecords: number;
  records: ExcelRecord[];
  monthlyBreakdown: Record<string, { points: number; records: number }>;
  weeklyBreakdown: Record<string, { points: number; records: number }>;
}

export interface ExcelRecord {
  date: string;
  refinery: string;
  points: number;
  observations: string;
  month: string;
  week: number;
  originalRow: any;
}

/**
 * Serviço principal para processar dados da pasta "registros monitorar"
 * Substitui a dependência do Supabase por leitura direta dos arquivos Excel
 */
export class ExcelDataService {
  private static readonly POINT_VALUE = 3.25;
  private static cachedData: ProcessedExcelData | null = null;
  private static lastCacheTime: number = 0;
  private static readonly CACHE_DURATION = 5 * 60 * 1000; // 5 minutos

  /**
   * Processa toda a pasta "registros monitorar" e retorna dados estruturados
   */
  static async processRegistrosFolder(): Promise<ProcessedExcelData> {
    // Verificar cache
    const now = Date.now();
    if (this.cachedData && (now - this.lastCacheTime) < this.CACHE_DURATION) {
      console.log('📋 Usando dados em cache');
      return this.cachedData;
    }

    console.log('🔍 Processando pasta "registros monitorar"...');
    
    const result: ProcessedExcelData = {
      employees: {},
      statistics: {
        totalFiles: 0,
        totalEmployees: 0,
        totalRecords: 0,
        totalPoints: 0,
        totalProfit: 0
      },
      lastProcessed: new Date().toISOString()
    };

    try {
      // Listar todos os arquivos Excel
      const excelFiles = await FileSystemService.listExcelFiles();
      console.log(`📁 Encontrados ${excelFiles.length} arquivos Excel`);

      // Processar cada arquivo
      for (const file of excelFiles) {
        try {
          console.log(`📄 Processando: ${file.name}`);
          const employeeData = await this.processExcelFile(file);
          this.mergeEmployeeData(result, employeeData);
          result.statistics.totalFiles++;
        } catch (error) {
          console.error(`❌ Erro ao processar ${file.name}:`, error);
        }
      }

      // Calcular estatísticas finais
      this.calculateStatistics(result);
      
      // Atualizar cache
      this.cachedData = result;
      this.lastCacheTime = now;
      
      console.log('✅ Processamento concluído:', result.statistics);
      return result;
      
    } catch (error) {
      console.error('❌ Erro ao processar pasta:', error);
      throw error;
    }
  }

  /**
   * Processa um arquivo Excel individual
   */
  private static async processExcelFile(file: { name: string; path: string; folder: string }): Promise<EmployeeExcelData> {
    try {
      // Extrair nome do funcionário
      const employeeName = this.extractEmployeeName(file.name);
      
      // Ler arquivo Excel
      const arrayBuffer = await FileSystemService.readExcelFile(file.path);
      const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
      
      // Pegar primeira planilha
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet);

      const employeeData: EmployeeExcelData = {
        name: employeeName,
        totalPoints: 0,
        totalRecords: 0,
        records: [],
        monthlyBreakdown: {},
        weeklyBreakdown: {}
      };

      // Processar cada linha
      jsonData.forEach((row: any, index: number) => {
        try {
          const record = this.parseExcelRow(row, employeeName, index);
          if (record) {
            employeeData.records.push(record);
            employeeData.totalPoints += record.points;
            employeeData.totalRecords++;

            // Agrupar por mês
            if (!employeeData.monthlyBreakdown[record.month]) {
              employeeData.monthlyBreakdown[record.month] = { points: 0, records: 0 };
            }
            employeeData.monthlyBreakdown[record.month].points += record.points;
            employeeData.monthlyBreakdown[record.month].records++;

            // Agrupar por semana
            const weekKey = `Semana ${record.week}`;
            if (!employeeData.weeklyBreakdown[weekKey]) {
              employeeData.weeklyBreakdown[weekKey] = { points: 0, records: 0 };
            }
            employeeData.weeklyBreakdown[weekKey].points += record.points;
            employeeData.weeklyBreakdown[weekKey].records++;
          }
        } catch (error) {
          console.warn(`Erro na linha ${index + 1} de ${file.name}:`, error);
        }
      });

      console.log(`✅ ${employeeName}: ${employeeData.totalRecords} registros, ${employeeData.totalPoints} pontos`);
      return employeeData;
      
    } catch (error) {
      console.error(`Erro ao processar ${file.name}:`, error);
      throw error;
    }
  }

  /**
   * Converte linha do Excel para registro estruturado
   */
  private static parseExcelRow(row: any, employeeName: string, rowIndex: number): ExcelRecord | null {
    try {
      // Tentar diferentes nomes de colunas (case insensitive)
      const dateValue = this.findColumnValue(row, ['data', 'date', 'Data', 'DATE']);
      const pointsValue = this.findColumnValue(row, ['pontos', 'points', 'Pontos', 'PONTOS']);
      const refineryValue = this.findColumnValue(row, ['refinaria', 'refinery', 'Refinaria', 'REFINARIA']);
      const observationsValue = this.findColumnValue(row, ['observacoes', 'observations', 'Observações', 'OBSERVACOES']);

      // Validar dados essenciais
      if (!dateValue) {
        console.warn(`Linha ${rowIndex + 1}: Data não encontrada`);
        return null;
      }

      const points = parseFloat(pointsValue) || 0;
      if (points <= 0) {
        console.warn(`Linha ${rowIndex + 1}: Pontos inválidos (${pointsValue})`);
        return null;
      }

      // Converter data
      const parsedDate = this.parseDate(dateValue);
      if (!parsedDate || isNaN(parsedDate.getTime())) {
        console.warn(`Linha ${rowIndex + 1}: Data inválida (${dateValue})`);
        return null;
      }

      // Calcular mês e semana baseado na lógica 26→25
      const month = this.getMonthFromDate(parsedDate);
      const week = CalculationsService.getWeekFromDate(parsedDate.toISOString().split('T')[0]);

      return {
        date: parsedDate.toISOString(),
        refinery: String(refineryValue || '').trim(),
        points: points,
        observations: String(observationsValue || '').trim(),
        month: month,
        week: week,
        originalRow: row
      };
      
    } catch (error) {
      console.warn(`Erro ao processar linha ${rowIndex + 1}:`, error);
      return null;
    }
  }

  /**
   * Busca valor de coluna com nomes alternativos
   */
  private static findColumnValue(row: any, possibleNames: string[]): any {
    for (const name of possibleNames) {
      if (row.hasOwnProperty(name) && row[name] !== undefined && row[name] !== null) {
        return row[name];
      }
    }
    return null;
  }

  /**
   * Converte diferentes formatos de data para Date
   */
  private static parseDate(dateValue: any): Date | null {
    try {
      if (dateValue instanceof Date) {
        return dateValue;
      }
      
      if (typeof dateValue === 'string') {
        // Formato brasileiro: DD/MM/YYYY
        if (dateValue.includes('/')) {
          const parts = dateValue.split('/');
          if (parts.length === 3) {
            const day = parseInt(parts[0]);
            const month = parseInt(parts[1]) - 1; // 0-based
            const year = parseInt(parts[2]);
            
            if (!isNaN(day) && !isNaN(month) && !isNaN(year)) {
              return new Date(year, month, day);
            }
          }
        }
        
        // Tentar parse direto
        const parsed = new Date(dateValue);
        if (!isNaN(parsed.getTime())) {
          return parsed;
        }
      }
      
      if (typeof dateValue === 'number') {
        // Excel serial date (dias desde 1900-01-01)
        const excelEpoch = new Date(1900, 0, 1);
        const date = new Date(excelEpoch.getTime() + (dateValue - 2) * 24 * 60 * 60 * 1000);
        return date;
      }
      
      return null;
      
    } catch (error) {
      console.warn('Erro ao converter data:', error);
      return null;
    }
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

    // Lógica 26→25: se dia >= 26, pertence ao próximo mês da empresa
    if (day >= 26) {
      targetMonth = month + 1;
      if (targetMonth > 12) {
        targetMonth = 1;
      }
    }

    return monthNames[targetMonth - 1];
  }

  /**
   * Extrai nome do funcionário do nome do arquivo
   */
  private static extractEmployeeName(fileName: string): string {
    // Remove extensão
    const nameWithoutExt = fileName.replace(/\.(xlsx|xls)$/i, '');
    
    // Extrair nome (antes do mês)
    const parts = nameWithoutExt.split(' ');
    return parts[0];
  }

  /**
   * Mescla dados de funcionário no resultado
   */
  private static mergeEmployeeData(result: ProcessedExcelData, employeeData: EmployeeExcelData) {
    const name = employeeData.name;
    
    if (!result.employees[name]) {
      result.employees[name] = employeeData;
    } else {
      // Mesclar com dados existentes
      const existing = result.employees[name];
      existing.totalPoints += employeeData.totalPoints;
      existing.totalRecords += employeeData.totalRecords;
      existing.records.push(...employeeData.records);

      // Mesclar breakdown mensal
      Object.entries(employeeData.monthlyBreakdown).forEach(([month, data]) => {
        if (!existing.monthlyBreakdown[month]) {
          existing.monthlyBreakdown[month] = { points: 0, records: 0 };
        }
        existing.monthlyBreakdown[month].points += data.points;
        existing.monthlyBreakdown[month].records += data.records;
      });

      // Mesclar breakdown semanal
      Object.entries(employeeData.weeklyBreakdown).forEach(([week, data]) => {
        if (!existing.weeklyBreakdown[week]) {
          existing.weeklyBreakdown[week] = { points: 0, records: 0 };
        }
        existing.weeklyBreakdown[week].points += data.points;
        existing.weeklyBreakdown[week].records += data.records;
      });
    }
  }

  /**
   * Calcula estatísticas finais
   */
  private static calculateStatistics(result: ProcessedExcelData) {
    result.statistics.totalEmployees = Object.keys(result.employees).length;
    
    Object.values(result.employees).forEach(employee => {
      result.statistics.totalRecords += employee.totalRecords;
      result.statistics.totalPoints += employee.totalPoints;
    });
    
    result.statistics.totalProfit = result.statistics.totalPoints * this.POINT_VALUE;
  }

  /**
   * Gera dados para gráficos baseados nos arquivos Excel
   */
  static async generateChartDataFromExcel(): Promise<{
    weeklyData: any[];
    monthlyData: any[];
    teamPerformance: any[];
  }> {
    const processedData = await this.processRegistrosFolder();
    const employees = Object.keys(processedData.employees);

    // Dados semanais
    const weeklyData = [];
    for (let week = 1; week <= 5; week++) {
      const weekData = { name: `Semana ${week}` };
      employees.forEach(employeeName => {
        const weekKey = `Semana ${week}`;
        weekData[employeeName] = processedData.employees[employeeName].weeklyBreakdown[weekKey]?.points || 0;
      });
      weeklyData.push(weekData);
    }

    // Dados mensais
    const allMonths = new Set<string>();
    Object.values(processedData.employees).forEach(employee => {
      Object.keys(employee.monthlyBreakdown).forEach(month => allMonths.add(month));
    });

    const sortedMonths = Array.from(allMonths).sort((a, b) => {
      const monthOrder = [
        'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
        'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'
      ];
      return monthOrder.indexOf(a) - monthOrder.indexOf(b);
    });

    const monthlyData = sortedMonths.map(month => {
      const monthData = { name: month };
      employees.forEach(employeeName => {
        monthData[employeeName] = processedData.employees[employeeName].monthlyBreakdown[month]?.points || 0;
      });
      return monthData;
    });

    // Performance da equipe
    const teamPerformance = employees.map(employeeName => ({
      name: employeeName,
      value: processedData.employees[employeeName].totalPoints,
      color: this.getEmployeeColor(employeeName)
    }));

    return {
      weeklyData,
      monthlyData,
      teamPerformance
    };
  }

  /**
   * Busca dados de um funcionário específico
   */
  static async getEmployeeDataFromExcel(employeeName: string): Promise<EmployeeExcelData | null> {
    try {
      const processedData = await this.processRegistrosFolder();
      return processedData.employees[employeeName] || null;
    } catch (error) {
      console.error(`Erro ao buscar dados de ${employeeName}:`, error);
      return null;
    }
  }

  /**
   * Calcula estatísticas gerais baseadas nos arquivos Excel
   */
  static async getGeneralStatsFromExcel() {
    try {
      const processedData = await this.processRegistrosFolder();
      const employees = Object.values(processedData.employees);

      let bestPerformer = '';
      let bestPoints = 0;
      let totalPointsForAverage = 0;
      let employeeCountForAverage = 0;

      employees.forEach(employee => {
        // Melhor performer
        if (employee.totalPoints > bestPoints) {
          bestPoints = employee.totalPoints;
          bestPerformer = employee.name;
        }

        // Média (excluindo Rodrigo se for freelancer)
        if (employee.name !== 'Rodrigo') {
          totalPointsForAverage += employee.totalPoints;
          employeeCountForAverage++;
        }
      });

      const avgTeam = employeeCountForAverage > 0 ? 
        Math.round(totalPointsForAverage / employeeCountForAverage) : 0;
      
      const totalGoal = 29500; // Meta mensal da equipe
      const progressPercentage = (processedData.statistics.totalPoints / totalGoal) * 100;

      return {
        bestPerformer,
        bestPoints,
        avgTeam,
        totalGoal: Math.round(totalGoal / 1000 * 10) / 10,
        progressPercentage: Math.round(progressPercentage * 10) / 10
      };
      
    } catch (error) {
      console.error('Erro ao calcular estatísticas:', error);
      return null;
    }
  }

  /**
   * Limpa cache forçando reprocessamento
   */
  static clearCache() {
    this.cachedData = null;
    this.lastCacheTime = 0;
    console.log('🗑️ Cache limpo');
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
}