import * as XLSX from 'xlsx';
import { CalculationsService } from './CalculationsService';

export interface ExcelEmployeeRecord {
  date: string;
  refinery: string;
  points: number;
  observations: string;
  month: string;
  week: number;
}

export interface ExcelEmployeeData {
  name: string;
  totalPoints: number;
  totalRecords: number;
  records: ExcelEmployeeRecord[];
  monthlyData: Record<string, { points: number; records: number }>;
  weeklyData: Record<string, { points: number; records: number }>;
}

export interface ExcelFolderData {
  employees: Record<string, ExcelEmployeeData>;
  statistics: {
    totalFiles: number;
    totalEmployees: number;
    totalRecords: number;
    totalPoints: number;
    totalProfit: number;
  };
  lastProcessed: string;
}

/**
 * Serviço para ler e processar arquivos Excel da pasta "registros monitorar"
 * Substitui completamente a dependência do Supabase para dados históricos
 */
export class ExcelFolderReader {
  private static readonly POINT_VALUE = 3.25; // R$ 3,25 por ponto
  private static cachedData: ExcelFolderData | null = null;
  private static lastCacheTime: number = 0;
  private static readonly CACHE_DURATION = 2 * 60 * 1000; // 2 minutos

  /**
   * Lê toda a pasta "registros monitorar" e processa os arquivos Excel
   */
  static async readRegistrosFolder(): Promise<ExcelFolderData> {
    // Verificar cache
    const now = Date.now();
    if (this.cachedData && (now - this.lastCacheTime) < this.CACHE_DURATION) {
      console.log('📋 Usando dados em cache da pasta Excel');
      return this.cachedData;
    }

    console.log('🔍 === LENDO PASTA "registros monitorar" ===');
    
    const result: ExcelFolderData = {
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
      // Listar todas as subpastas conhecidas
      const folders = ['mes 4', 'mes 5', 'mes 6', 'mes 7'];
      
      for (const folder of folders) {
        console.log(`📁 Processando pasta: ${folder}`);
        
        // Arquivos conhecidos em cada pasta
        const files = await this.getFilesInFolder(folder);
        
        for (const fileName of files) {
          try {
            console.log(`📄 Processando arquivo: ${fileName}`);
            const filePath = `registros monitorar/${folder}/${fileName}`;
            
            // Verificar se arquivo existe
            if (await this.fileExists(filePath)) {
              const employeeData = await this.processExcelFile(filePath, fileName);
              this.mergeEmployeeData(result, employeeData);
              result.statistics.totalFiles++;
            } else {
              console.warn(`⚠️ Arquivo não encontrado: ${filePath}`);
            }
          } catch (error) {
            console.error(`❌ Erro ao processar ${fileName}:`, error);
          }
        }
      }

      // Calcular estatísticas finais
      this.calculateFinalStatistics(result);
      
      // Atualizar cache
      this.cachedData = result;
      this.lastCacheTime = now;
      
      console.log('✅ Processamento da pasta concluído:', result.statistics);
      return result;
      
    } catch (error) {
      console.error('❌ Erro ao processar pasta registros monitorar:', error);
      throw error;
    }
  }

  /**
   * Lista arquivos conhecidos em cada pasta
   */
  private static async getFilesInFolder(folder: string): Promise<string[]> {
    const knownFiles: Record<string, string[]> = {
      'mes 4': [
        'Matheus Abril.xlsx',
        'Maurício Abril.xlsx', 
        'Rodrigo Abril.xlsx'
      ],
      'mes 5': [
        'Matheus Maio.xlsx',
        'Maurício Maio.xlsx',
        'Wesley Maio.xlsx'
      ],
      'mes 6': [
        'Matheus Junho.xlsx',
        'Maurício Junho.xlsx',
        'Wesley Junho.xlsx'
      ],
      'mes 7': [
        'Matheus Julho.xlsx',
        'Maurício Julho.xlsx',
        'Wesley Julho.xlsx'
      ]
    };

    return knownFiles[folder] || [];
  }

  /**
   * Verifica se arquivo existe
   */
  private static async fileExists(filePath: string): Promise<boolean> {
    try {
      const response = await fetch(`/${filePath}`, { method: 'HEAD' });
      return response.ok;
    } catch {
      return false;
    }
  }

  /**
   * Processa um arquivo Excel individual
   */
  private static async processExcelFile(filePath: string, fileName: string): Promise<ExcelEmployeeData> {
    try {
      // Extrair nome do funcionário do arquivo
      const employeeName = this.extractEmployeeName(fileName);
      
      console.log(`👤 Processando funcionário: ${employeeName}`);
      
      // Ler arquivo Excel
      const response = await fetch(`/${filePath}`);
      if (!response.ok) {
        throw new Error(`Arquivo não encontrado: ${filePath}`);
      }
      
      const arrayBuffer = await response.arrayBuffer();
      const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
      
      // Pegar primeira planilha
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet);

      console.log(`📊 ${jsonData.length} linhas encontradas no Excel`);

      const employeeData: ExcelEmployeeData = {
        name: employeeName,
        totalPoints: 0,
        totalRecords: 0,
        records: [],
        monthlyData: {},
        weeklyData: {}
      };

      // Processar cada linha do Excel
      jsonData.forEach((row: any, index: number) => {
        try {
          const record = this.parseExcelRow(row, employeeName, index);
          if (record) {
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
          console.warn(`⚠️ Erro na linha ${index + 1}:`, error);
        }
      });

      console.log(`✅ ${employeeName}: ${employeeData.totalRecords} registros, ${employeeData.totalPoints} pontos`);
      return employeeData;
      
    } catch (error) {
      console.error(`❌ Erro ao processar ${filePath}:`, error);
      throw error;
    }
  }

  /**
   * Converte linha do Excel para registro estruturado
   */
  private static parseExcelRow(row: any, employeeName: string, rowIndex: number): ExcelEmployeeRecord | null {
    try {
      // Buscar colunas com nomes alternativos (case insensitive)
      const dateValue = this.findColumnValue(row, ['Data', 'data', 'DATE', 'Date']);
      const pointsValue = this.findColumnValue(row, ['Pontos', 'pontos', 'PONTOS', 'Points']);
      const refineryValue = this.findColumnValue(row, ['Refinaria', 'refinaria', 'REFINARIA', 'Refinery']);
      const observationsValue = this.findColumnValue(row, ['Observações', 'observacoes', 'OBSERVACOES', 'Observations']);

      // Validar dados essenciais
      if (!dateValue) {
        return null; // Linha sem data válida
      }

      const points = parseFloat(pointsValue) || 0;
      if (points <= 0) {
        return null; // Linha sem pontos válidos
      }

      // Converter data do Excel
      const parsedDate = this.parseExcelDate(dateValue);
      if (!parsedDate || isNaN(parsedDate.getTime())) {
        console.warn(`Data inválida na linha ${rowIndex + 1}: ${dateValue}`);
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
        week: week
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
      if (row.hasOwnProperty(name) && row[name] !== undefined && row[name] !== null && row[name] !== '') {
        return row[name];
      }
    }
    return null;
  }

  /**
   * Converte valor de data do Excel para Date
   */
  private static parseExcelDate(dateValue: any): Date | null {
    try {
      if (dateValue instanceof Date) {
        return dateValue;
      }
      
      if (typeof dateValue === 'string') {
        // Formato brasileiro: DD/MM/YYYY ou DD/MM/YY
        if (dateValue.includes('/')) {
          const parts = dateValue.split('/');
          if (parts.length === 3) {
            let day = parseInt(parts[0]);
            let month = parseInt(parts[1]) - 1; // 0-based
            let year = parseInt(parts[2]);
            
            // Ajustar ano de 2 dígitos
            if (year < 100) {
              year += year < 50 ? 2000 : 1900;
            }
            
            if (!isNaN(day) && !isNaN(month) && !isNaN(year)) {
              return new Date(year, month, day);
            }
          }
        }
        
        // Tentar outros formatos
        const parsed = new Date(dateValue);
        if (!isNaN(parsed.getTime())) {
          return parsed;
        }
      }
      
      if (typeof dateValue === 'number') {
        // Excel serial date (dias desde 1900-01-01)
        // Correção para bug do Excel (ano 1900 não é bissexto)
        const excelEpoch = new Date(1899, 11, 30); // 30 de dezembro de 1899
        const date = new Date(excelEpoch.getTime() + dateValue * 24 * 60 * 60 * 1000);
        return date;
      }
      
      return null;
      
    } catch (error) {
      return null;
    }
  }

  /**
   * Determina mês baseado na lógica 26→25 da empresa
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
  private static mergeEmployeeData(result: ExcelFolderData, employeeData: ExcelEmployeeData) {
    const name = employeeData.name;
    
    if (!result.employees[name]) {
      result.employees[name] = employeeData;
    } else {
      // Mesclar com dados existentes
      const existing = result.employees[name];
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
  private static calculateFinalStatistics(result: ExcelFolderData) {
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
  static async generateChartData(): Promise<{
    weeklyData: any[];
    monthlyData: any[];
    teamPerformance: any[];
  }> {
    const folderData = await this.readRegistrosFolder();
    const employees = Object.keys(folderData.employees);

    console.log('📊 Gerando dados dos gráficos a partir dos arquivos Excel');
    console.log('👥 Funcionários encontrados:', employees);

    // Dados semanais (5 semanas do ciclo)
    const weeklyData = [];
    for (let week = 1; week <= 5; week++) {
      const weekData = { name: `Semana ${week}` };
      employees.forEach(employeeName => {
        const weekKey = `Semana ${week}`;
        weekData[employeeName] = folderData.employees[employeeName].weeklyData[weekKey]?.points || 0;
      });
      weeklyData.push(weekData);
    }

    // Dados mensais
    const allMonths = new Set<string>();
    Object.values(folderData.employees).forEach(employee => {
      Object.keys(employee.monthlyData).forEach(month => allMonths.add(month));
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
        monthData[employeeName] = folderData.employees[employeeName].monthlyData[month]?.points || 0;
      });
      return monthData;
    });

    // Performance da equipe (gráfico de pizza)
    const teamPerformance = employees.map(employeeName => ({
      name: employeeName,
      value: folderData.employees[employeeName].totalPoints,
      color: this.getEmployeeColor(employeeName)
    }));

    console.log('📊 Dados dos gráficos gerados:', {
      weeklyData: weeklyData.length,
      monthlyData: monthlyData.length,
      teamPerformance: teamPerformance.length
    });

    return {
      weeklyData,
      monthlyData,
      teamPerformance
    };
  }

  /**
   * Calcula estatísticas gerais baseadas nos arquivos Excel
   */
  static async getGeneralStats() {
    try {
      const folderData = await this.readRegistrosFolder();
      const employees = Object.values(folderData.employees);

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

        // Média da equipe (excluindo Rodrigo se for freelancer)
        if (employee.name !== 'Rodrigo') {
          totalPointsForAverage += employee.totalPoints;
          employeeCountForAverage++;
        }
      });

      const avgTeam = employeeCountForAverage > 0 ? 
        Math.round(totalPointsForAverage / employeeCountForAverage) : 0;
      
      const totalGoal = 29500; // Meta mensal da equipe
      const progressPercentage = (folderData.statistics.totalPoints / totalGoal) * 100;

      return {
        bestPerformer,
        bestPoints,
        avgTeam,
        totalGoal: Math.round(totalGoal / 1000 * 10) / 10, // 29.5K
        progressPercentage: Math.round(progressPercentage * 10) / 10
      };
      
    } catch (error) {
      console.error('Erro ao calcular estatísticas dos arquivos Excel:', error);
      return {
        bestPerformer: '',
        bestPoints: 0,
        avgTeam: 0,
        totalGoal: 29.5,
        progressPercentage: 0
      };
    }
  }

  /**
   * Busca dados de um funcionário específico
   */
  static async getEmployeeData(employeeName: string): Promise<ExcelEmployeeData | null> {
    try {
      const folderData = await this.readRegistrosFolder();
      return folderData.employees[employeeName] || null;
    } catch (error) {
      console.error(`Erro ao buscar dados de ${employeeName}:`, error);
      return null;
    }
  }

  /**
   * Limpa cache forçando reprocessamento
   */
  static clearCache() {
    this.cachedData = null;
    this.lastCacheTime = 0;
    console.log('🗑️ Cache da pasta Excel limpo');
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
   * Exporta dados processados para Excel
   */
  static async exportProcessedData(): Promise<void> {
    try {
      const folderData = await this.readRegistrosFolder();
      
      // Criar workbook
      const wb = XLSX.utils.book_new();

      // Aba de resumo
      const summaryData = [
        ['Estatística', 'Valor'],
        ['Total de Arquivos', folderData.statistics.totalFiles],
        ['Total de Funcionários', folderData.statistics.totalEmployees],
        ['Total de Registros', folderData.statistics.totalRecords],
        ['Total de Pontos', folderData.statistics.totalPoints],
        ['Lucro Total', this.formatCurrency(folderData.statistics.totalProfit)]
      ];
      
      const summaryWs = XLSX.utils.aoa_to_sheet(summaryData);
      XLSX.utils.book_append_sheet(wb, summaryWs, 'Resumo');

      // Aba por funcionário
      Object.entries(folderData.employees).forEach(([name, employee]) => {
        const employeeSheet = employee.records.map(record => ({
          Data: new Date(record.date).toLocaleDateString('pt-BR'),
          Refinaria: record.refinery,
          Pontos: record.points,
          Observações: record.observations,
          Mês: record.month,
          Semana: record.week
        }));
        
        const ws = XLSX.utils.json_to_sheet(employeeSheet);
        
        // Ajustar largura das colunas
        const colWidths = [
          { wch: 12 }, // Data
          { wch: 12 }, // Refinaria
          { wch: 8 },  // Pontos
          { wch: 40 }, // Observações
          { wch: 10 }, // Mês
          { wch: 8 }   // Semana
        ];
        ws['!cols'] = colWidths;
        
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