import * as XLSX from 'xlsx';

export interface ExcelRecord {
  employee: string;
  date: Date;
  points: number;
  month: string;
  refinery?: string;
}

export interface EmployeeData {
  total: number;
  records: number;
  months: { [key: string]: { points: number; records: number } };
  refineries: { [key: string]: number };
}

export interface ProcessedExcelData {
  employees: { [key: string]: EmployeeData };
  months: { [key: string]: any };
  records: ExcelRecord[];
  statistics: {
    total_files: number;
    total_employees: number;
    total_records: number;
    total_points: number;
    total_profit: number; // R$ 3.25 por ponto
  };
}

export class ExcelProcessorService {
  private static readonly POINT_VALUE = 3.25; // R$ 3,25 por ponto

  static async processExcelFiles(files: FileList): Promise<ProcessedExcelData> {
    const processedData: ProcessedExcelData = {
      employees: {},
      months: {},
      records: [],
      statistics: {
        total_files: files.length,
        total_employees: 0,
        total_records: 0,
        total_points: 0,
        total_profit: 0
      }
    };

    for (const file of Array.from(files)) {
      if (file.name.match(/\.(xlsx|xls)$/i)) {
        try {
          const fileData = await this.extractDataFromExcel(file);
          this.mergeData(processedData, fileData);
        } catch (error) {
          console.error(`Erro ao processar ${file.name}:`, error);
        }
      }
    }

    this.calculateFinalStatistics(processedData);
    return processedData;
  }

  private static async extractDataFromExcel(file: File): Promise<Partial<ProcessedExcelData>> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target?.result as ArrayBuffer);
          const workbook = XLSX.read(data, { type: 'array' });
          const sheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[sheetName];
          const jsonData = XLSX.utils.sheet_to_json(worksheet);

          const employeeName = file.name.replace(/\.(xlsx|xls)$/i, '');
          const fileData: Partial<ProcessedExcelData> = {
            employees: {},
            records: []
          };

          // Inicializar dados do funcionário
          if (employeeName) {
            fileData.employees![employeeName] = {
              total: 0,
              records: 0,
              months: {},
              refineries: {}
            };
          }

          // Processar cada linha
          jsonData.forEach((row: any) => {
            try {
              const dateValue = this.parseDate(row.Data || row.data || row.DATE);
              const pointsValue = parseFloat(row.Pontos || row.pontos || row.PONTOS || row.Total || row.total || 0);
              const refineryValue = String(row.Refinaria || row.refinaria || row.REFINARIA || '').trim();

              if (pointsValue > 0 && employeeName) {
                const monthKey = this.getMonthFromDate(dateValue);
                
                // Adicionar aos totais do funcionário
                fileData.employees![employeeName].total += pointsValue;
                fileData.employees![employeeName].records += 1;

                // Adicionar aos dados mensais
                if (!fileData.employees![employeeName].months[monthKey]) {
                  fileData.employees![employeeName].months[monthKey] = {
                    points: 0,
                    records: 0
                  };
                }
                fileData.employees![employeeName].months[monthKey].points += pointsValue;
                fileData.employees![employeeName].months[monthKey].records += 1;

                // Adicionar refinaria
                if (refineryValue) {
                  if (!fileData.employees![employeeName].refineries[refineryValue]) {
                    fileData.employees![employeeName].refineries[refineryValue] = 0;
                  }
                  fileData.employees![employeeName].refineries[refineryValue] += pointsValue;
                }

                // Adicionar registro individual
                fileData.records!.push({
                  employee: employeeName,
                  date: dateValue,
                  points: pointsValue,
                  month: monthKey,
                  refinery: refineryValue
                });
              }
            } catch (error) {
              console.warn('Erro ao processar linha:', error);
            }
          });

          resolve(fileData);
        } catch (error) {
          reject(error);
        }
      };

      reader.onerror = () => reject(new Error('Erro ao ler arquivo'));
      reader.readAsArrayBuffer(file);
    });
  }

  private static parseDate(dateValue: any): Date {
    if (!dateValue) return new Date();
    
    // Tentar diferentes formatos de data
    if (dateValue instanceof Date) return dateValue;
    
    if (typeof dateValue === 'string') {
      // Tentar formatos comuns: DD/MM/YYYY, DD-MM-YYYY, etc.
      const dateStr = dateValue.replace(/[-]/g, '/');
      const parsed = new Date(dateStr);
      if (!isNaN(parsed.getTime())) return parsed;
    }
    
    if (typeof dateValue === 'number') {
      // Excel serial date
      return new Date((dateValue - 25569) * 86400 * 1000);
    }
    
    return new Date();
  }

  private static getMonthFromDate(date: Date): string {
    try {
      const day = date.getDate();
      const month = date.getMonth() + 1;
      const year = date.getFullYear();

      // Lógica: se dia >= 26, pertence ao mês atual
      // Se dia < 26, pertence ao mês anterior
      let targetMonth = month;
      let targetYear = year;

      if (day >= 26) {
        // Já no próximo ciclo - manter mês e ano atuais
        targetMonth = month;
        targetYear = year;
      } else {
        // Ainda no ciclo anterior - ir para mês anterior
        if (month === 1) {
          targetMonth = 12;
          targetYear = year - 1;
        } else {
          targetMonth = month - 1;
        }
      }

      return `${targetMonth.toString().padStart(2, '0')}/${targetYear}`;
    } catch (error) {
      return 'Sem Data';
    }
  }

  private static mergeData(processedData: ProcessedExcelData, fileData: Partial<ProcessedExcelData>) {
    // Mesclar funcionários
    if (fileData.employees) {
      Object.entries(fileData.employees).forEach(([employee, data]) => {
        if (!processedData.employees[employee]) {
          processedData.employees[employee] = data;
        } else {
          // Somar dados
          processedData.employees[employee].total += data.total;
          processedData.employees[employee].records += data.records;
          
          // Mesclar meses
          Object.entries(data.months).forEach(([month, monthData]) => {
            if (!processedData.employees[employee].months[month]) {
              processedData.employees[employee].months[month] = monthData;
            } else {
              processedData.employees[employee].months[month].points += monthData.points;
              processedData.employees[employee].months[month].records += monthData.records;
            }
          });
          
          // Mesclar refinarias
          Object.entries(data.refineries).forEach(([refinery, points]) => {
            if (!processedData.employees[employee].refineries[refinery]) {
              processedData.employees[employee].refineries[refinery] = 0;
            }
            processedData.employees[employee].refineries[refinery] += points;
          });
        }
      });
    }

    // Adicionar registros
    if (fileData.records) {
      processedData.records.push(...fileData.records);
    }
  }

  private static calculateFinalStatistics(processedData: ProcessedExcelData) {
    processedData.statistics.total_employees = Object.keys(processedData.employees).length;
    processedData.statistics.total_records = processedData.records.length;
    
    // Calcular pontos totais
    processedData.statistics.total_points = Object.values(processedData.employees)
      .reduce((sum, emp) => sum + emp.total, 0);
    
    // Calcular lucro total (pontos × R$ 3,25)
    processedData.statistics.total_profit = processedData.statistics.total_points * this.POINT_VALUE;
  }

  static formatCurrency(value: number): string {
    return new Intl.NumberFormat('pt-BR', {
      style: 'currency',
      currency: 'BRL'
    }).format(value);
  }

  static calculateMonthlyData(processedData: ProcessedExcelData): any[] {
    const monthlyData: { [key: string]: { [employee: string]: number } } = {};
    
    // Agrupar por mês
    Object.entries(processedData.employees).forEach(([employee, data]) => {
      Object.entries(data.months).forEach(([month, monthData]) => {
        if (!monthlyData[month]) {
          monthlyData[month] = {};
        }
        monthlyData[month][employee] = monthData.points;
      });
    });
    
    // Converter para formato do gráfico
    return Object.entries(monthlyData)
      .sort(([a], [b]) => a.localeCompare(b))
      .map(([month, employees]) => ({
        name: month,
        ...employees
      }));
  }
}