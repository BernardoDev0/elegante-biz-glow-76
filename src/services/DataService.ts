import { supabase } from '@/integrations/supabase/client';
import { Employee, Entry } from './EmployeeService';
import { CalculationsService } from './CalculationsService';
import { ExcelDataService } from './ExcelDataService';

export interface ChartData {
  weeklyData: any[];
  monthlyData: any[];
  teamPerformance: any[];
  employeeStats: Record<string, any>;
}

export interface GeneralStats {
  bestPerformer: string;
  bestPoints: number;
  avgTeam: number;
  totalGoal: number;
  progressPercentage: number;
}

/**
 * Serviço centralizado para busca de dados do Supabase
 * Responsável por abstrair queries complexas e fornecer dados estruturados
 * 
 * MODO HÍBRIDO: Pode usar Supabase OU arquivos Excel da pasta "registros monitorar"
 */
export class DataService {
  private static useExcelMode = true; // Alternar entre Excel e Supabase
  
  private static readonly EMPLOYEE_COLORS = {
    'Rodrigo': '#8b5cf6',
    'Maurício': '#f59e0b', 
    'Matheus': '#10b981',
    'Wesley': '#ef4444'
  };

  /**
   * Alterna entre modo Excel e Supabase
   */
  static setDataSource(useExcel: boolean) {
    this.useExcelMode = useExcel;
    console.log(`📊 Modo de dados alterado para: ${useExcel ? 'Excel' : 'Supabase'}`);
  }

  /**
   * Busca todos os funcionários com cache
   */
  static async getEmployees(): Promise<Employee[]> {
    const { data, error } = await supabase
      .from('employee')
      .select('*')
      .order('real_name');

    if (error) {
      console.error('Erro ao buscar funcionários:', error);
      return [];
    }

    return data as Employee[];
  }

  /**
   * Busca entradas com filtros e paginação
   */
  static async getEntries(filters?: {
    employeeId?: number;
    startDate?: string;
    endDate?: string;
    limit?: number;
  }): Promise<Entry[]> {
    let query = supabase
      .from('entry')
      .select('*')
      .order('date', { ascending: false });

    if (filters?.employeeId) {
      query = query.eq('employee_id', filters.employeeId);
    }
    if (filters?.startDate) {
      query = query.gte('date', filters.startDate);
    }
    if (filters?.endDate) {
      query = query.lte('date', filters.endDate);
    }
    if (filters?.limit) {
      query = query.limit(filters.limit);
    }

    const { data, error } = await query;
    
    if (error) {
      console.error('Erro ao buscar entradas:', error);
      return [];
    }

    return data as Entry[];
  }

  /**
   * Calcula pontos de um funcionário em um período
   */
  static async getEmployeePoints(
    employeeId: number, 
    startDate: string, 
    endDate: string
  ): Promise<number> {
    const { data, error } = await supabase
      .from('entry')
      .select('points')
      .eq('employee_id', employeeId)
      .gte('date', startDate)
      .lte('date', endDate);

    if (error) {
      console.error('Erro ao calcular pontos:', error);
      return 0;
    }

    return data?.reduce((sum, entry) => sum + (entry.points || 0), 0) || 0;
  }

  /**
   * Gera dados para gráficos semanais
   */
  static async getWeeklyChartData(): Promise<any[]> {
    const employees = await this.getEmployees();
    if (!employees.length) return [];

    const weeklyData = [];
    const cycleStart = CalculationsService.getCurrentCycleStart();
    
    // Gerar 5 semanas do ciclo
    for (let week = 1; week <= 5; week++) {
      const weekDates = CalculationsService.getWeekDates(week.toString());
      const weekData = { name: `Semana ${week}` };
      
      for (const employee of employees) {
        const points = await this.getEmployeePoints(
          employee.id, 
          weekDates.start, 
          weekDates.end
        );
        weekData[employee.real_name] = points;
      }
      
      weeklyData.push(weekData);
    }

    return weeklyData;
  }

  /**
   * Gera dados para gráficos mensais
   * IMPORTANTE: Usa o período customizado da empresa (26→25)
   */
  static async getMonthlyChartData(): Promise<any[]> {
    const employees = await this.getEmployees();
    if (!employees.length) return [];

    const monthlyData = [];
    const months = [
      'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
      'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'
    ];
    
    const currentDate = new Date();
    const currentDay = currentDate.getDate();
    const currentCalendarMonth = currentDate.getMonth(); // 0-based (0=Janeiro)
    const currentYear = currentDate.getFullYear();
    
    // Determinar qual é o "mês atual da empresa" baseado na regra 26-25
    let currentCompanyMonth = currentCalendarMonth;
    
    console.log('🔍 DEBUG DataService:');
    console.log('📅 Data atual:', currentDate.toLocaleDateString('pt-BR'));
    console.log('📊 Dia atual:', currentDay);
    console.log('📊 Mês calendário atual (0-based):', currentCalendarMonth);
    
    // Se já passou do dia 25, estamos no próximo mês da empresa
    if (currentDay >= 26) {
      // Estamos no próximo ciclo - manter mês atual
      currentCompanyMonth = currentCalendarMonth;
      console.log('✅ Já passou do dia 25 - estamos no próximo ciclo');
    } else {
      // Ainda no ciclo anterior - ir para mês anterior
      currentCompanyMonth = currentCalendarMonth - 1;
      if (currentCompanyMonth < 0) {
        currentCompanyMonth = 11; // Dezembro do ano anterior
      }
      console.log('⚠️ Ainda no ciclo anterior');
    }
    
    console.log('🎯 Mês atual da empresa (0-based):', currentCompanyMonth);
    console.log('🎯 Nome do mês atual da empresa:', months[currentCompanyMonth]);
    
    // Mostrar últimos 7 meses da empresa (incluindo o atual)
    const monthsToShow = [];
    for (let i = 6; i >= 0; i--) {
      let monthIndex = currentCompanyMonth - i;
      let yearForMonth = currentYear;
      
      if (monthIndex < 0) {
        monthIndex += 12;
        yearForMonth -= 1;
      }
      
      monthsToShow.push({ monthIndex, yearForMonth });
    }
    
    for (const { monthIndex, yearForMonth } of monthsToShow) {
      // monthIndex é 0-based, mas getMonthCycleDates espera 1-based
      const monthDates = CalculationsService.getMonthCycleDates(monthIndex + 1, yearForMonth);
      const monthData = { name: months[monthIndex] };
      
      console.log(`📅 Processando mês: ${months[monthIndex]} (${monthIndex + 1}/${yearForMonth})`);
      console.log(`📅 Período: ${monthDates.start} até ${monthDates.end}`);
      
      for (const employee of employees) {
        const points = await this.getEmployeePoints(
          employee.id,
          monthDates.start,
          monthDates.end
        );
        monthData[employee.real_name] = points;
      }
      
      monthlyData.push(monthData);
    }

    return monthlyData;
  }

  /**
   * Gera dados para gráfico de pizza da equipe
   */
  static async getTeamPerformanceData(): Promise<any[]> {
    const employees = await this.getEmployees();
    if (!employees.length) return [];

    const monthDates = CalculationsService.getMonthCycleDates();
    const teamData = [];

    for (const employee of employees) {
      const points = await this.getEmployeePoints(
        employee.id,
        monthDates.start,
        monthDates.end
      );

      teamData.push({
        name: employee.real_name,
        value: points,
        color: this.EMPLOYEE_COLORS[employee.real_name] || '#6b7280'
      });
    }

    return teamData;
  }

  /**
   * Gera todos os dados de gráficos de uma vez
   * Agora suporta tanto Supabase quanto arquivos Excel
   */
  static async getChartData(): Promise<ChartData> {
    if (this.useExcelMode) {
      console.log('📊 Usando dados dos arquivos Excel');
      const excelChartData = await ExcelDataService.generateChartDataFromExcel();
      return {
        ...excelChartData,
        employeeStats: {}
      };
    }
    
    console.log('📊 Usando dados do Supabase');
    const [weeklyData, monthlyData, teamPerformance] = await Promise.all([
      this.getWeeklyChartData(),
      this.getMonthlyChartData(),
      this.getTeamPerformanceData()
    ]);

    return {
      weeklyData,
      monthlyData,
      teamPerformance,
      employeeStats: {}
    };
  }

  /**
   * Calcula estatísticas gerais
   * Agora suporta tanto Supabase quanto arquivos Excel
   */
  static async getGeneralStats(): Promise<GeneralStats> {
    if (this.useExcelMode) {
      console.log('📊 Calculando estatísticas dos arquivos Excel');
      const excelStats = await ExcelDataService.getGeneralStatsFromExcel();
      return excelStats || {
        bestPerformer: '',
        bestPoints: 0,
        avgTeam: 0,
        totalGoal: 29.5,
        progressPercentage: 0
      };
    }
    
    console.log('📊 Calculando estatísticas do Supabase');
    const employees = await this.getEmployees();
    const monthDates = CalculationsService.getMonthCycleDates();

    let bestPerformer = '';
    let bestPoints = 0;
    let totalPoints = 0;
    let totalPointsForAverage = 0;
    let employeeCountForAverage = 0;

    for (const employee of employees) {
      const points = await this.getEmployeePoints(
        employee.id,
        monthDates.start,
        monthDates.end
      );
      
      totalPoints += points;

      if (points > bestPoints) {
        bestPoints = points;
        bestPerformer = employee.real_name;
      }

      // Para média: excluir Rodrigo (freelancer)
      if (employee.real_name !== 'Rodrigo') {
        totalPointsForAverage += points;
        employeeCountForAverage++;
      }
    }

    const avgTeam = employeeCountForAverage > 0 ? 
      Math.round(totalPointsForAverage / employeeCountForAverage) : 0;
    
    const totalGoalTeam = 29500; // Meta mensal da equipe
    const progressPercentage = totalGoalTeam > 0 ? 
      (totalPoints / totalGoalTeam * 100) : 0;

    return {
      bestPerformer,
      bestPoints,
      avgTeam,
      totalGoal: Math.round(totalGoalTeam / 1000 * 10) / 10,
      progressPercentage: Math.round(progressPercentage * 10) / 10
    };
  }

  /**
   * Cor específica para cada funcionário
   */
  static getEmployeeColor(employeeName: string): string {
    return this.EMPLOYEE_COLORS[employeeName] || '#6b7280';
  }
}