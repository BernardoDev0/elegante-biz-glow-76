// Portando lógica do utils/calculations.py para TypeScript
// Mantendo a lógica original de semanas 26→25

export interface WeekDates {
  start: string;
  end: string;
}

export interface MonthCycleDates {
  start: string;
  end: string;
}

export class CalculationsService {
  
  /**
   * Obtém início do ciclo atual (dia 26)
   */
  static getCurrentCycleStart(): Date {
    const today = new Date();
    const currentDay = today.getDate();
    let cycleYear = today.getFullYear();
    let cycleMonth = today.getMonth() + 1;

    // Se já passou do dia 25, estamos no próximo ciclo
    if (currentDay >= 26) {
      // Estamos no próximo ciclo - manter mês atual
      cycleMonth = today.getMonth() + 1;
    } else {
      // Ainda no ciclo anterior - ir para mês anterior
      cycleMonth -= 1;
      if (cycleMonth < 1) {
        cycleMonth = 12;
        cycleYear -= 1;
      }
    }

    return new Date(cycleYear, cycleMonth - 1, 26);
  }

  // Lógica original: getWeekDates() do calculations.py
  static getWeekDates(weekStr: string): WeekDates {
    const weekNum = parseInt(weekStr);
    if (weekNum < 1 || weekNum > 5) {
      throw new Error('Semana deve estar entre 1 e 5');
    }

    const today = new Date();
    const currentMonth = today.getMonth() + 1; // getMonth() returns 0-11
    const currentYear = today.getFullYear();
    
    // Lógica baseada no ciclo 26→25
    let cycleYear = currentYear;
    let cycleMonth = currentMonth;

    // Se já passou do dia 25, estamos no próximo ciclo
    if (today.getDate() >= 26) {
      // Manter mês atual - já está no próximo ciclo
      cycleMonth = currentMonth;
    } else {
      // Ir para ciclo anterior
      cycleMonth -= 1;
      if (cycleMonth < 1) {
        cycleMonth = 12;
        cycleYear -= 1;
      }
    }

    // Calcular início do ciclo (dia 26)
    const cycleStart = new Date(cycleYear, cycleMonth - 1, 26);
    
    // Cada semana tem 7 dias, começando do dia 26
    const weekStartDate = new Date(cycleStart);
    weekStartDate.setDate(cycleStart.getDate() + (weekNum - 1) * 7);
    
    const weekEndDate = new Date(weekStartDate);
    weekEndDate.setDate(weekStartDate.getDate() + 6);

    return {
      start: weekStartDate.toISOString().split('T')[0],
      end: weekEndDate.toISOString().split('T')[0]
    };
  }

  // Obter semana atual baseada no ciclo 26→25 (LÓGICA CORRIGIDA)
  static getCurrentWeek(): number {
    const today = new Date();
    const currentDay = today.getDate();
    const currentMonth = today.getMonth() + 1;
    const currentYear = today.getFullYear();

    // Determinar início do ciclo atual (dia 26)
    let cycleYear = currentYear;
    let cycleMonth = currentMonth;

    // Se já passou do dia 25, estamos no próximo ciclo (igual ao DataService)
    if (currentDay >= 26) {
      // Estamos no próximo ciclo - manter mês atual
      cycleMonth = currentMonth;
    } else {
      // Ainda no ciclo anterior - ir para mês anterior
      cycleMonth -= 1;
      if (cycleMonth < 1) {
        cycleMonth = 12;
        cycleYear -= 1;
      }
    }

    // Início do ciclo sempre no dia 26
    const cycleStart = new Date(cycleYear, cycleMonth - 1, 26);
    const daysDiff = Math.floor((today.getTime() - cycleStart.getTime()) / (1000 * 60 * 60 * 24));
    
    // Calcular semana (1-5) - cada semana tem 7 dias
    const week = Math.floor(daysDiff / 7) + 1;
    return Math.min(Math.max(week, 1), 5);
  }

  // Obter datas do ciclo mensal (26→25) - LÓGICA PYTHON CORRIGIDA
  static getMonthCycleDates(month?: number, year?: number): MonthCycleDates {
    const today = new Date();
    let targetMonth = month || today.getMonth() + 1;
    let targetYear = year || today.getFullYear();

    // Se não especificado, pegar ciclo atual baseado no dia 26
    if (!month && !year) {
      // Se já passou do dia 25, estamos no próximo ciclo da empresa
      if (today.getDate() >= 26) {
        // Já no próximo ciclo - manter mês atual
        targetMonth = today.getMonth() + 1;
        targetYear = today.getFullYear();
      } else {
        // Ainda no ciclo anterior - ir para mês anterior  
        targetMonth = today.getMonth(); // já subtrai 1 pois getMonth() é 0-based
        if (targetMonth < 1) {
          targetMonth = 12;
          targetYear -= 1;
        }
      }
    }

    // Início do ciclo: dia 26 do mês especificado
    const cycleStart = new Date(targetYear, targetMonth - 1, 26);
    
    // Fim do ciclo: dia 25 do próximo mês  
    let endMonth = targetMonth + 1;
    let endYear = targetYear;
    if (endMonth > 12) {
      endMonth = 1;
      endYear += 1;
    }
    const cycleEnd = new Date(endYear, endMonth - 1, 25);

    return {
      start: cycleStart.toISOString().split('T')[0],
      end: cycleEnd.toISOString().split('T')[0]
    };
  }

  // Determinar semana de uma data específica (baseado no ciclo 26→25)
  static getWeekFromDate(dateStr: string): number {
    const date = new Date(dateStr);
    const day = date.getDate();
    const month = date.getMonth() + 1;
    const year = date.getFullYear();

    // Encontrar início do ciclo mensal para esta data (sempre dia 26)
    let cycleYear = year;
    let cycleMonth = month;

    // Se já passou do dia 25, pertence ao próximo ciclo
    if (day >= 26) {
      // Manter mês atual - já está no próximo ciclo
      cycleMonth = month;
    } else {
      // Ir para ciclo anterior
      cycleMonth -= 1;
      if (cycleMonth < 1) {
        cycleMonth = 12;
        cycleYear -= 1;
      }
    }

    const cycleStart = new Date(cycleYear, cycleMonth - 1, 26);
    const daysDiff = Math.floor((date.getTime() - cycleStart.getTime()) / (1000 * 60 * 60 * 24));
    
    // Calcular qual semana (1-5 dentro do ciclo mensal)
    const week = Math.floor(daysDiff / 7) + 1;
    return Math.min(Math.max(week, 1), 5);
  }

  // Semanas disponíveis (sempre 1-5)
  static getAvailableWeeks(): string[] {
    return ['1', '2', '3', '4', '5'];
  }

  // Calcular meta diária baseada no funcionário
  static getDailyGoal(employee: { username: string; weekly_goal?: number }): number {
    // Lógica original: Matheus (E89P) tem meta especial
    if (employee.username === 'E89P') {
      return 535; // 2675 / 5 dias
    }
    return 475; // 2375 / 5 dias (meta padrão)
  }

  // Calcular meta semanal baseada no funcionário
  static getWeeklyGoal(employee: { username: string; weekly_goal?: number }): number {
    if (employee.weekly_goal) {
      return employee.weekly_goal;
    }
    // Fallback para lógica original
    if (employee.username === 'E89P') {
      return 2675;
    }
    return 2375;
  }

  // Calcular meta mensal baseada no funcionário
  static getMonthlyGoal(employee: { username: string }): number {
    // Lógica original do employee.py
    if (employee.username === 'E89P') {
      return 10500;
    }
    return 9500;
  }

  // Calcular porcentagem de progresso
  static calculateProgressPercentage(current: number, goal: number): number {
    if (goal === 0) return 0;
    return Math.round((current / goal) * 100 * 10) / 10; // Rounded to 1 decimal
  }

  // Determinar status do funcionário (performance)
  static getEmployeeStatus(progressPercentage: number): 'at-risk' | 'on-track' | 'top-performer' {
    if (progressPercentage < 70) {
      return 'at-risk';
    } else if (progressPercentage > 110) {
      return 'top-performer';
    }
    return 'on-track';
  }

  // Formatar data para exibição (BR format)
  static formatDateBR(dateStr: string): string {
    const date = new Date(dateStr);
    return date.toLocaleDateString('pt-BR');
  }

  // Formatar timestamp para exibição (BR format)
  static formatTimestampBR(dateStr: string): { date: string; time: string } {
    const date = new Date(dateStr);
    return {
      date: date.toLocaleDateString('pt-BR'),
      time: date.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' })
    };
  }
}