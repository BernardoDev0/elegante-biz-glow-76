import { supabase } from '@/integrations/supabase/client';
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
   * Formato igual ao da imagem: Data, Refinaria, Pontos, Observações
   */
  static async exportEmployeeDataToZip(): Promise<void> {
    try {
      // Buscar todos os funcionários
      const { data: employees, error: employeesError } = await supabase
        .from('employee')
        .select('id, real_name')
        .order('real_name');

      if (employeesError || !employees) {
        throw new Error('Erro ao buscar funcionários');
      }

      // Buscar todas as entradas
      const { data: entries, error: entriesError } = await supabase
        .from('entry')
        .select(`
          id,
          date,
          points,
          observations,
          refinery,
          employee_id
        `)
        .order('date', { ascending: false })
        .limit(1000);

      if (entriesError || !entries) {
        throw new Error('Erro ao buscar registros');
      }

      // Criar um ZIP
      const zip = new JSZip();

      // Agrupar entradas por funcionário
      const entriesByEmployee = new Map<number, typeof entries>();
      entries.forEach(entry => {
        if (!entriesByEmployee.has(entry.employee_id)) {
          entriesByEmployee.set(entry.employee_id, []);
        }
        entriesByEmployee.get(entry.employee_id)?.push(entry);
      });

      // Criar arquivo Excel para cada funcionário
      for (const employee of employees) {
        const employeeEntries = entriesByEmployee.get(employee.id) || [];
        
        // Formatar dados para Excel (igual ao formato da imagem)
        let exportData: ExportEntry[];
        
        if (employeeEntries.length > 0) {
          exportData = employeeEntries.map(entry => ({
            Data: format(new Date(entry.date), 'dd/MM/yyyy HH:mm:ss', { locale: ptBR }),
            Refinaria: entry.refinery || '',
            Pontos: entry.points || 0,
            Observações: entry.observations || ''
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
        const fileName = `${employee.real_name} ${currentMonth}.xlsx`;
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
   * Busca dados reais dos funcionários para os gráficos
   * @deprecated Usar DataService.getChartData() em vez disso
   */
  static async getEmployeesChartData() {
    try {
      // Buscar funcionários
      const { data: employees, error: employeesError } = await supabase
        .from('employee')
        .select('*')
        .order('real_name');

      if (employeesError) {
        throw new Error('Erro ao buscar funcionários');
      }

      const chartData = {
        weeklyData: [],
        monthlyData: [],
        employeeStats: {},
        teamPerformance: []
      };

      if (!employees || employees.length === 0) {
        return chartData;
      }

      // Gerar dados semanais usando lógica do ciclo 26→25
      const weeklyData = [];
      
      // Determinar ciclo atual baseado no dia 26
      const today = new Date();
      const currentDay = today.getDate();
      let cycleYear = today.getFullYear();
      let cycleMonth = today.getMonth() + 1;

      if (currentDay >= 26) {
        // Já no próximo ciclo - manter mês atual
        cycleMonth = today.getMonth() + 1;
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
      
      // Gerar 5 semanas do ciclo
      for (let week = 1; week <= 5; week++) {
        const weekStart = new Date(cycleStart);
        weekStart.setDate(cycleStart.getDate() + (week - 1) * 7);
        
        const weekEnd = new Date(weekStart);
        weekEnd.setDate(weekStart.getDate() + 6);

        const weekData = { name: `Semana ${week}` };
        
        for (const employee of employees) {
          const { data: entries } = await supabase
            .from('entry')
            .select('points')
            .eq('employee_id', employee.id)
            .gte('date', weekStart.toISOString().split('T')[0])
            .lte('date', weekEnd.toISOString().split('T')[0]);

          const points = entries?.reduce((sum, entry) => sum + entry.points, 0) || 0;
          weekData[employee.real_name] = points;
        }
        
        weeklyData.push(weekData);
      }

      chartData.weeklyData = weeklyData;

      // Gerar dados mensais usando MESMA LÓGICA do DataService
      const monthlyData = [];
      const months = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];
      
      const currentDateMonthly = new Date();
      const currentDayMonthly = currentDateMonthly.getDate();
      const currentCalendarMonth = currentDateMonthly.getMonth(); // 0-based (0=Janeiro)
      const currentYear = currentDateMonthly.getFullYear();
      
      // Determinar qual é o "mês atual da empresa" baseado na regra 26-25
      let currentCompanyMonth = currentCalendarMonth;
      
      // Se já passou do dia 25, estamos no próximo mês da empresa
      if (currentDayMonthly >= 26) {
        // Estamos no próximo ciclo - manter mês atual
        currentCompanyMonth = currentCalendarMonth;
      } else {
        // Ainda no ciclo anterior - ir para mês anterior
        currentCompanyMonth = currentCalendarMonth - 1;
        if (currentCompanyMonth < 0) {
          currentCompanyMonth = 11; // Dezembro do ano anterior
        }
      }
      
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
        // Usar ciclo 26-25 para o mês
        const monthStart = new Date(yearForMonth, monthIndex, 26);
        let monthEnd = new Date(yearForMonth, monthIndex + 1, 25);
        
        // Se monthIndex + 1 > 11, ajustar para próximo ano
        if (monthIndex === 11) {
          monthEnd = new Date(yearForMonth + 1, 0, 25);
        }
        
        // Total de pontos do mês com dados por funcionário
        const monthData = { name: months[monthIndex] };
        
        for (const employee of employees) {
          const { data: employeeEntries } = await supabase
            .from('entry')
            .select('points')
            .eq('employee_id', employee.id)
            .gte('date', monthStart.toISOString().split('T')[0])
            .lte('date', monthEnd.toISOString().split('T')[0]);

          const points = employeeEntries?.reduce((sum, entry) => sum + entry.points, 0) || 0;
          monthData[employee.real_name] = points;
        }
        
        monthlyData.push(monthData);
      }

      chartData.monthlyData = monthlyData;

      // Para cada funcionário, calcular estatísticas
      for (const employee of employees) {
        // Dados do mês atual (ciclo 26-25)
        const today = new Date();
        const currentDay = today.getDate();
        let monthStart: Date;
        let monthEnd: Date;

        if (currentDay >= 26) {
          monthStart = new Date(today.getFullYear(), today.getMonth(), 26);
          monthEnd = new Date(today.getFullYear(), today.getMonth() + 1, 25);
        } else {
          monthStart = new Date(today.getFullYear(), today.getMonth() - 1, 26);
          monthEnd = new Date(today.getFullYear(), today.getMonth(), 25);
        }

        const { data: monthlyEntries } = await supabase
          .from('entry')
          .select('points')
          .eq('employee_id', employee.id)
          .gte('date', monthStart.toISOString().split('T')[0])
          .lte('date', monthEnd.toISOString().split('T')[0]);

        const monthlyPoints = monthlyEntries?.reduce((sum, entry) => sum + entry.points, 0) || 0;

        // Dados para gráfico de pizza (team performance)
        chartData.teamPerformance.push({
          name: employee.real_name,
          value: monthlyPoints,
          color: ExportService.getEmployeeColor(employee.real_name)
        });
      }

      return chartData;
      
    } catch (error) {
      console.error('Erro ao buscar dados dos gráficos:', error);
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
   * IMPORTANTE: Exclui Rodrigo das médias (funcionário freelancer)
   * @deprecated Usar DataService.getGeneralStats() em vez disso
   */
  static async getGeneralStats() {
    try {
      const { data: employees } = await supabase
        .from('employee')
        .select('id, real_name, weekly_goal');

      if (!employees) return null;

      // Dados do mês atual (ciclo 26-25)
      const today = new Date();
      const currentDay = today.getDate();
      let monthStart: Date;
      let monthEnd: Date;

      if (currentDay >= 26) {
        monthStart = new Date(today.getFullYear(), today.getMonth(), 26);
        monthEnd = new Date(today.getFullYear(), today.getMonth() + 1, 25);
      } else {
        monthStart = new Date(today.getFullYear(), today.getMonth() - 1, 26);
        monthEnd = new Date(today.getFullYear(), today.getMonth(), 25);
      }

      let bestPerformer = '';
      let bestPoints = 0;
      let totalPoints = 0;
      let totalGoalTeam = 29500; // Meta mensal da equipe definida pelo usuário
      
      // Para média da equipe: excluir Rodrigo (funcionário freelancer)
      let totalPointsForAverage = 0;
      let employeeCountForAverage = 0;

      for (const employee of employees) {
        const { data: entries } = await supabase
          .from('entry')
          .select('points')
          .eq('employee_id', employee.id)
          .gte('date', monthStart.toISOString().split('T')[0])
          .lte('date', monthEnd.toISOString().split('T')[0]);

        const points = entries?.reduce((sum, entry) => sum + entry.points, 0) || 0;
        totalPoints += points;

        // Verificar melhor performer
        if (points > bestPoints) {
          bestPoints = points;
          bestPerformer = employee.real_name;
        }

        // Para média: excluir Rodrigo
        if (employee.real_name !== 'Rodrigo') {
          totalPointsForAverage += points;
          employeeCountForAverage++;
        }
      }

      // Média da equipe sem Rodrigo
      const avgTeam = employeeCountForAverage > 0 ? Math.round(totalPointsForAverage / employeeCountForAverage) : 0;
      
      // Progresso geral baseado na meta de 29.500
      const progressPercentage = totalGoalTeam > 0 ? (totalPoints / totalGoalTeam * 100) : 0;

      return {
        bestPerformer,
        bestPoints,
        avgTeam, // Média sem Rodrigo
        totalGoal: Math.round(totalGoalTeam / 1000 * 10) / 10, // 29.5K
        progressPercentage: Math.round(progressPercentage * 10) / 10
      };
      
    } catch (error) {
      console.error('Erro ao calcular estatísticas:', error);
      return null;
    }
  }
}