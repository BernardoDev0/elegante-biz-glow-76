import React, { useState, useEffect, useRef, useCallback } from 'react';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { ChartContainer, ChartTooltip, ChartTooltipContent } from '@/components/ui/chart';
import { Bar, BarChart, ResponsiveContainer, XAxis, YAxis, CartesianGrid, Legend } from 'recharts';
import { useToast } from '@/hooks/use-toast';
import { Loader2, FolderOpen, FileSpreadsheet, Users, TrendingUp, DollarSign } from 'lucide-react';

interface ExcelData {
  employees: Record<string, EmployeeData>;
  statistics?: Statistics;
}

interface EmployeeData {
  total_points?: number;
  totalPoints?: number;
  records?: Array<{
    date: string;
    points: number;
    refinery?: string;
    observations?: string;
    month?: string;
  }>;
  months?: Record<string, {
    records: number;
    points: number;
    total?: number;
  }>;
}

interface Statistics {
  total_files?: number;
  totalFiles?: number;
  total_records?: number;
  totalRecords?: number;
  total_points?: number;
  totalPoints?: number;
}

interface ChartDataPoint {
  month: string;
  [employeeName: string]: string | number;
}

interface EmployeeColors {
  borderColor: string;
  backgroundColor: string;
}

const excelChartConfig = {
  Rodrigo: {
    label: "Rodrigo",
    color: "rgba(168, 85, 247, 1)", // Roxo vivo
  },
  Maurício: {
    label: "Maurício", 
    color: "rgba(59, 130, 246, 1)", // Azul vivo
  },
  Matheus: {
    label: "Matheus",
    color: "rgba(34, 197, 94, 1)", // Verde vivo
  },
  Wesley: {
    label: "Wesley",
    color: "rgba(239, 68, 68, 1)", // Vermelho vivo
  },
};

export default function ExcelDashboard() {
  const [isLoading, setIsLoading] = useState(false);
  const [data, setData] = useState<ExcelData | null>(null);
  const [statistics, setStatistics] = useState<Statistics | null>(null);
  const { toast } = useToast();
  const animationFrameRef = useRef<number>();

  // Carregar pasta de arquivos Excel
  const loadFolder = async () => {
    if (isLoading) return;

    console.log('🔍 === DEBUG: INICIANDO LOAD FOLDER ===');
    setIsLoading(true);
    
    try {
      console.log('🔍 Fazendo requisição para /api/excel/load_folder');
      
      const response = await fetch('/api/excel/load_folder', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          folder_path: 'registros monitorar'
        })
      });

      console.log('📡 Resposta recebida:', response.status);

      const responseData = await response.json();
      console.log('📊 Dados recebidos:', responseData);

      if (responseData.status === 'success') {
        setData(responseData.data || {});
        setStatistics(responseData.statistics || {});
        
        toast({
          title: "Sucesso",
          description: "Arquivos carregados com sucesso!",
        });
      } else {
        throw new Error(responseData.message || 'Erro ao carregar arquivos');
      }
    } catch (error) {
      console.error('Excel Dashboard: Erro ao carregar pasta:', error);
      toast({
        title: "Erro",
        description: `Erro: ${error instanceof Error ? error.message : 'Erro desconhecido'}`,
        variant: "destructive",
      });
    } finally {
      setIsLoading(false);
    }
  };

  // Calcular lucro total
  const calculateProfit = (totalPoints: number): string => {
    const profit = totalPoints * 3.25;
    return profit.toLocaleString('pt-BR', { 
      style: 'currency', 
      currency: 'BRL' 
    });
  };

  // Animar valores com otimização
  const animateValue = useCallback((
    element: HTMLElement, 
    start: number, 
    end: number | string, 
    duration: number
  ) => {
    // Se o valor final é uma string (como "R$ 1.234,56"), não animar
    if (typeof end === 'string') {
      element.textContent = end;
      return;
    }

    const startTime = performance.now();
    const numericEnd = Number(end) || 0;

    const animate = (currentTime: number) => {
      const elapsed = currentTime - startTime;
      const progress = Math.min(elapsed / duration, 1);
      
      // Easing suave
      const easeOut = 1 - Math.pow(2, -10 * progress);
      let current = start + (numericEnd - start) * easeOut;
      
      if (progress >= 1) {
        current = numericEnd;
      } else {
        current = Math.round(current);
      }
      
      element.textContent = current.toLocaleString('pt-BR');

      if (progress < 1) {
        animationFrameRef.current = requestAnimationFrame(animate);
      }
    };

    animationFrameRef.current = requestAnimationFrame(animate);
  }, []);

  // Obter mês a partir de uma data (considerando mês do dia 26 ao 25)
  const getMonthFromDate = (dateString: string | Date): string | null => {
    console.log('🔍 Processando data:', dateString);
    
    try {
      let date: Date;
      
      // Se é uma string, tentar parsear
      if (typeof dateString === 'string') {
        date = new Date(dateString);
      } else if (dateString instanceof Date) {
        date = dateString;
      } else {
        console.log('  ❌ Formato de data inválido:', typeof dateString);
        return null;
      }
      
      console.log('  📅 Data parseada:', date);
      
      if (isNaN(date.getTime())) {
        console.log('  ❌ Data inválida');
        return null;
      }
      
      // LÓGICA CORRETA: Se dia >= 26, pertence ao mês seguinte
      const day = date.getDate();
      let month = date.getMonth();
      let year = date.getFullYear();
      
      let target_month: number;
      let target_year: number;

      if (day >= 26) {
        // Mês seguinte
        if (month === 11) { // Dezembro
          target_month = 0; // Janeiro
          target_year = year + 1;
        } else {
          target_month = month + 1;
          target_year = year;
        }
      } else {
        // Mês atual
        target_month = month;
        target_year = year;
      }
      
      const monthNames = [
        'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
        'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'
      ];
      
      const monthName = monthNames[target_month];
      console.log(`  📅 Mês calculado: ${monthName} (dia ${day})`);
      
      return monthName;
    } catch (error) {
      console.error('❌ Erro ao processar data:', error);
      return null;
    }
  };

  // Obter cores para cada funcionário (igual ao da aba Gráficos)
  const getEmployeeColors = (employeeName: string): EmployeeColors => {
    const colorMap: Record<string, EmployeeColors> = {
      'Matheus': {
        borderColor: 'rgba(34, 197, 94, 1)',    // Verde vivo
        backgroundColor: 'rgba(34, 197, 94, 0.8)'
      },
      'Maurício': {
        borderColor: 'rgba(59, 130, 246, 1)',   // Azul vivo
        backgroundColor: 'rgba(59, 130, 246, 0.8)'
      },
      'Rodrigo': {
        borderColor: 'rgba(168, 85, 247, 1)',   // Roxo vivo
        backgroundColor: 'rgba(168, 85, 247, 0.8)'
      },
      'Wesley': {
        borderColor: 'rgba(239, 68, 68, 1)',    // Vermelho vivo
        backgroundColor: 'rgba(239, 68, 68, 0.8)'
      }
    };
    
    return colorMap[employeeName] || {
      borderColor: 'rgba(107, 114, 128, 1)',      // Cinza escuro
      backgroundColor: 'rgba(107, 114, 128, 0.8)'
    };
  };

  // Converter mês para formato limpo
  const convertToCleanMonth = (month: string): string => {
    const monthMap: Record<string, string> = {
      '01': 'Janeiro',
      '02': 'Fevereiro', 
      '03': 'Março',
      '04': 'Abril',
      '05': 'Maio',
      '06': 'Junho',
      '07': 'Julho',
      '08': 'Agosto',
      '09': 'Setembro',
      '10': 'Outubro',
      '11': 'Novembro',
      '12': 'Dezembro'
    };
    
    // Se é formato "MM/YYYY", converter para nome do mês
    if (month.includes('/')) {
      const [monthNum, year] = month.split('/');
      const monthName = monthMap[monthNum] || month;
      return `${monthName} ${year}`;
    }
    
    // Se já é nome do mês, verificar se tem ano
    if (month.includes(' ')) {
      return month; // Já está no formato correto
    }
    
    // Se é só nome do mês sem ano, adicionar ano atual
    const currentYear = new Date().getFullYear();
    return `${month} ${currentYear}`;
  };

  // Extrair meses dos registros individuais
  const extractMonthsFromRecords = (records: any, monthsSet: Set<string>): void => {
    console.log('🔍 Extraindo meses dos registros:', records);
    
    // Verificar se records é um array ou um número
    if (typeof records === 'number') {
      console.log('❌ Records é um número, não um array. Tentando buscar registros individuais...');
      return;
    }
    
    if (!Array.isArray(records)) {
      console.log('❌ Records não é um array:', typeof records);
      return;
    }
    
    console.log(`📋 Processando ${records.length} registros individuais`);
    
    records.forEach((record, index) => {
      console.log(`  📋 Registro ${index + 1}:`, record);
      
      if (record.date) {
        const month = getMonthFromDate(record.date);
        if (month) {
          monthsSet.add(month);
          console.log(`    📅 Mês extraído: ${month}`);
        }
      }
    });
  };

  // Extrair meses de todos os registros
  const extractMonthsFromAllRecords = (monthsSet: Set<string>): void => {
    console.log('🔍 Extraindo meses de todos os registros...');
    
    if (data?.employees) {
      Object.values(data.employees).forEach(empData => {
        if (empData.records) {
          extractMonthsFromRecords(empData.records, monthsSet);
        }
      });
    }
  };

  // Obter meses únicos dos dados do Excel
  const getMonthsFromExcelData = (): string[] => {
    console.log('🔍 === DEBUG: EXTRAINDO MESES DOS DADOS ===');
    
    const months = new Set<string>();
    const processedRecords = new Set<string>(); // Para evitar duplicação
    
    if (data?.employees) {
      console.log('📊 Processando dados dos funcionários...');
      
      Object.values(data.employees).forEach((empData, index) => {
        console.log(`\n👤 Funcionário ${index + 1}:`, empData);
        
        // Extrair meses APENAS dos registros individuais
        if (empData.records && Array.isArray(empData.records)) {
          console.log(`  📋 ${empData.records.length} registros individuais encontrados`);
          
          empData.records.forEach((record, recordIndex) => {
            console.log(`    📋 Registro ${recordIndex + 1}:`, record);
            
            if (record.date) {
              // Verificar se já foi processado
              const recordKey = `${record.date}_${record.points}`;
              if (processedRecords.has(recordKey)) {
                console.log(`    ⚠️ Registro já processado: ${recordKey}`);
                return; // Pular registro duplicado
              }
              processedRecords.add(recordKey);
              
              const month = getMonthFromDate(record.date);
              console.log(`    📅 Mês extraído: ${month}`);
              if (month) months.add(month);
            }
          });
        } else {
          console.log('  ❌ Nenhum registro individual encontrado');
        }
      });
    }
    
    console.log('📅 Meses únicos encontrados:', Array.from(months));
    
    // Se não encontrou meses específicos, tentar extrair dos registros
    if (months.size === 0) {
      console.log('🔍 Tentando extrair meses dos registros...');
      extractMonthsFromAllRecords(months);
    }
    
    // Converter meses para formato mais limpo
    const cleanMonths = cleanMonthNames(Array.from(months));
    console.log('📅 Meses limpos:', cleanMonths);
    
    return cleanMonths;
  };

  // Limpar nomes dos meses para formato mais legível
  const cleanMonthNames = (months: string[]): string[] => {
    const monthMap: Record<string, string> = {
      '01': 'Janeiro',
      '02': 'Fevereiro', 
      '03': 'Março',
      '04': 'Abril',
      '05': 'Maio',
      '06': 'Junho',
      '07': 'Julho',
      '08': 'Agosto',
      '09': 'Setembro',
      '10': 'Outubro',
      '11': 'Novembro',
      '12': 'Dezembro'
    };
    
    // Converter meses para formato limpo e remover duplicatas
    const cleanMonths = new Set<string>();
    
    months.forEach(month => {
      let cleanMonth: string;
      
      // Se é formato "MM/YYYY", converter para nome do mês
      if (month.includes('/')) {
        const [monthNum, year] = month.split('/');
        const monthName = monthMap[monthNum] || month;
        cleanMonth = `${monthName} ${year}`;
      } else {
        // Se já é nome do mês, verificar se tem ano
        if (month.includes(' ')) {
          cleanMonth = month; // Já está no formato correto
        } else {
          // Se é só nome do mês sem ano, adicionar ano atual
          const currentYear = new Date().getFullYear();
          cleanMonth = `${month} ${currentYear}`;
        }
      }
      
      cleanMonths.add(cleanMonth);
    });
    
    // Ordenar meses cronologicamente
    const monthOrder = [
      'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
      'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'
    ];
    
    const sortedMonths = Array.from(cleanMonths).sort((a, b) => {
      const aParts = a.split(' ');
      const bParts = b.split(' ');
      
      const aMonth = aParts[0];
      const bMonth = bParts[0];
      const aYear = parseInt(aParts[1]) || 0;
      const bYear = parseInt(bParts[1]) || 0;
      
      // Primeiro ordenar por ano
      if (aYear !== bYear) {
        return aYear - bYear;
      }

      // Depois por mês
      return monthOrder.indexOf(aMonth) - monthOrder.indexOf(bMonth);
    });
    
    console.log('📅 Meses limpos e ordenados:', sortedMonths);
    return sortedMonths;
  };

  // Processar registros de um funcionário para criar dados por mês
  const processEmployeeRecords = (empData: EmployeeData, months: string[]): number[] => {
    console.log('🔍 === DEBUG: PROCESSANDO REGISTROS ===');
    console.log('Funcionário:', empData);
    console.log('Meses:', months);
    
    // Inicializar dados por mês
    const monthlyPoints: Record<string, number> = {};
    months.forEach(month => {
      monthlyPoints[month] = 0;
    });
    
    console.log('📊 Inicialização monthlyPoints:', monthlyPoints);
    
    // Calcular meses APENAS a partir dos registros individuais
    if (empData.records && Array.isArray(empData.records)) {
      console.log(`📋 Calculando a partir de ${empData.records.length} registros individuais...`);
      
      // Set para evitar duplicação
      const processedRecords = new Set<string>();
      
      // Debug para cada registro
      empData.records.forEach((record, index) => {
        console.log(`\n📋 === REGISTRO ${index + 1} ===`);
        console.log('Registro:', record);
        
        if (record.date && record.points) {
          console.log(`  📅 Data: ${record.date}, Pontos: ${record.points}`);
          
          // Verificar se já foi processado
          const recordKey = `${record.date}_${record.points}`;
          console.log(`  🔑 Chave do registro: ${recordKey}`);
          
          if (processedRecords.has(recordKey)) {
            console.log(`  ⚠️ Registro já processado: ${recordKey}`);
            return; // Pular registro duplicado
          }
          processedRecords.add(recordKey);
          
          const month = getMonthFromDate(record.date);
          console.log(`  📅 Mês calculado: ${month}`);
          
          if (month) {
            // Converter para formato limpo
            const cleanMonth = convertToCleanMonth(month);
            console.log(`  📅 Mês limpo: ${cleanMonth}`);
            console.log(`  📅 MonthlyPoints antes:`, monthlyPoints);
            
            if (monthlyPoints.hasOwnProperty(cleanMonth)) {
              const oldValue = monthlyPoints[cleanMonth];
              monthlyPoints[cleanMonth] += record.points;
              console.log(`  ✅ Adicionado ${record.points} pontos para ${cleanMonth}`);
              console.log(`  📊 ${oldValue} + ${record.points} = ${monthlyPoints[cleanMonth]}`);
            } else {
              console.log(`  ❌ Mês não encontrado: ${cleanMonth}`);
              console.log(`  📊 Meses disponíveis:`, Object.keys(monthlyPoints));
            }
            
            console.log(`  📅 MonthlyPoints depois:`, monthlyPoints);
          } else {
            console.log(`  ❌ Mês inválido para data: ${record.date}`);
          }
        } else {
          console.log(`  ❌ Dados inválidos: date=${record.date}, points=${record.points}`);
        }
      });
    } else {
      // Fallback: distribuir pontos totais pelos meses
      console.log('📊 Distribuindo pontos totais pelos meses');
      const totalPoints = empData.total_points || empData.totalPoints || 0;
      const pointsPerMonth = Math.round(totalPoints / months.length);
      
      months.forEach(month => {
        monthlyPoints[month] = pointsPerMonth;
      });
    }
    
    // Converter para array na ordem dos meses
    const result = months.map(month => monthlyPoints[month] || 0);
    console.log('📊 Resultado final:', result);
    console.log('📊 Soma total:', result.reduce((a, b) => a + b, 0));
    
    return result;
  };

  // Preparar dados do gráfico
  const prepareChartData = (): ChartDataPoint[] => {
    console.log('🔍 === DEBUG: PREPARANDO DADOS DO GRÁFICO ===');
    console.log('State data:', data);
    
    if (!data?.employees) {
      console.log('❌ Dados ou employees não encontrados');
      return [];
    }

    const employees = Object.keys(data.employees);
    console.log('👥 Funcionários encontrados:', employees);
    
    // Obter meses únicos dos dados do Excel
    const months = getMonthsFromExcelData();
    console.log('📅 Meses extraídos:', months);
    
    // Preparar datasets para cada funcionário
    return months.map(month => {
      const dataPoint: ChartDataPoint = { month };
      
      employees.forEach(employeeName => {
        console.log(`\n🔍 Processando funcionário: ${employeeName}`);
        const empData = data.employees[employeeName] || {};
        console.log('Dados do funcionário:', empData);
        
        // Processar registros individuais para criar dados por mês
        const monthlyData = processEmployeeRecords(empData, [month]);
        console.log(`📊 Dados mensais para ${employeeName}:`, monthlyData);
        
        dataPoint[employeeName] = monthlyData[0] || 0;
      });
      
      return dataPoint;
    });
  };

  // Funções de debug (equivalentes ao JavaScript original)
  const debugDataStructure = () => {
    console.log('🔍 === DEBUG MASSIVO: ESTRUTURA DOS DADOS ===');
    console.log('State completo:', { data, statistics });
    
    if (data) {
      console.log('📊 Dados encontrados:', data);
      
      if (data.employees) {
        console.log('👥 Estrutura dos funcionários:');
        Object.entries(data.employees).forEach(([name, empData]) => {
          console.log(`\n👤 ${name}:`, empData);
          
          if (empData.records) {
            console.log(`  📋 ${empData.records.length} registros encontrados`);
            if (Array.isArray(empData.records)) {
              empData.records.slice(0, 3).forEach((record, index) => {
                console.log(`    Registro ${index + 1}:`, record);
              });
            } else {
              console.log(`    ❌ Records não é array: ${typeof empData.records}`);
            }
          }
          
          if (empData.months) {
            console.log(`  📅 Meses:`, Object.keys(empData.months));
            Object.entries(empData.months).forEach(([month, monthData]) => {
              console.log(`    ${month}:`, monthData);
            });
          }
        });
      } else {
        console.log('❌ Nenhum funcionário encontrado');
      }
    } else {
      console.log('❌ Nenhum dado encontrado');
    }
  };

  const debugDuplication = () => {
    console.log('🔍 === DEBUG DUPLICAÇÃO ===');
    
    if (!data?.employees) {
      console.log('❌ Nenhum dado para analisar');
      return;
    }
    
    Object.entries(data.employees).forEach(([employeeName, empData]) => {
      console.log(`\n👤 === ANÁLISE ${employeeName} ===`);
      
      // Verificar registros individuais
      if (empData.records && Array.isArray(empData.records)) {
        console.log(`📋 ${empData.records.length} registros individuais:`);
        
        // Agrupar por mês para verificar duplicação
        const recordsByMonth: Record<string, any[]> = {};
        empData.records.forEach((record, index) => {
          const month = record.month || getMonthFromDate(record.date);
          if (!recordsByMonth[month || '']) {
            recordsByMonth[month || ''] = [];
          }
          recordsByMonth[month || ''].push({
            index: index,
            date: record.date,
            points: record.points,
            month: month
          });
        });
        
        console.log('📅 Registros agrupados por mês:');
        Object.entries(recordsByMonth).forEach(([month, records]) => {
          const totalPoints = records.reduce((sum, r) => sum + r.points, 0);
          console.log(`  ${month}: ${records.length} registros, ${totalPoints} pontos totais`);
          records.forEach(record => {
            console.log(`    - Registro ${record.index}: ${record.date} = ${record.points} pontos`);
          });
        });
      }
      
      // Verificar dados por mês (se existirem)
      if (empData.months) {
        console.log('📅 Dados por mês (backend):');
        Object.entries(empData.months).forEach(([month, monthData]) => {
          console.log(`  ${month}: ${monthData.points || monthData.total || 0} pontos`);
        });
      }
    });
  };

  const debugMonthProcessing = () => {
    console.log('🔍 === DEBUG PROCESSAMENTO DE MESES ===');
    
    const months = getMonthsFromExcelData();
    console.log('📅 Meses extraídos:', months);
    
    if (data?.employees) {
      Object.entries(data.employees).forEach(([employeeName, empData]) => {
        console.log(`\n👤 === PROCESSAMENTO ${employeeName} ===`);
        
        const monthlyData = processEmployeeRecords(empData, months);
        console.log(`📊 Dados mensais para ${employeeName}:`, monthlyData);
        
        // Verificar se há valores duplicados
        const nonZeroValues = monthlyData.filter(val => val > 0);
        console.log(`📈 Valores não-zero:`, nonZeroValues);
        
        if (nonZeroValues.length > 0) {
          const sum = nonZeroValues.reduce((a, b) => a + b, 0);
          console.log(`📊 Soma total: ${sum}`);
        }
      });
    }
  };

  const showMessage = (message: string, type: 'info' | 'success' | 'error' = 'info') => {
    toast({
      title: type === 'success' ? "Sucesso" : type === 'error' ? "Erro" : "Informação",
      description: message,
      variant: type === 'error' ? "destructive" : "default",
    });
  };

  // Expor funções de debug no window (para uso no console)
  useEffect(() => {
    if (typeof window !== 'undefined') {
      (window as any).excelDebug = {
        debugDataStructure,
        debugDuplication,
        debugMonthProcessing,
        loadFolder,
        showMessage
      };
      
      console.log('🔧 === EXCEL DASHBOARD DEBUG ===');
      console.log('Comandos disponíveis no console:');
      console.log('  window.excelDebug.debugDataStructure() - Analisar estrutura dos dados');
      console.log('  window.excelDebug.debugDuplication() - Debug específico para duplicação');
      console.log('  window.excelDebug.debugMonthProcessing() - Debug para processamento de meses');
      console.log('  window.excelDebug.loadFolder() - Carregar pasta Excel (DADOS REAIS)');
    }
    
    return () => {
      if (typeof window !== 'undefined') {
        delete (window as any).excelDebug;
      }
    };
  }, [data, statistics]);

  // Cleanup na desmontagem
  useEffect(() => {
    return () => {
      if (animationFrameRef.current) {
        cancelAnimationFrame(animationFrameRef.current);
      }
    };
  }, []);

  const chartData = prepareChartData();
  const totalPoints = statistics?.total_points || statistics?.totalPoints || 0;

  return (
    <div className="space-y-6">
      {/* Load Button */}
      <Card>
        <CardHeader>
          <CardTitle className="flex items-center gap-2">
            <FolderOpen className="h-5 w-5" />
            Carregar Arquivos Excel
          </CardTitle>
        </CardHeader>
        <CardContent>
          <Button 
            onClick={loadFolder} 
            disabled={isLoading}
            className="w-full sm:w-auto"
          >
            {isLoading ? (
              <>
                <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                Carregando...
              </>
            ) : (
              <>
                <FolderOpen className="mr-2 h-4 w-4" />
                Carregar Pasta
              </>
            )}
          </Button>
        </CardContent>
      </Card>

      {/* Statistics Grid */}
      {statistics && (
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
          <Card>
            <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2">
              <CardTitle className="text-sm font-medium">Total de Arquivos</CardTitle>
              <FileSpreadsheet className="h-4 w-4 text-muted-foreground" />
            </CardHeader>
            <CardContent>
              <div className="text-2xl font-bold">
                {(statistics.total_files || statistics.totalFiles || 0).toLocaleString('pt-BR')}
              </div>
              <p className="text-xs text-muted-foreground">
                {(statistics.total_files || statistics.totalFiles || 0)} arquivos processados
              </p>
            </CardContent>
          </Card>

          <Card>
            <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2">
              <CardTitle className="text-sm font-medium">Total de Registros</CardTitle>
              <Users className="h-4 w-4 text-muted-foreground" />
            </CardHeader>
            <CardContent>
              <div className="text-2xl font-bold">
                {(statistics.total_records || statistics.totalRecords || 0).toLocaleString('pt-BR')}
              </div>
              <p className="text-xs text-muted-foreground">
                {(statistics.total_records || statistics.totalRecords || 0).toLocaleString('pt-BR')} registros totais
              </p>
            </CardContent>
          </Card>

          <Card>
            <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2">
              <CardTitle className="text-sm font-medium">Total de Pontos</CardTitle>
              <TrendingUp className="h-4 w-4 text-muted-foreground" />
            </CardHeader>
            <CardContent>
              <div className="text-2xl font-bold">
                {totalPoints.toLocaleString('pt-BR')}
              </div>
              <p className="text-xs text-muted-foreground">
                {totalPoints.toLocaleString('pt-BR')} pontos totais
              </p>
            </CardContent>
          </Card>

          <Card>
            <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2">
              <CardTitle className="text-sm font-medium">Lucro Total</CardTitle>
              <DollarSign className="h-4 w-4 text-muted-foreground" />
            </CardHeader>
            <CardContent>
              <div className="text-2xl font-bold">
                {calculateProfit(totalPoints)}
              </div>
              <p className="text-xs text-muted-foreground">
                {calculateProfit(totalPoints)} (pontos x R$ 3,25)
              </p>
            </CardContent>
          </Card>
        </div>
      )}

      {/* Chart */}
      {chartData.length > 0 && (
        <Card>
          <CardHeader>
            <CardTitle>Evolução dos Funcionários</CardTitle>
          </CardHeader>
          <CardContent>
            <ChartContainer config={excelChartConfig} className="h-[400px]">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={chartData}>
                  <CartesianGrid strokeDasharray="3 3" stroke="hsl(var(--border))" />
                  <XAxis 
                    dataKey="month" 
                    stroke="hsl(var(--muted-foreground))"
                    fontSize={12}
                  />
                  <YAxis 
                    stroke="hsl(var(--muted-foreground))"
                    fontSize={12}
                  />
                  <ChartTooltip 
                    content={<ChartTooltipContent />}
                    formatter={(value, name) => [
                      `${Number(value).toLocaleString('pt-BR')} pontos`,
                      name
                    ]}
                    labelFormatter={(label) => `📅 ${label}`}
                  />
                  <Legend />
                  {Object.keys(excelChartConfig).map((employeeName, index) => (
                    <Bar
                      key={employeeName}
                      dataKey={employeeName}
                      fill={`var(--color-${employeeName})`}
                      radius={[4, 4, 0, 0]}
                    />
                  ))}
                </BarChart>
              </ResponsiveContainer>
            </ChartContainer>
          </CardContent>
        </Card>
      )}

      {/* Loading State */}
      {isLoading && (
        <Card>
          <CardContent className="flex items-center justify-center py-8">
            <div className="flex items-center space-x-2">
              <Loader2 className="h-6 w-6 animate-spin" />
              <span>Carregando arquivos Excel...</span>
            </div>
          </CardContent>
        </Card>
      )}
    </div>
  );
}