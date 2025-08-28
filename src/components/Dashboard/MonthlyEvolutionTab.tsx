import { useState, useEffect } from "react";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Legend } from 'recharts';
import { TrendingUp } from "lucide-react";
import { EmployeeService, Entry } from "@/services/EmployeeService";
import { CalculationsService } from "@/services/CalculationsService";
import { format, startOfMonth, endOfMonth, eachWeekOfInterval, startOfWeek, endOfWeek, parseISO } from "date-fns";
import { ptBR } from "date-fns/locale";

interface MonthlyEvolutionTabProps {
  employeeId: number;
  monthlyGoal: number;
}

interface WeeklyData {
  week: string;
  weekLabel: string;
  points: number;
  goal: number;
}

export const MonthlyEvolutionTab = ({ employeeId, monthlyGoal }: MonthlyEvolutionTabProps) => {
  const [entries, setEntries] = useState<Entry[]>([]);
  const [weeklyData, setWeeklyData] = useState<WeeklyData[]>([]);
  const [loading, setLoading] = useState(true);

  const loadMonthlyData = async () => {
    try {
      setLoading(true);
      
      // Buscar dados do mês atual
      const monthDates = CalculationsService.getMonthCycleDates();
      const monthEntries = await EmployeeService.getEmployeeEntries(
        employeeId,
        100, // Limite alto para pegar todos os dados do mês
        0,
        monthDates
      );
      
      setEntries(monthEntries);
      
      // Calcular dados por semana usando o sistema 26→25
      const weeklyGoal = monthlyGoal / 5; // Sempre 5 semanas
      
      const chartData: WeeklyData[] = [];
      
      // Iterar pelas 5 semanas do sistema
      for (let weekNum = 1; weekNum <= 5; weekNum++) {
        try {
          const weekDates = CalculationsService.getWeekDates(weekNum.toString());
          const weekEntries = monthEntries.filter(entry => {
            const entryDate = entry.date.split('T')[0]; // Pegar apenas a data
            return entryDate >= weekDates.start && entryDate <= weekDates.end;
          });
          
          const weekPoints = weekEntries.reduce((sum, entry) => sum + entry.points, 0);
          
          chartData.push({
            week: `week-${weekNum}`,
            weekLabel: `Semana ${weekNum}`,
            points: weekPoints,
            goal: Math.round(weeklyGoal)
          });
        } catch (error) {
          // Se der erro ao calcular uma semana, adicionar dados zerados
          chartData.push({
            week: `week-${weekNum}`,
            weekLabel: `Semana ${weekNum}`,
            points: 0,
            goal: Math.round(weeklyGoal)
          });
        }
      }
      
      setWeeklyData(chartData);
      
    } catch (error) {
      console.error('Erro ao carregar dados mensais:', error);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadMonthlyData();
  }, [employeeId, monthlyGoal]);

  if (loading) {
    return (
      <Card className="bg-gradient-card border-border/50 shadow-card">
        <CardContent className="flex items-center justify-center py-8">
          <div className="text-center">
            <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-primary mx-auto mb-4"></div>
            <p className="text-muted-foreground">Carregando evolução mensal...</p>
          </div>
        </CardContent>
      </Card>
    );
  }

  const totalPoints = weeklyData.reduce((sum, week) => sum + week.points, 0);
  const totalGoal = weeklyData.reduce((sum, week) => sum + week.goal, 0);
  const progress = totalGoal > 0 ? (totalPoints / totalGoal) * 100 : 0;

  return (
    <Card className="bg-gradient-card border-border/50 shadow-card">
      <CardHeader className="pb-4">
        <CardTitle className="text-xl font-medium text-foreground">
          Evolução Mensal
        </CardTitle>
      </CardHeader>
      
      <CardContent className="space-y-6">
        {/* Gráfico Principal */}
        <div className="bg-card/50 rounded-xl border border-border/20 p-6">
          <div className="h-80">
            <ResponsiveContainer width="100%" height="100%">
              <LineChart data={weeklyData} margin={{ top: 20, right: 30, left: 20, bottom: 20 }}>
                <CartesianGrid 
                  strokeDasharray="3 3" 
                  stroke="hsl(var(--border))" 
                  opacity={0.3}
                />
                <XAxis 
                  dataKey="weekLabel" 
                  tick={{ fontSize: 12, fill: 'hsl(var(--muted-foreground))' }}
                  axisLine={false}
                  tickLine={false}
                />
                <YAxis 
                  tick={{ fontSize: 12, fill: 'hsl(var(--muted-foreground))' }}
                  axisLine={false}
                  tickLine={false}
                />
                <Tooltip 
                  labelFormatter={(value) => value}
                  formatter={(value: number, name: string) => [
                    value.toLocaleString(), 
                    name === 'points' ? 'Pontos' : 'Meta'
                  ]}
                  contentStyle={{
                    backgroundColor: 'hsl(var(--card))',
                    border: '1px solid hsl(var(--border))',
                    borderRadius: '8px',
                    boxShadow: '0 4px 12px hsl(var(--foreground) / 0.1)'
                  }}
                />
                <Legend 
                  wrapperStyle={{ paddingTop: '20px' }}
                  iconType="circle"
                />
                <Line 
                  type="monotone" 
                  dataKey="points" 
                  stroke="hsl(var(--primary))" 
                  strokeWidth={3}
                  dot={{ fill: 'hsl(var(--primary))', strokeWidth: 2, r: 5 }}
                  activeDot={{ r: 7, fill: 'hsl(var(--primary))', strokeWidth: 2 }}
                  name="Pontos"
                />
                <Line 
                  type="monotone" 
                  dataKey="goal" 
                  stroke="hsl(var(--muted-foreground))" 
                  strokeWidth={2}
                  strokeDasharray="8 4"
                  dot={false}
                  name="Meta"
                />
              </LineChart>
            </ResponsiveContainer>
          </div>
        </div>

        {/* Resumo de Performance */}
        <div className="grid grid-cols-3 gap-4">
          <div className="text-center p-4 bg-primary/5 rounded-lg border border-primary/20">
            <div className="text-2xl font-bold text-primary mb-1">
              {totalPoints}
            </div>
            <div className="text-sm text-muted-foreground">
              Total de Pontos
            </div>
          </div>
          
          <div className="text-center p-4 bg-secondary/5 rounded-lg border border-secondary/20">
            <div className="text-2xl font-bold text-secondary mb-1">
              {totalGoal}
            </div>
            <div className="text-sm text-muted-foreground">
              Meta Total
            </div>
          </div>
          
          <div className="text-center p-4 bg-card rounded-lg border border-border">
            <div className={`text-2xl font-bold mb-1 ${progress >= 100 ? 'text-green-600' : progress >= 80 ? 'text-yellow-600' : 'text-red-600'}`}>
              {progress.toFixed(1)}%
            </div>
            <div className="text-sm text-muted-foreground">
              Progresso
            </div>
          </div>
        </div>
      </CardContent>
    </Card>
  );
};