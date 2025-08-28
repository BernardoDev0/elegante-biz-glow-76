import React, { useState, useEffect } from 'react';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table';
import { ChartContainer, ChartTooltip, ChartTooltipContent } from '@/components/ui/chart';
import { Line, LineChart, Bar, BarChart, ResponsiveContainer, XAxis, YAxis, CartesianGrid, Legend } from 'recharts';
import { useToast } from '@/hooks/use-toast';

interface MonthlyData {
  labels: string[];
  points: number[];
  goals: number[];
}

interface WeeklyData {
  labels: string[];
  points: number[];
}

interface HistoryEntry {
  id: string;
  date: string;
  refinery: string;
  points: number;
  observations: string;
}

interface PaginationInfo {
  current_page: number;
  total_pages: number;
  total_entries: number;
}

interface EmployeeDashboardProps {
  employeeId: string;
  monthlyData?: MonthlyData;
  weeklyData?: WeeklyData;
}

const monthlyChartConfig = {
  points: {
    label: "Pontos",
    color: "hsl(var(--primary))",
  },
  goals: {
    label: "Meta",
    color: "hsl(var(--warning))",
  },
};

const weeklyChartConfig = {
  points: {
    label: "Pontos",
    color: "hsl(var(--primary))",
  },
};

export default function EmployeeDashboard({ employeeId, monthlyData, weeklyData }: EmployeeDashboardProps) {
  const [currentPage, setCurrentPage] = useState(1);
  const [selectedWeek, setSelectedWeek] = useState('');
  const [historyEntries, setHistoryEntries] = useState<HistoryEntry[]>([]);
  const [pagination, setPagination] = useState<PaginationInfo | null>(null);
  const [loading, setLoading] = useState(false);
  const [activeTab, setActiveTab] = useState('dashboard');
  const { toast } = useToast();

  // Transform data for charts
  const monthlyChartData = monthlyData ? 
    monthlyData.labels.map((label, index) => ({
      month: label,
      points: monthlyData.points[index] || 0,
      goals: monthlyData.goals[index] || 0,
    })) : [];

  const weeklyChartData = weeklyData ?
    weeklyData.labels.map((label, index) => ({
      week: label,
      points: weeklyData.points[index] || 0,
    })) : [];

  // Fetch history entries
  const fetchHistory = async (page: number, week?: string) => {
    setLoading(true);
    try {
      let url = `/api/entries?page=${page}&employee_id=${employeeId}`;
      if (week && week !== '') {
        url += `&week=${week}`;
      }

      const response = await fetch(url);
      const data = await response.json();
      
      if (data.entries) {
        setHistoryEntries(data.entries);
        setCurrentPage(page);
        setPagination(data.pagination);
      }
    } catch (error) {
      console.error('Erro ao buscar histórico:', error);
      toast({
        title: "Erro",
        description: "Não foi possível carregar o histórico.",
        variant: "destructive",
      });
    } finally {
      setLoading(false);
    }
  };

  // Handle week filter change
  const handleWeekChange = (week: string) => {
    setSelectedWeek(week);
    if (activeTab === 'history') {
      fetchHistory(1, week);
    }
  };

  // Handle tab change
  const handleTabChange = (tab: string) => {
    setActiveTab(tab);
    if (tab === 'history' && historyEntries.length === 0) {
      fetchHistory(1, selectedWeek);
    }
  };

  // Format date for display
  const formatDate = (dateString: string) => {
    const [datePart, timePart] = dateString.split(' ');
    const [yyyy, mm, dd] = datePart.split('-');
    const brDate = `${dd}/${mm}/${yyyy}`;
    const brTime = timePart ? timePart.substring(0, 5) : '-';
    return { brDate, brTime };
  };

  useEffect(() => {
    // Handle URL hash for tab navigation
    const hash = window.location.hash.replace('#', '');
    if (hash && ['dashboard', 'history'].includes(hash)) {
      setActiveTab(hash);
    }
  }, []);

  return (
    <div className="space-y-6">
      <Tabs value={activeTab} onValueChange={handleTabChange} className="w-full">
        <TabsList className="grid w-full grid-cols-2">
          <TabsTrigger value="dashboard">Dashboard</TabsTrigger>
          <TabsTrigger value="history">Histórico</TabsTrigger>
        </TabsList>

        <TabsContent value="dashboard" className="space-y-6">
          {/* Monthly Evolution Chart */}
          {monthlyData && (
            <Card>
              <CardHeader>
                <CardTitle>Evolução Mensal</CardTitle>
              </CardHeader>
              <CardContent>
                <ChartContainer config={monthlyChartConfig} className="h-[300px]">
                  <ResponsiveContainer width="100%" height="100%">
                    <LineChart data={monthlyChartData}>
                      <CartesianGrid strokeDasharray="3 3" stroke="hsl(var(--border))" />
                      <XAxis 
                        dataKey="month" 
                        stroke="hsl(var(--muted-foreground))"
                      />
                      <YAxis 
                        stroke="hsl(var(--muted-foreground))"
                      />
                      <ChartTooltip content={<ChartTooltipContent />} />
                      <Legend />
                      <Line
                        type="monotone"
                        dataKey="points"
                        stroke="var(--color-points)"
                        strokeWidth={2}
                        dot={{ fill: "var(--color-points)" }}
                      />
                      <Line
                        type="monotone"
                        dataKey="goals"
                        stroke="var(--color-goals)"
                        strokeWidth={2}
                        strokeDasharray="5 5"
                        dot={{ fill: "var(--color-goals)" }}
                      />
                    </LineChart>
                  </ResponsiveContainer>
                </ChartContainer>
              </CardContent>
            </Card>
          )}

          {/* Weekly Performance Chart */}
          {weeklyData && (
            <Card>
              <CardHeader>
                <CardTitle>Desempenho Semanal</CardTitle>
              </CardHeader>
              <CardContent>
                <ChartContainer config={weeklyChartConfig} className="h-[300px]">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={weeklyChartData}>
                      <CartesianGrid strokeDasharray="3 3" stroke="hsl(var(--border))" />
                      <XAxis 
                        dataKey="week" 
                        stroke="hsl(var(--muted-foreground))"
                      />
                      <YAxis 
                        stroke="hsl(var(--muted-foreground))"
                      />
                      <ChartTooltip content={<ChartTooltipContent />} />
                      <Bar dataKey="points" fill="var(--color-points)" />
                    </BarChart>
                  </ResponsiveContainer>
                </ChartContainer>
              </CardContent>
            </Card>
          )}
        </TabsContent>

        <TabsContent value="history" className="space-y-6">
          {/* Week Filter */}
          <Card>
            <CardHeader>
              <CardTitle>Filtros</CardTitle>
            </CardHeader>
            <CardContent>
              <Select value={selectedWeek} onValueChange={handleWeekChange}>
                <SelectTrigger className="w-[200px]">
                  <SelectValue placeholder="Selecione uma semana" />
                </SelectTrigger>
                <SelectContent>
                  <SelectItem value="">Todas as semanas</SelectItem>
                  {/* Add week options here based on available data */}
                  <SelectItem value="2024-W01">Semana 1</SelectItem>
                  <SelectItem value="2024-W02">Semana 2</SelectItem>
                  <SelectItem value="2024-W03">Semana 3</SelectItem>
                  <SelectItem value="2024-W04">Semana 4</SelectItem>
                </SelectContent>
              </Select>
            </CardContent>
          </Card>

          {/* History Table */}
          <Card>
            <CardHeader>
              <CardTitle>Histórico de Registros</CardTitle>
            </CardHeader>
            <CardContent>
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead>Data</TableHead>
                    <TableHead>Hora</TableHead>
                    <TableHead>Refinaria</TableHead>
                    <TableHead>Pontos</TableHead>
                    <TableHead>Observações</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {loading ? (
                    <TableRow>
                      <TableCell colSpan={5} className="text-center">
                        Carregando...
                      </TableCell>
                    </TableRow>
                  ) : historyEntries.length > 0 ? (
                    historyEntries.map((entry) => {
                      const { brDate, brTime } = formatDate(entry.date);
                      return (
                        <TableRow key={entry.id}>
                          <TableCell>{brDate}</TableCell>
                          <TableCell>{brTime}</TableCell>
                          <TableCell>{entry.refinery || ''}</TableCell>
                          <TableCell>{entry.points}</TableCell>
                          <TableCell>{entry.observations || ''}</TableCell>
                        </TableRow>
                      );
                    })
                  ) : (
                    <TableRow>
                      <TableCell colSpan={5} className="text-center">
                        Nenhum registro encontrado.
                      </TableCell>
                    </TableRow>
                  )}
                </TableBody>
              </Table>

              {/* Pagination */}
              {pagination && pagination.total_pages > 1 && (
                <div className="flex items-center justify-between space-x-2 py-4">
                  <div className="text-sm text-muted-foreground">
                    Página {pagination.current_page} de {pagination.total_pages}
                  </div>
                  <div className="flex space-x-2">
                    <button
                      className="px-3 py-1 text-sm border rounded hover:bg-accent"
                      onClick={() => fetchHistory(currentPage - 1, selectedWeek)}
                      disabled={currentPage <= 1}
                    >
                      Anterior
                    </button>
                    <button
                      className="px-3 py-1 text-sm border rounded hover:bg-accent"
                      onClick={() => fetchHistory(currentPage + 1, selectedWeek)}
                      disabled={currentPage >= pagination.total_pages}
                    >
                      Próxima
                    </button>
                  </div>
                </div>
              )}
            </CardContent>
          </Card>
        </TabsContent>
      </Tabs>
    </div>
  );
}