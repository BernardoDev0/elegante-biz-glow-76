import { useState, useEffect } from "react";
import { useNavigate } from "react-router-dom";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Textarea } from "@/components/ui/textarea";
import { Progress } from "@/components/ui/progress";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { LogOut, TrendingUp, Target, Calendar } from "lucide-react";
import { useToast } from "@/hooks/use-toast";
import { EmployeeService, Employee } from "@/services/EmployeeService";
import { CalculationsService } from "@/services/CalculationsService";
import { HistoryTab } from "@/components/Dashboard/HistoryTab";
import { MonthlyEvolutionTab } from "@/components/Dashboard/MonthlyEvolutionTab";

interface DashboardMetrics {
  todayPoints: number;
  weekPoints: number;
  monthPoints: number;
  dailyGoal: number;
  weeklyGoal: number;
  monthlyGoal: number;
}

const Dashboard = () => {
  const [currentUser, setCurrentUser] = useState<Employee | null>(null);
  const [selectedWeek, setSelectedWeek] = useState("1");
  const [pontos, setPontos] = useState("");
  const [observacoes, setObservacoes] = useState("");
  const [selectedRefinery, setSelectedRefinery] = useState("");
  const [metrics, setMetrics] = useState<DashboardMetrics>({
    todayPoints: 0,
    weekPoints: 0,
    monthPoints: 0,
    dailyGoal: 475,
    weeklyGoal: 2375,
    monthlyGoal: 9500
  });
  const [loading, setLoading] = useState(true);
  const navigate = useNavigate();
  const { toast } = useToast();

  const progressoDiario = CalculationsService.calculateProgressPercentage(metrics.todayPoints, metrics.dailyGoal);
  const progressoSemanal = CalculationsService.calculateProgressPercentage(metrics.weekPoints, metrics.weeklyGoal);
  const progressoMensal = CalculationsService.calculateProgressPercentage(metrics.monthPoints, metrics.monthlyGoal);

  // Carregar dados do funcionário e métricas
  const loadEmployeeData = async (employeeId: number) => {
    try {
      setLoading(true);
      
      // Buscar dados do funcionário
      const employee = await EmployeeService.getEmployeeById(employeeId);
      if (!employee) {
        toast({
          title: "Erro",
          description: "Funcionário não encontrado.",
          variant: "destructive"
        });
        navigate("/login");
        return;
      }

      setCurrentUser(employee);

      // Definir refinaria padrão
      setSelectedRefinery(employee.default_refinery || "RPBC");

      // Calcular metas baseadas no funcionário
      const dailyGoal = CalculationsService.getDailyGoal(employee);
      const weeklyGoal = CalculationsService.getWeeklyGoal(employee);
      const monthlyGoal = CalculationsService.getMonthlyGoal(employee);

      // Calcular pontos atuais
      const todayPoints = await EmployeeService.getTodayPoints(employeeId);
      
      // Calcular pontos da semana selecionada
      const weekDates = CalculationsService.getWeekDates(selectedWeek);
      const weekPoints = await EmployeeService.getWeekPoints(employeeId, weekDates);

      // Calcular pontos mensais
      const monthDates = CalculationsService.getMonthCycleDates();
      const monthPoints = await EmployeeService.getMonthPoints(employeeId, monthDates);

      setMetrics({
        todayPoints,
        weekPoints,
        monthPoints,
        dailyGoal,
        weeklyGoal,
        monthlyGoal
      });

    } catch (error) {
      console.error('Erro ao carregar dados:', error);
      toast({
        title: "Erro",
        description: "Não foi possível carregar os dados.",
        variant: "destructive"
      });
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    const userData = localStorage.getItem("currentUser");
    if (!userData) {
      navigate("/login");
      return;
    }
    
    try {
      const user = JSON.parse(userData);
      if (user.id) {
        loadEmployeeData(user.id);
      } else {
        navigate("/login");
      }
    } catch (error) {
      console.error('Erro ao parsear dados do usuário:', error);
      navigate("/login");
    }
  }, [navigate, selectedWeek]); // Recarregar quando mudar semana

  const handleLogout = () => {
    localStorage.removeItem("currentUser");
    toast({
      title: "Logout realizado",
      description: "Até logo!"
    });
    navigate("/login");
  };

  const handleRegistrar = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!pontos || !observacoes || !currentUser) {
      toast({
        title: "Erro",
        description: "Por favor, preencha todos os campos.",
        variant: "destructive"
      });
      return;
    }

    try {
      setLoading(true);
      
      // Criar novo registro
      const newEntry = await EmployeeService.createEntry({
        employee_id: currentUser.id,
        date: new Date().toISOString(),
        refinery: selectedRefinery,
        points: parseInt(pontos),
        observations: observacoes
      });

      if (newEntry) {
        toast({
          title: "Registro salvo!",
          description: `${pontos} pontos registrados com sucesso.`
        });
        
        // Limpar formulário
        setPontos("");
        setObservacoes("");
        
        // Recarregar métricas
        await loadEmployeeData(currentUser.id);
      } else {
        throw new Error('Falha ao salvar registro');
      }
    } catch (error) {
      console.error('Erro ao registrar pontos:', error);
      toast({
        title: "Erro",
        description: "Não foi possível salvar o registro.",
        variant: "destructive"
      });
    } finally {
      setLoading(false);
    }
  };

  if (loading || !currentUser) {
    return (
      <div className="min-h-screen bg-background flex items-center justify-center">
        <div className="text-center">
          <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-primary mx-auto mb-4"></div>
          <p className="text-muted-foreground">Carregando dados...</p>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-background">
      {/* Header */}
      <header className="bg-gradient-header border-b border-border/50 shadow-header">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="flex justify-between items-center h-16">
            <div>
              <h1 className="text-2xl font-bold bg-gradient-primary bg-clip-text text-transparent">
                Dashboard de {currentUser.real_name || currentUser.name}
              </h1>
              <p className="text-sm text-muted-foreground">{currentUser.default_refinery}</p>
            </div>
            
            <div className="flex items-center space-x-4">
              <div className="flex items-center space-x-2">
                <Label className="text-sm text-muted-foreground">Semana:</Label>
                <Select value={selectedWeek} onValueChange={setSelectedWeek}>
                  <SelectTrigger className="w-32 bg-secondary border-border">
                    <SelectValue />
                  </SelectTrigger>
                  <SelectContent className="bg-card border-border">
                    <SelectItem value="1">Semana 1</SelectItem>
                    <SelectItem value="2">Semana 2</SelectItem>
                    <SelectItem value="3">Semana 3</SelectItem>
                    <SelectItem value="4">Semana 4</SelectItem>
                    <SelectItem value="5">Semana 5</SelectItem>
                  </SelectContent>
                </Select>
              </div>
              
              <Button 
                variant="outline" 
                onClick={handleLogout}
                className="border-border hover:bg-destructive hover:text-destructive-foreground"
              >
                <LogOut className="w-4 h-4 mr-2" />
                Sair
              </Button>
            </div>
          </div>
        </div>
      </header>

      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
        {/* Cards de Métricas */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-6 mb-8">
          <Card className="bg-gradient-card border-border/50 shadow-card">
            <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2">
              <CardTitle className="text-sm font-medium text-muted-foreground">
                Pontos Hoje
              </CardTitle>
              <TrendingUp className="h-4 w-4 text-dashboard-primary" />
            </CardHeader>
            <CardContent>
              <div className="text-2xl font-bold text-dashboard-primary mb-1">
                {metrics.todayPoints}
              </div>
              <p className="text-xs text-muted-foreground mb-2">
                Meta diária: {metrics.dailyGoal} pontos
              </p>
              <Progress value={progressoDiario} className="h-2" />
              <p className="text-xs text-muted-foreground mt-1">
                {progressoDiario}% da meta diária
              </p>
            </CardContent>
          </Card>

          <Card className="bg-gradient-card border-border/50 shadow-card">
            <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2">
              <CardTitle className="text-sm font-medium text-muted-foreground">
                Pontos Semana {selectedWeek}
              </CardTitle>
              <Target className="h-4 w-4 text-dashboard-secondary" />
            </CardHeader>
            <CardContent>
              <div className="text-2xl font-bold text-dashboard-secondary mb-1">
                {metrics.weekPoints}
              </div>
              <p className="text-xs text-muted-foreground mb-2">
                Meta semanal: {metrics.weeklyGoal} pontos
              </p>
              <Progress value={progressoSemanal} className="h-2" />
              <p className="text-xs text-muted-foreground mt-1">
                {progressoSemanal}% da meta semanal
              </p>
            </CardContent>
          </Card>

          <Card className="bg-gradient-card border-border/50 shadow-card">
            <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2">
              <CardTitle className="text-sm font-medium text-muted-foreground">
                Pontos Mensais
              </CardTitle>
              <Calendar className="h-4 w-4 text-dashboard-success" />
            </CardHeader>
            <CardContent>
              <div className="text-2xl font-bold text-dashboard-success mb-1">
                {metrics.monthPoints}
              </div>
              <p className="text-xs text-muted-foreground mb-2">
                Meta mensal: {metrics.monthlyGoal} pontos
              </p>
              <Progress value={progressoMensal} className="h-2" />
              <p className="text-xs text-muted-foreground mt-1">
                {progressoMensal}% da meta mensal
              </p>
            </CardContent>
          </Card>
        </div>

        {/* Abas */}
        <Tabs defaultValue="registrar" className="w-full">
          <TabsList className="grid w-full grid-cols-3 bg-card/50 border border-border">
            <TabsTrigger 
              value="evolucao" 
              className="data-[state=active]:bg-gradient-primary data-[state=active]:text-white"
            >
              Evolução Mensal
            </TabsTrigger>
            <TabsTrigger 
              value="registrar"
              className="data-[state=active]:bg-gradient-primary data-[state=active]:text-white"
            >
              Registrar Pontos
            </TabsTrigger>
            <TabsTrigger 
              value="historico"
              className="data-[state=active]:bg-gradient-primary data-[state=active]:text-white"
            >
              Histórico
            </TabsTrigger>
          </TabsList>

          <TabsContent value="evolucao" className="mt-6">
            <MonthlyEvolutionTab 
              employeeId={currentUser.id} 
              monthlyGoal={metrics.monthlyGoal}
            />
          </TabsContent>

          <TabsContent value="registrar" className="mt-6">
            <Card className="bg-gradient-card border-border/50 shadow-card">
              <CardHeader>
                <CardTitle>Novo Registro</CardTitle>
              </CardHeader>
              <CardContent>
                <form onSubmit={handleRegistrar} className="space-y-4">
                  <div className="space-y-2">
                    <Label htmlFor="refinaria">Refinaria:</Label>
                    <Select value={selectedRefinery} onValueChange={setSelectedRefinery}>
                      <SelectTrigger>
                        <SelectValue placeholder="Selecione a refinaria" />
                      </SelectTrigger>
                      <SelectContent className="bg-card border-border">
                        <SelectItem value="RPBC">RPBC</SelectItem>
                        <SelectItem value="REVAP">REVAP</SelectItem>
                        <SelectItem value="REPAR">REPAR</SelectItem>
                        <SelectItem value="REFAP">REFAP</SelectItem>
                        <SelectItem value="REMAN">REMAN</SelectItem>
                        <SelectItem value="REDUC">REDUC</SelectItem>
                        <SelectItem value="REGAP">REGAP</SelectItem>
                        <SelectItem value="RELAN">RELAN</SelectItem>
                        <SelectItem value="RNEST">RNEST</SelectItem>
                      </SelectContent>
                    </Select>
                  </div>

                  <div className="space-y-2">
                    <Label htmlFor="pontos">Pontos:</Label>
                    <Input
                      id="pontos"
                      type="number"
                      value={pontos}
                      onChange={(e) => setPontos(e.target.value)}
                      placeholder="Digite os pontos"
                    />
                  </div>

                  <div className="space-y-2">
                    <Label htmlFor="observacoes">Observações:</Label>
                    <Textarea
                      id="observacoes"
                      value={observacoes}
                      onChange={(e) => setObservacoes(e.target.value)}
                      placeholder="Digite suas observações"
                      className="min-h-24"
                    />
                  </div>

                  <Button 
                    type="submit" 
                    className="bg-gradient-primary hover:opacity-90 text-white"
                    disabled={loading}
                  >
                    {loading ? 'Registrando...' : 'Registrar'}
                  </Button>
                </form>
              </CardContent>
            </Card>
          </TabsContent>

          <TabsContent value="historico" className="mt-6">
            <HistoryTab employeeId={currentUser.id} />
          </TabsContent>
        </Tabs>
      </div>
    </div>
  );
};

export default Dashboard;