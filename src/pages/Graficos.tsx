import { useState, useEffect } from "react";
import { useNavigate } from "react-router-dom";
import { MetricCard } from "@/components/Dashboard/MetricCard";
import { EmployeeCard } from "@/components/Dashboard/EmployeeCard";
import { ProgressSection } from "@/components/Dashboard/ProgressSection";
import { EmployeeService, Employee } from "@/services/EmployeeService";
import { CalculationsService } from "@/services/CalculationsService";

interface EmployeeMetrics extends Employee {
  weeklyPoints: number;
  weeklyGoal: number;
  monthlyPoints: number;
  monthlyGoal: number;
  weeklyProgress: number;
  monthlyProgress: number;
  status: "below" | "on-track" | "above";
}

const Index = () => {
  const [selectedWeek, setSelectedWeek] = useState("1");
  const [employees, setEmployees] = useState<EmployeeMetrics[]>([]);
  const [loading, setLoading] = useState(true);
  const navigate = useNavigate();

  // Carregar dados dos funcionários
  const loadEmployeesData = async () => {
    try {
      setLoading(true);
      const allEmployees = await EmployeeService.getAllEmployees();
      
      // Calcular métricas para cada funcionário
      const employeesWithMetrics = await Promise.all(
        allEmployees.map(async (employee) => {
          const weekDates = CalculationsService.getWeekDates(selectedWeek);
          const monthDates = CalculationsService.getMonthCycleDates();
          
          const weeklyPoints = await EmployeeService.getWeekPoints(employee.id, weekDates);
          const monthlyPoints = await EmployeeService.getMonthPoints(employee.id, monthDates);
          
          const weeklyGoal = CalculationsService.getWeeklyGoal(employee);
          const monthlyGoal = CalculationsService.getMonthlyGoal(employee);
          
          const weeklyProgress = CalculationsService.calculateProgressPercentage(weeklyPoints, weeklyGoal);
          const monthlyProgress = CalculationsService.calculateProgressPercentage(monthlyPoints, monthlyGoal);
          
          const status = weeklyProgress >= 90 ? "above" : weeklyProgress >= 70 ? "on-track" : "below";
          
          return {
            ...employee,
            weeklyPoints,
            weeklyGoal,
            monthlyPoints,
            monthlyGoal,
            weeklyProgress,
            monthlyProgress,
            status
          } as EmployeeMetrics;
        })
      );
      
      setEmployees(employeesWithMetrics);
    } catch (error) {
      console.error('Erro ao carregar dados dos funcionários:', error);
    } finally {
      setLoading(false);
    }
  };

  // Verificar se o usuário tem permissão (CEO)
  useEffect(() => {
    const userData = localStorage.getItem("currentUser");
    if (!userData) {
      navigate("/login");
      return;
    }
    
    const user = JSON.parse(userData);
    if (user.role !== "ceo") {
      navigate("/dashboard");
      return;
    }
    
    loadEmployeesData();
  }, [navigate, selectedWeek]);

  // Calcular métricas totais
  const totalWeeklyPoints = employees.reduce((sum, emp) => sum + emp.weeklyPoints, 0);
  const totalMonthlyPoints = employees.reduce((sum, emp) => sum + emp.monthlyPoints, 0);
  const totalMonthlyGoal = employees.reduce((sum, emp) => sum + emp.monthlyGoal, 0);
  const teamProgress = (totalMonthlyPoints / totalMonthlyGoal) * 100;

  if (loading) {
    return (
      <div className="min-h-screen bg-background flex items-center justify-center">
        <div className="text-center">
          <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-primary mx-auto mb-4"></div>
          <p className="text-muted-foreground">Carregando dados da equipe...</p>
        </div>
      </div>
    );
  }

  return (
    <div className="space-y-6">
      {/* Header */}
      <div>
        <div>
          <h1 className="text-2xl font-bold text-foreground">Equipe</h1>
          <p className="text-muted-foreground">Visão geral da performance da equipe</p>
        </div>
      </div>

      {/* Métricas Principais */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
        {[
          {
            title: "Total de Pontos Desta Semana",
            value: totalWeeklyPoints.toLocaleString(),
            icon: "trending" as const,
            variant: "primary" as const
          },
          {
            title: "Pontos Mensais Totais",
            value: `${(totalMonthlyPoints / 1000).toFixed(3)}`,
            icon: "target" as const,
            variant: "default" as const
          },
          {
            title: "Progresso da Equipe Esta Semana",
            value: `${teamProgress.toFixed(1)}%`,
            icon: "users" as const,
            variant: teamProgress >= 50 ? ("success" as const) : ("warning" as const)
          }
        ].map((metric, index) => (
          <div key={metric.title}>
            <MetricCard {...metric} />
          </div>
        ))}
      </div>

      {/* Progresso Geral */}
      <div>
        <ProgressSection
          totalPoints={totalMonthlyPoints}
          totalGoal={totalMonthlyGoal} 
          completedPercentage={teamProgress}
          selectedWeek={selectedWeek}
          onWeekChange={setSelectedWeek}
        />
      </div>

      {/* Cards dos Funcionários */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {employees.map((employee, index) => (
          <div key={employee.name}>
            <EmployeeCard
              name={employee.real_name || employee.name}
              role={employee.role}
              weeklyPoints={employee.weeklyPoints}
              weeklyGoal={employee.weeklyGoal}
              monthlyPoints={employee.monthlyPoints}
              monthlyGoal={employee.monthlyGoal}
              weeklyProgress={employee.weeklyProgress}
              monthlyProgress={employee.monthlyProgress}
              status={employee.status}
            />
          </div>
        ))}
      </div>
    </div>
  );
};

export default Index;