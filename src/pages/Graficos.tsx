import { useState, useEffect } from "react";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Badge } from "@/components/ui/badge";
import { BarChart3, Target, Download, TrendingUp, Users, User } from "lucide-react";
import { useToast } from "@/hooks/use-toast";
import { ExportService } from "@/services/ExportService";
import { DataService } from "@/services/DataService";
import { ExcelDataService } from "@/services/ExcelDataService";
import { ChartContainer } from "@/components/Charts/ChartContainer";
import { WeeklyChart } from "@/components/Charts/WeeklyChart";
import { MonthlyChart } from "@/components/Charts/MonthlyChart";
import { TeamChart } from "@/components/Charts/TeamChart";
import { ChartTypeSelector } from "@/components/Charts/ChartTypeSelector";
import { EmployeeControls } from "@/components/Charts/EmployeeControls";

export default function Graficos() {
  const [selectedChart, setSelectedChart] = useState("weekly");
  const [viewMode, setViewMode] = useState<"team" | "individual">("team");
  const [hiddenEmployees, setHiddenEmployees] = useState<Set<string>>(new Set());
  const [loading, setLoading] = useState(false);
  const [chartData, setChartData] = useState<any>(null);
  const [stats, setStats] = useState<any>(null);
  const [dataSource, setDataSource] = useState<"excel" | "supabase">("excel");
  const { toast } = useToast();

  // Carregar dados reais do Supabase
  useEffect(() => {
    loadData();
  }, []);

  const loadData = async () => {
    setLoading(true);
    try {
      // Configurar fonte de dados
      DataService.setDataSource(dataSource === "excel");
      
      let chartDataResult, statsResult;
      
      if (dataSource === "excel") {
        console.log('📊 Carregando dados dos arquivos Excel...');
        [chartDataResult, statsResult] = await Promise.all([
          ExcelDataService.generateChartDataFromExcel(),
          ExcelDataService.getGeneralStatsFromExcel()
        ]);
        
        // Adaptar formato para compatibilidade
        chartDataResult = {
          ...chartDataResult,
          employeeStats: {}
        };
      } else {
        console.log('📊 Carregando dados do Supabase...');
        [chartDataResult, statsResult] = await Promise.all([
          DataService.getChartData(),
          DataService.getGeneralStats()
        ]);
      }
      
      setChartData(chartDataResult);
      setStats(statsResult);
    } catch (error) {
      console.error('Erro ao carregar dados:', error);
      toast({
        title: "Erro",
        description: `Erro ao carregar dados ${dataSource === "excel" ? "dos arquivos Excel" : "do Supabase"}`,
        variant: "destructive",
      });
    } finally {
      setLoading(false);
    }
  };

  const handleDataSourceChange = async (newSource: "excel" | "supabase") => {
    setDataSource(newSource);
    
    // Limpar cache se mudando para Excel
    if (newSource === "excel") {
      ExcelDataService.clearCache();
    }
    
    // Recarregar dados
    await loadData();
    
    toast({
      title: "Fonte de dados alterada",
      description: `Agora usando dados ${newSource === "excel" ? "dos arquivos Excel" : "do Supabase"}`,
    });
  };
  const toggleEmployee = (employee: string) => {
    setHiddenEmployees(prev => {
      const newSet = new Set(prev);
      if (newSet.has(employee)) {
        newSet.delete(employee);
      } else {
        newSet.add(employee);
      }
      return newSet;
    });
  };

  const handleExportData = async () => {
    setLoading(true);
    try {
      if (dataSource === "excel") {
        // Exportar dados processados dos arquivos Excel
        const processedData = await ExcelDataService.processRegistrosFolder();
        await this.exportExcelData(processedData);
      } else {
        // Exportar dados do Supabase
        await ExportService.exportEmployeeDataToZip();
      }
      
      toast({
        title: "Sucesso",
        description: "Arquivos exportados com sucesso! Verifique sua pasta de downloads.",
      });
    } catch (error) {
      console.error('Erro ao exportar:', error);
      toast({
        title: "Erro",
        description: "Erro ao exportar dados",
        variant: "destructive",
      });
    } finally {
      setLoading(false);
    }
  };

  const exportExcelData = async (processedData: any) => {
    // Implementar exportação dos dados processados
    console.log('📤 Exportando dados processados dos arquivos Excel');
    // Por enquanto, usar o ExportService existente
    await ExportService.exportEmployeeDataToZip();
  };
  const renderChart = () => {
    if (loading || !chartData) {
      return (
        <div className="flex items-center justify-center h-64">
          <div className="text-center">
            <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-primary mx-auto mb-2"></div>
            <p className="text-muted-foreground">
              Carregando dados {dataSource === "excel" ? "dos arquivos Excel" : "do Supabase"}...
            </p>
          </div>
        </div>
      );
    }

    switch (selectedChart) {
      case "weekly":
        return <WeeklyChart data={chartData.weeklyData} hiddenEmployees={hiddenEmployees} />;
      
      case "progress":
        return (
          <MonthlyChart 
            data={chartData.monthlyData} 
            hiddenEmployees={hiddenEmployees} 
            viewMode={viewMode} 
          />
        );
      
      case "team":
        return <TeamChart data={chartData.teamPerformance} />;
      
      default:
        return null;
    }
  };

  const getChartTitle = () => {
    switch (selectedChart) {
      case "weekly":
        return "Progresso Semanal";
      case "progress":
        return "Progresso Mensal";
      case "team":
        return "Distribuição da Equipe";
      default:
        return "Gráfico";
    }
  };

  const getChartIcon = () => {
    switch (selectedChart) {
      case "weekly":
        return <BarChart3 className="h-5 w-5" />;
      case "progress":
        return <TrendingUp className="h-5 w-5" />;
      case "team":
        return <Target className="h-5 w-5" />;
      default:
        return <BarChart3 className="h-5 w-5" />;
    }
  };

  return (
    <div className="min-h-screen bg-background p-6 space-y-6">
      {/* Header compacto */}
      <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4">
        <div>
          <h1 className="text-3xl font-bold text-foreground">Gráficos</h1>
          <p className="text-muted-foreground">
            Visualização de dados e métricas da equipe 
            {dataSource === "excel" ? "(Arquivos Excel)" : "(Supabase)"}
          </p>
        </div>
        
        <div className="flex items-center gap-3">
          {/* Seletor de fonte de dados */}
          <div className="flex items-center gap-2">
            <Button
              variant={dataSource === "excel" ? "dashboard-active" : "outline"}
              size="sm"
              onClick={() => handleDataSourceChange("excel")}
            >
              📁 Excel
            </Button>
            <Button
              variant={dataSource === "supabase" ? "dashboard-active" : "outline"}
              size="sm"
              onClick={() => handleDataSourceChange("supabase")}
            >
              🗄️ Supabase
            </Button>
          </div>
          
          <Button
            onClick={handleExportData}
            disabled={loading}
            variant="outline"
            className="gap-2"
          >
            <Download className="h-4 w-4" />
            {loading ? 'Exportando...' : 'Exportar Excel'}
          </Button>
          
          <Badge variant="outline" className="text-dashboard-info border-dashboard-info/30">
            <Target className="h-3 w-3 mr-1" />
            {dataSource === "excel" ? "Modo Excel" : "Modo Supabase"}
          </Badge>
        </div>
      </div>

      {/* Chart Type Selector */}
      <ChartTypeSelector 
        selectedChart={selectedChart} 
        onChartChange={setSelectedChart} 
      />

      {/* Main Content Grid */}
      <div className="grid grid-cols-12 gap-6">
        {/* Chart Area */}
        <div className="col-span-12 xl:col-span-8">
          <Card className="bg-gradient-card shadow-card border-border">
            <CardHeader className="pb-4">
              <CardTitle className="flex items-center gap-2 text-foreground">
                {getChartIcon()}
                {getChartTitle()}
              </CardTitle>
            </CardHeader>
            <CardContent className="p-6 pt-0">
              <div className="h-96">
                {renderChart()}
              </div>
            </CardContent>
          </Card>
        </div>

        {/* Controls Sidebar */}
        <div className="col-span-12 xl:col-span-4">
          <Card className="bg-gradient-card shadow-card border-border">
            <CardHeader className="pb-3">
              <CardTitle className="text-sm text-foreground">Controles</CardTitle>
            </CardHeader>
            <CardContent className="space-y-4">
              {/* Modo de Visualização */}
              {selectedChart === "progress" && (
                <div className="space-y-3">
                  <h4 className="text-xs font-medium text-muted-foreground uppercase tracking-wide">Visualização</h4>
                  <div className="grid grid-cols-2 gap-2">
                    <Button
                      variant={viewMode === "team" ? "dashboard-active" : "outline"}
                      onClick={() => setViewMode("team")}
                      className="h-8 text-xs"
                      size="sm"
                    >
                      <Users className="h-3 w-3 mr-1" />
                      Equipe
                    </Button>
                    <Button
                      variant={viewMode === "individual" ? "dashboard-active" : "outline"}
                      onClick={() => setViewMode("individual")}
                      className="h-8 text-xs"
                      size="sm"
                    >
                      <User className="h-3 w-3 mr-1" />
                      Individual
                    </Button>
                  </div>
                </div>
              )}

              {/* Funcionários */}
              <div className="space-y-3">
                <h4 className="text-xs font-medium text-muted-foreground uppercase tracking-wide">Funcionários</h4>
                <EmployeeControls
                  viewMode={viewMode}
                  onViewModeChange={setViewMode}
                  hiddenEmployees={hiddenEmployees}
                  onToggleEmployee={toggleEmployee}
                  selectedChart={selectedChart}
                />
              </div>
            </CardContent>
          </Card>
        </div>
      </div>

      {/* Statistics Cards - Separated Below Chart */}
      {stats && (
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
          {/* Melhor Card */}
          <Card className="bg-gradient-card shadow-card border-border">
            <CardContent className="p-4">
              <div className="text-center">
                <div className="text-xs text-muted-foreground mb-1">Melhor</div>
                <div className="font-semibold text-dashboard-success text-lg">{stats.bestPerformer}</div>
                <div className="text-sm text-dashboard-success">{stats.bestPoints} pts</div>
              </div>
            </CardContent>
          </Card>

          {/* Média Card */}
          <Card className="bg-gradient-card shadow-card border-border">
            <CardContent className="p-4">
              <div className="text-center">
                <div className="text-xs text-muted-foreground mb-1">Média</div>
                <div className="font-semibold text-dashboard-info text-lg">{stats.avgTeam} pts</div>
                <div className="text-xs text-muted-foreground">por funcionário</div>
              </div>
            </CardContent>
          </Card>

          {/* Meta Card */}
          <Card className="bg-gradient-card shadow-card border-border">
            <CardContent className="p-4">
              <div className="text-center">
                <div className="text-xs text-muted-foreground mb-1">Meta</div>
                <div className="font-semibold text-dashboard-warning text-lg">{stats.totalGoal}K</div>
                <div className="text-xs text-muted-foreground">mensal</div>
              </div>
            </CardContent>
          </Card>

          {/* Progresso Card */}
          <Card className="bg-gradient-card shadow-card border-border">
            <CardContent className="p-4">
              <div className="text-center">
                <div className="text-xs text-muted-foreground mb-1">Progresso</div>
                <div className={`font-semibold text-lg ${
                  stats.progressPercentage >= 80 
                    ? "text-dashboard-success" 
                    : stats.progressPercentage >= 60 
                    ? "text-dashboard-warning"
                    : "text-dashboard-danger"
                }`}>
                  {stats.progressPercentage}%
                </div>
                <div className="text-xs text-muted-foreground">da meta</div>
              </div>
            </CardContent>
          </Card>
        </div>
      )}
    </div>
  );
}