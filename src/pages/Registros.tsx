import { useState, useEffect } from "react";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Badge } from "@/components/ui/badge";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { useToast } from "@/hooks/use-toast";
import { ExcelFolderService } from "@/services/ExcelFolderService";
import { format, parseISO } from 'date-fns';
import { ptBR } from 'date-fns/locale';
import * as XLSX from 'xlsx';
import { CalculationsService } from "@/services/CalculationsService";
import { 
  Search, 
  Filter, 
  Download, 
  Trash2, 
  Edit, 
  Calendar,
  Clock,
  User,
  Building2,
  Hash,
  MessageSquare
} from "lucide-react";

interface EntryRecord {
  id: number;
  date: string;
  time: string;
  employee: string;
  refinery: string;
  points: number;
  observations: string;
  status: "completed" | "absent" | "pending";
}

export default function Registros() {
  const { toast } = useToast();
  const [selectedWeek, setSelectedWeek] = useState("todas");
  const [selectedEmployee, setSelectedEmployee] = useState("todos");
  const [searchTerm, setSearchTerm] = useState("");
  const [records, setRecords] = useState<EntryRecord[]>([]);
  const [employees, setEmployees] = useState<string[]>([]);
  const [loading, setLoading] = useState(true);

  // Carregar dados dos arquivos Excel
  useEffect(() => {
    loadRecords();
  }, []);

  const loadRecords = async () => {
    try {
      setLoading(true);
      
      // Buscar registros dos arquivos Excel
      const excelRecords = await ExcelFolderService.getFilteredRecords({
        employee: selectedEmployee,
        week: selectedWeek,
        searchTerm: searchTerm
      });

      // Transformar dados para o formato da interface
      const transformedRecords: EntryRecord[] = excelRecords.map((record, index) => ({
        id: index + 1, // ID sequencial
        date: format(new Date(record.date), 'dd/MM/yyyy', { locale: ptBR }),
        time: format(new Date(record.date), 'HH:mm', { locale: ptBR }),
        employee: record.employee,
        refinery: record.refinery,
        points: record.points,
        observations: record.observations,
        status: record.points > 0 ? "completed" : "absent"
      }));

      setRecords(transformedRecords);

      // Extrair funcionários únicos
      const folderData = await ExcelFolderService.processRegistrosFolder();
      const uniqueEmployees = Object.keys(folderData.employees);
      setEmployees(uniqueEmployees);

    } catch (error) {
      console.error('Erro ao carregar registros:', error);
      toast({
        title: "Erro",
        description: "Erro ao carregar registros",
        variant: "destructive",
      });
    } finally {
      setLoading(false);
    }
  };

  // Recarregar quando filtros mudarem
  useEffect(() => {
    loadRecords();
  }, [selectedEmployee, selectedWeek, searchTerm]);

  const filteredRecords = records.filter(record => {
    const matchesSearch = record.employee.toLowerCase().includes(searchTerm.toLowerCase()) ||
                         record.refinery.toLowerCase().includes(searchTerm.toLowerCase()) ||
                         record.observations.toLowerCase().includes(searchTerm.toLowerCase());
    
    const matchesEmployee = selectedEmployee === "todos" || record.employee === selectedEmployee;
    
    // Os filtros já foram aplicados no ExcelFolderService
    const matchesWeek = true;
    
    return matchesSearch && matchesEmployee && matchesWeek;
  });

  const exportToExcel = () => {
    try {
      // Preparar dados para exportação
      const exportData = filteredRecords.map(record => ({
        'Data': record.date,
        'Horário': record.time,
        'Funcionário': record.employee,
        'Refinaria': record.refinery,
        'Pontos': record.points,
        'Observações': record.observations,
        'Status': record.status === 'completed' ? 'Concluído' : 
                  record.status === 'absent' ? 'Ausente' : 'Pendente'
      }));

      // Criar workbook
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.json_to_sheet(exportData);

      // Ajustar largura das colunas
      const colWidths = [
        { wch: 12 }, // Data
        { wch: 8 },  // Horário
        { wch: 15 }, // Funcionário
        { wch: 10 }, // Refinaria
        { wch: 8 },  // Pontos
        { wch: 30 }, // Observações
        { wch: 10 }  // Status
      ];
      ws['!cols'] = colWidths;

      XLSX.utils.book_append_sheet(wb, ws, "Registros");

      // Gerar nome do arquivo com data atual
      const fileName = `registros_${format(new Date(), 'yyyy-MM-dd')}.xlsx`;
      
      // Salvar arquivo
      XLSX.writeFile(wb, fileName);

      toast({
        title: "Sucesso",
        description: `Arquivo ${fileName} baixado com sucesso!`,
      });

    } catch (error) {
      console.error('Erro ao exportar:', error);
      toast({
        title: "Erro",
        description: "Erro ao exportar dados para Excel",
        variant: "destructive",
      });
    }
  };

  const getStatusBadge = (status: string) => {
    switch (status) {
      case "completed":
        return <Badge variant="default" className="bg-dashboard-success/20 text-dashboard-success border-dashboard-success/30">Concluído</Badge>;
      case "absent":
        return <Badge variant="destructive">Ausente</Badge>;
      case "pending":
        return <Badge variant="secondary" className="bg-dashboard-warning/20 text-dashboard-warning border-dashboard-warning/30">Pendente</Badge>;
      default:
        return <Badge variant="outline">Desconhecido</Badge>;
    }
  };

  const totalPoints = filteredRecords.reduce((sum, record) => sum + record.points, 0);
  const completedRecords = filteredRecords.filter(r => r.status === "completed").length;

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="flex flex-col gap-4">
        <div>
          <h1 className="text-2xl font-bold text-foreground">Registros</h1>
          <p className="text-muted-foreground">Gerenciamento e histórico de registros da equipe</p>
        </div>

        {/* Summary Cards */}
        <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
          <Card className="bg-gradient-card shadow-card border-border">
            <CardContent className="p-4">
              <div className="flex items-center gap-2">
                <Hash className="h-4 w-4 text-dashboard-primary" />
                <span className="text-sm font-medium text-foreground">Total de Registros</span>
              </div>
              <p className="text-2xl font-bold text-dashboard-primary mt-2">{filteredRecords.length}</p>
            </CardContent>
          </Card>
          
          <Card className="bg-gradient-card shadow-card border-border">
            <CardContent className="p-4">
              <div className="flex items-center gap-2">
                <User className="h-4 w-4 text-dashboard-success" />
                <span className="text-sm font-medium text-foreground">Concluídos</span>
              </div>
              <p className="text-2xl font-bold text-dashboard-success mt-2">{completedRecords}</p>
            </CardContent>
          </Card>
          
          <Card className="bg-gradient-card shadow-card border-border">
            <CardContent className="p-4">
              <div className="flex items-center gap-2">
                <Building2 className="h-4 w-4 text-dashboard-info" />
                <span className="text-sm font-medium text-foreground">Total de Pontos</span>
              </div>
              <p className="text-2xl font-bold text-dashboard-info mt-2">{totalPoints.toLocaleString()}</p>
            </CardContent>
          </Card>
          
          <Card className="bg-gradient-card shadow-card border-border">
            <CardContent className="p-4">
              <div className="flex items-center gap-2">
                <Clock className="h-4 w-4 text-dashboard-warning" />
                <span className="text-sm font-medium text-foreground">Média Diária</span>
              </div>
              <p className="text-2xl font-bold text-dashboard-warning mt-2">{Math.round(totalPoints / 7)}</p>
            </CardContent>
          </Card>
        </div>
      </div>

      {/* Filters */}
      <Card className="bg-gradient-card shadow-card border-border">
        <CardHeader>
          <CardTitle className="flex items-center gap-2 text-foreground">
            <Filter className="h-5 w-5" />
            Filtros
          </CardTitle>
        </CardHeader>
        <CardContent>
          <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
            <div className="space-y-2">
              <label className="text-sm font-medium text-foreground">Semana:</label>
              <Select value={selectedWeek} onValueChange={setSelectedWeek}>
                <SelectTrigger className="bg-secondary border-border">
                  <SelectValue />
                </SelectTrigger>
                <SelectContent>
                  <SelectItem value="todas">Todas</SelectItem>
                  <SelectItem value="1">Semana 1</SelectItem>
                  <SelectItem value="2">Semana 2</SelectItem>
                  <SelectItem value="3">Semana 3</SelectItem>
                  <SelectItem value="4">Semana 4</SelectItem>
                  <SelectItem value="5">Semana 5</SelectItem>
                </SelectContent>
              </Select>
            </div>
            
            <div className="space-y-2">
              <label className="text-sm font-medium text-foreground">Funcionário:</label>
              <Select value={selectedEmployee} onValueChange={setSelectedEmployee}>
                <SelectTrigger className="bg-secondary border-border">
                  <SelectValue />
                </SelectTrigger>
                <SelectContent>
                  <SelectItem value="todos">Todos</SelectItem>
                  {employees.map(employee => (
                    <SelectItem key={employee} value={employee}>{employee}</SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>
            
            <div className="space-y-2">
              <label className="text-sm font-medium text-foreground">Buscar:</label>
              <div className="relative">
                <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 h-4 w-4 text-muted-foreground" />
                <Input
                  placeholder="Buscar registros..."
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                  className="pl-10 bg-secondary border-border"
                />
              </div>
            </div>
            
            <div className="space-y-2">
              <label className="text-sm font-medium text-foreground">Ações:</label>
              <div className="flex gap-2">
                <Button 
                  variant="dashboard" 
                  size="sm" 
                  className="flex-1"
                  onClick={exportToExcel}
                  disabled={filteredRecords.length === 0}
                >
                  <Download className="h-4 w-4 mr-1" />
                  Exportar
                </Button>
              </div>
            </div>
          </div>
        </CardContent>
      </Card>

      {/* Records Table */}
      <Card className="bg-gradient-card shadow-card border-border">
        <CardHeader>
          <CardTitle className="flex items-center justify-between text-foreground">
            <div className="flex items-center gap-2">
              <Calendar className="h-5 w-5" />
              Seus Registros
            </div>
            <Badge variant="outline" className="text-dashboard-info border-dashboard-info/30">
              {filteredRecords.length} registros encontrados
            </Badge>
          </CardTitle>
        </CardHeader>
        <CardContent>
          <Table>
            <TableHeader>
              <TableRow className="border-border hover:bg-secondary/20">
                <TableHead className="text-foreground">Data</TableHead>
                <TableHead className="text-foreground">Horário</TableHead>
                <TableHead className="text-foreground">Funcionário</TableHead>
                <TableHead className="text-foreground">Refinaria</TableHead>
                <TableHead className="text-foreground">Pontos</TableHead>
                <TableHead className="text-foreground">Observações</TableHead>
                <TableHead className="text-foreground">Status</TableHead>
                <TableHead className="text-foreground">Ações</TableHead>
              </TableRow>
            </TableHeader>
            <TableBody>
              {loading ? (
                <TableRow>
                  <TableCell colSpan={8} className="text-center py-8 text-muted-foreground">
                    <Clock className="h-8 w-8 mx-auto mb-2 opacity-50 animate-spin" />
                    Carregando registros...
                  </TableCell>
                </TableRow>
              ) : filteredRecords.length === 0 ? (
                <TableRow>
                  <TableCell colSpan={8} className="text-center py-8 text-muted-foreground">
                    <MessageSquare className="h-8 w-8 mx-auto mb-2 opacity-50" />
                    Nenhum registro encontrado.
                  </TableCell>
                </TableRow>
              ) : (
                filteredRecords.map((record) => (
                  <TableRow key={record.id} className="border-border hover:bg-secondary/10">
                    <TableCell className="text-foreground">{record.date}</TableCell>
                    <TableCell className="text-muted-foreground">{record.time}</TableCell>
                    <TableCell className="font-medium text-foreground">{record.employee}</TableCell>
                    <TableCell className="text-foreground">{record.refinery}</TableCell>
                    <TableCell className="font-mono text-dashboard-primary font-bold">
                      {record.points.toLocaleString()}
                    </TableCell>
                    <TableCell className="text-muted-foreground max-w-xs truncate">
                      {record.observations}
                    </TableCell>
                    <TableCell>{getStatusBadge(record.status)}</TableCell>
                    <TableCell>
                      <div className="flex gap-1">
                        <Button variant="ghost" size="sm" className="h-8 w-8 p-0 hover:bg-dashboard-primary/20">
                          <Edit className="h-4 w-4 text-dashboard-primary" />
                        </Button>
                        <Button variant="ghost" size="sm" className="h-8 w-8 p-0 hover:bg-dashboard-danger/20">
                          <Trash2 className="h-4 w-4 text-dashboard-danger" />
                        </Button>
                      </div>
                    </TableCell>
                  </TableRow>
                ))
              )}
            </TableBody>
          </Table>
        </CardContent>
      </Card>
    </div>
  );
}