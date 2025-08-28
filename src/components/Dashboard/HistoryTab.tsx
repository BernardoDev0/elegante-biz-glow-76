import { useState, useEffect } from "react";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { Badge } from "@/components/ui/badge";
import { Input } from "@/components/ui/input";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Button } from "@/components/ui/button";
import { Calendar, MapPin, MessageSquare, TrendingUp, Search, Filter } from "lucide-react";
import { EmployeeService, Entry } from "@/services/EmployeeService";
import { format } from "date-fns";
import { ptBR } from "date-fns/locale";

interface HistoryTabProps {
  employeeId: number;
}

export const HistoryTab = ({ employeeId }: HistoryTabProps) => {
  const [entries, setEntries] = useState<Entry[]>([]);
  const [loading, setLoading] = useState(true);
  const [searchTerm, setSearchTerm] = useState("");
  const [filterType, setFilterType] = useState<"all" | "week" | "month">("all");
  const [page, setPage] = useState(0);
  const [hasMore, setHasMore] = useState(true);

  const ITEMS_PER_PAGE = 10;

  const loadEntries = async (reset = false) => {
    try {
      setLoading(true);
      const currentPage = reset ? 0 : page;
      
      let dateFilter: { start?: string; end?: string } | undefined;
      const now = new Date();
      
      if (filterType === "week") {
        const weekStart = new Date(now);
        weekStart.setDate(now.getDate() - now.getDay());
        dateFilter = { start: weekStart.toISOString().split('T')[0] };
      } else if (filterType === "month") {
        const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
        dateFilter = { start: monthStart.toISOString().split('T')[0] };
      }
      
      const newEntries = await EmployeeService.getEmployeeEntries(
        employeeId,
        ITEMS_PER_PAGE,
        currentPage * ITEMS_PER_PAGE,
        dateFilter
      );
      
      if (reset) {
        setEntries(newEntries);
      } else {
        setEntries(prev => [...prev, ...newEntries]);
      }
      
      setHasMore(newEntries.length === ITEMS_PER_PAGE);
      if (reset) setPage(0);
      else setPage(prev => prev + 1);
      
    } catch (error) {
      console.error('Erro ao carregar histórico:', error);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadEntries(true);
  }, [employeeId, filterType]);

  const filteredEntries = entries.filter(entry =>
    entry.observations.toLowerCase().includes(searchTerm.toLowerCase()) ||
    entry.refinery.toLowerCase().includes(searchTerm.toLowerCase())
  );

  const formatDate = (dateString: string) => {
    try {
      return format(new Date(dateString), "dd/MM/yyyy", { locale: ptBR });
    } catch {
      return dateString;
    }
  };

  const formatTime = (dateString: string) => {
    try {
      return format(new Date(dateString), "HH:mm", { locale: ptBR });
    } catch {
      return "";
    }
  };

  const getPointsBadgeVariant = (points: number) => {
    if (points >= 500) return "default";
    if (points >= 300) return "secondary";
    return "outline";
  };

  return (
    <Card className="bg-gradient-card border-border/50 shadow-card">
      <CardHeader>
        <CardTitle className="flex items-center gap-2">
          <Calendar className="w-5 h-5" />
          Histórico de Registros
        </CardTitle>
        
        {/* Filtros */}
        <div className="flex flex-col sm:flex-row gap-4 mt-4">
          <div className="relative flex-1">
            <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 text-muted-foreground w-4 h-4" />
            <Input
              placeholder="Buscar por observações ou refinaria..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="pl-10"
            />
          </div>
          
          <Select value={filterType} onValueChange={(value: "all" | "week" | "month") => setFilterType(value)}>
            <SelectTrigger className="w-full sm:w-40">
              <Filter className="w-4 h-4 mr-2" />
              <SelectValue />
            </SelectTrigger>
            <SelectContent>
              <SelectItem value="all">Todos</SelectItem>
              <SelectItem value="week">Esta Semana</SelectItem>
              <SelectItem value="month">Este Mês</SelectItem>
            </SelectContent>
          </Select>
        </div>
      </CardHeader>
      
      <CardContent>
        {loading && entries.length === 0 ? (
          <div className="flex items-center justify-center py-8">
            <div className="text-center">
              <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-primary mx-auto mb-4"></div>
              <p className="text-muted-foreground">Carregando histórico...</p>
            </div>
          </div>
        ) : filteredEntries.length === 0 ? (
          <div className="text-center py-8">
            <Calendar className="w-12 h-12 text-muted-foreground mx-auto mb-4" />
            <h3 className="text-lg font-medium text-foreground mb-2">
              Nenhum registro encontrado
            </h3>
            <p className="text-muted-foreground">
              {searchTerm || filterType !== "all" 
                ? "Tente ajustar os filtros de busca"
                : "Ainda não há registros para este funcionário"
              }
            </p>
          </div>
        ) : (
          <>
            <div className="rounded-lg border border-border overflow-hidden">
              <Table>
                <TableHeader>
                  <TableRow className="bg-muted/50">
                    <TableHead className="font-semibold">Data</TableHead>
                    <TableHead className="font-semibold">Hora</TableHead>
                    <TableHead className="font-semibold">Refinaria</TableHead>
                    <TableHead className="font-semibold">Pontos</TableHead>
                    <TableHead className="font-semibold">Observações</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {filteredEntries.map((entry) => (
                    <TableRow key={entry.id} className="hover:bg-muted/30">
                      <TableCell className="font-medium">
                        <div className="flex items-center gap-2">
                          <Calendar className="w-4 h-4 text-muted-foreground" />
                          {formatDate(entry.date)}
                        </div>
                      </TableCell>
                      <TableCell>
                        <div className="flex items-center gap-2 text-muted-foreground">
                          {formatTime(entry.date)}
                        </div>
                      </TableCell>
                      <TableCell>
                        <div className="flex items-center gap-2">
                          <MapPin className="w-4 h-4 text-muted-foreground" />
                          <span className="font-medium">{entry.refinery}</span>
                        </div>
                      </TableCell>
                      <TableCell>
                        <Badge 
                          variant={getPointsBadgeVariant(entry.points)}
                          className="flex items-center gap-1"
                        >
                          <TrendingUp className="w-3 h-3" />
                          {entry.points}
                        </Badge>
                      </TableCell>
                      <TableCell className="max-w-xs">
                        <div className="flex items-start gap-2">
                          <MessageSquare className="w-4 h-4 text-muted-foreground mt-0.5" />
                          <span className="text-sm break-words">{entry.observations}</span>
                        </div>
                      </TableCell>
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            </div>
            
            {hasMore && (
              <div className="mt-6 text-center">
                <Button 
                  variant="outline" 
                  onClick={() => loadEntries(false)}
                  disabled={loading}
                  className="min-w-32"
                >
                  {loading ? 'Carregando...' : 'Carregar mais'}
                </Button>
              </div>
            )}
          </>
        )}
      </CardContent>
    </Card>
  );
};