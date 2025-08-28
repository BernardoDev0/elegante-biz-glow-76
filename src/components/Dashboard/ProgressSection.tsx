import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Progress } from "@/components/ui/progress";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";

interface ProgressSectionProps {
  totalPoints: number;
  totalGoal: number;
  completedPercentage: number;
  selectedWeek: string;
  onWeekChange: (week: string) => void;
}

export function ProgressSection({ 
  totalPoints, 
  totalGoal, 
  completedPercentage, 
  selectedWeek, 
  onWeekChange 
}: ProgressSectionProps) {
  return (
    <Card className="bg-gradient-card shadow-card border-border mb-6">
      <CardHeader>
        <div className="flex items-center justify-between">
          <CardTitle className="text-lg font-semibold text-foreground flex items-center gap-2">
            <div className="w-4 h-4 bg-gradient-primary rounded"></div>
            Progresso Geral da Equipe - Mensal
          </CardTitle>
          <div className="flex items-center gap-4 text-sm">
            <span className="text-dashboard-info font-medium">
              {totalPoints} de {totalGoal.toLocaleString()} pontos
            </span>
            <span className="text-dashboard-danger font-bold">
              {completedPercentage.toFixed(1)}% concluído
            </span>
          </div>
        </div>
      </CardHeader>
      <CardContent>
        <div className="space-y-4">
          <Progress 
            value={completedPercentage} 
            className="h-3 bg-secondary"
          />
          
          <div className="flex items-center gap-2">
            <span className="text-sm text-muted-foreground">Selecionar Semana:</span>
            <Select value={selectedWeek} onValueChange={onWeekChange}>
              <SelectTrigger className="w-32 bg-secondary border-border">
                <SelectValue />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="1">Semana 1</SelectItem>
                <SelectItem value="2">Semana 2</SelectItem>
                <SelectItem value="3">Semana 3</SelectItem>
                <SelectItem value="4">Semana 4</SelectItem>
                <SelectItem value="5">Semana 5</SelectItem>
              </SelectContent>
            </Select>
          </div>
        </div>
      </CardContent>
    </Card>
  );
}