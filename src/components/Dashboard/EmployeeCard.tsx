import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Progress } from "@/components/ui/progress";
import { Avatar, AvatarFallback } from "@/components/ui/avatar";
import { AlertTriangle, CheckCircle, Clock } from "lucide-react";

interface EmployeeCardProps {
  name: string;
  role?: string;
  weeklyPoints: number;
  weeklyGoal: number;
  monthlyPoints: number;
  monthlyGoal: number;
  weeklyProgress: number;
  monthlyProgress: number;
  status: "below" | "on-track" | "above";
}

export function EmployeeCard({
  name,
  role = "Funcionário",
  weeklyPoints,
  weeklyGoal,
  monthlyPoints,
  monthlyGoal,
  weeklyProgress,
  monthlyProgress,
  status
}: EmployeeCardProps) {
  const getInitials = (name: string) => {
    return name.split(' ').map(n => n[0]).join('').substring(0, 2).toUpperCase();
  };

  const getStatusConfig = () => {
    switch (status) {
      case "above":
        return {
          badge: "Acima da Meta",
          variant: "default" as const,
          icon: CheckCircle,
          color: "text-dashboard-success"
        };
      case "on-track":
        return {
          badge: "No Caminho",
          variant: "secondary" as const,
          icon: Clock,
          color: "text-dashboard-warning"
        };
      default:
        return {
          badge: "Abaixo da Meta",
          variant: "destructive" as const,
          icon: AlertTriangle,
          color: "text-dashboard-danger"
        };
    }
  };

  const statusConfig = getStatusConfig();
  const StatusIcon = statusConfig.icon;

  return (
    <Card className="bg-gradient-card shadow-card border-border">
      <CardHeader className="pb-3">
        <div className="flex items-center gap-3">
          <Avatar className="h-10 w-10 bg-dashboard-primary/20 border border-dashboard-primary/30">
            <AvatarFallback className="bg-dashboard-primary text-primary-foreground font-medium">
              {getInitials(name)}
            </AvatarFallback>
          </Avatar>
          <div className="flex-1">
            <CardTitle className="text-lg text-foreground">{name}</CardTitle>
            <p className="text-sm text-muted-foreground">{role}</p>
          </div>
          <Badge 
            variant={statusConfig.variant} 
            className={`gap-1 ${status === 'below' ? 'bg-red-500/20 backdrop-blur-sm border border-red-500/30 text-red-100' : ''}`}
          >
            <StatusIcon className="h-3 w-3" />
            {statusConfig.badge}
          </Badge>
        </div>
      </CardHeader>
      
      <CardContent className="space-y-4">
        <div className="grid grid-cols-2 gap-4">
          <div className="space-y-2">
            <div className="flex items-center gap-1">
              <div className="w-2 h-2 rounded-full bg-dashboard-info"></div>
              <span className="text-sm font-medium text-foreground">Semanal</span>
            </div>
            <div className="text-xl font-bold text-dashboard-info">{weeklyPoints}</div>
            <div className="text-xs text-muted-foreground">Meta: {weeklyGoal}</div>
            <Progress value={weeklyProgress} className="h-2" />
            <div className="text-xs text-muted-foreground">{weeklyProgress.toFixed(1)}% da meta semanal</div>
          </div>
          
          <div className="space-y-2">
            <div className="flex items-center gap-1">
              <div className="w-2 h-2 rounded-full bg-dashboard-secondary"></div>
              <span className="text-sm font-medium text-foreground">Mensal</span>
            </div>
            <div className="text-xl font-bold text-dashboard-secondary">{monthlyPoints}</div>
            <div className="text-xs text-muted-foreground">Meta: {monthlyGoal}</div>
            <Progress value={monthlyProgress} className="h-2" />
            <div className="text-xs text-muted-foreground">{monthlyProgress.toFixed(1)}% da meta mensal</div>
          </div>
        </div>
      </CardContent>
    </Card>
  );
}