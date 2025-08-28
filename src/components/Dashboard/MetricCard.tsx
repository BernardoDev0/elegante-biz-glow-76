import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { TrendingUp, Target, Users, BarChart3 } from "lucide-react";

interface MetricCardProps {
  title: string;
  value: string | number;
  subtitle?: string;
  icon?: "trending" | "target" | "users" | "chart";
  variant?: "default" | "primary" | "success" | "warning";
}

const iconMap = {
  trending: TrendingUp,
  target: Target,
  users: Users,
  chart: BarChart3,
};

export function MetricCard({ title, value, subtitle, icon, variant = "default" }: MetricCardProps) {
  const Icon = icon ? iconMap[icon] : null;
  
  const getVariantStyles = () => {
    switch (variant) {
      case "primary":
        return "bg-gradient-primary/20 backdrop-blur-sm border border-primary/30 shadow-glow";
      case "success":
        return "bg-gradient-to-br from-dashboard-success/10 to-dashboard-success/5 border-dashboard-success/30";
      case "warning":
        return "bg-gradient-to-br from-dashboard-warning/10 to-dashboard-warning/5 border-dashboard-warning/30";
      default:
        return "bg-gradient-card shadow-card border-border";
    }
  };

  const getIconColor = () => {
    switch (variant) {
      case "primary":
        return "text-primary-foreground";
      case "success":
        return "text-dashboard-success";
      case "warning":
        return "text-dashboard-warning";
      default:
        return "text-dashboard-primary";
    }
  };

  return (
    <Card className={`border ${getVariantStyles()}`}>
      <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2">
        <CardTitle className="text-sm font-medium text-muted-foreground">
          {title}
        </CardTitle>
        {Icon && (
          <Icon className={`h-4 w-4 ${getIconColor()}`} />
        )}
      </CardHeader>
      <CardContent>
        <div className="text-2xl font-bold text-foreground mb-1">
          {value}
        </div>
        {subtitle && (
          <p className="text-xs text-muted-foreground">
            {subtitle}
          </p>
        )}
      </CardContent>
    </Card>
  );
}