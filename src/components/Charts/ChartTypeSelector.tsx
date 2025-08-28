import { Button } from "@/components/ui/button";
import { BarChart3, TrendingUp, Target } from "lucide-react";

interface ChartTypeSelectorProps {
  selectedChart: string;
  onChartChange: (chart: string) => void;
}

export function ChartTypeSelector({ selectedChart, onChartChange }: ChartTypeSelectorProps) {
  const chartTypes = [
    { id: "weekly", label: "Progresso Semanal", icon: BarChart3 },
    { id: "progress", label: "Progresso Mensal", icon: TrendingUp },
    { id: "team", label: "Distribuição da Equipe", icon: Target }
  ];

  return (
    <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
      {chartTypes.map(({ id, label, icon: Icon }) => (
        <Button
          key={id}
          variant={selectedChart === id ? "dashboard-active" : "dashboard"}
          onClick={() => onChartChange(id)}
          className="h-16 flex-col gap-2"
        >
          <Icon className="h-5 w-5" />
          <span className="text-sm">{label}</span>
        </Button>
      ))}
    </div>
  );
}