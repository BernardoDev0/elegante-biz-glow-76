import { Button } from "@/components/ui/button";
import { Users, BarChart3, Table, FileSpreadsheet } from "lucide-react";

interface NavigationTabsProps {
  activeTab: string;
  onTabChange: (tab: string) => void;
}

const tabs = [
  { id: "equipe", label: "Equipe", icon: Users },
  { id: "graficos", label: "Gráficos", icon: BarChart3 },
  { id: "registros", label: "Registros", icon: Table },
  
];

export function NavigationTabs({ activeTab, onTabChange }: NavigationTabsProps) {
  return (
    <div className="grid grid-cols-2 lg:grid-cols-4 gap-4 mb-8">
      {tabs.map((tab) => {
        const Icon = tab.icon;
        const isActive = activeTab === tab.id;
        
        return (
          <Button
            key={tab.id}
            variant={isActive ? "dashboard-active" : "dashboard"}
            className="h-16 flex-col gap-2"
            onClick={() => onTabChange(tab.id)}
          >
            <Icon className="h-5 w-5" />
            <span className="text-sm font-medium">{tab.label}</span>
          </Button>
        );
      })}
    </div>
  );
}