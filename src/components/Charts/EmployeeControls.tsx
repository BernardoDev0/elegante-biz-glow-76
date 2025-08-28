import { Button } from "@/components/ui/button";
import { Badge } from "@/components/ui/badge";
import { Eye, EyeOff, Users, User } from "lucide-react";

interface EmployeeControlsProps {
  viewMode: "team" | "individual";
  onViewModeChange: (mode: "team" | "individual") => void;
  hiddenEmployees: Set<string>;
  onToggleEmployee: (employee: string) => void;
  selectedChart: string;
}

const employeeColors = {
  'Rodrigo': '#8b5cf6',
  'Maurício': '#f59e0b', 
  'Matheus': '#10b981',
  'Wesley': '#ef4444'
};

export function EmployeeControls({ 
  viewMode, 
  onViewModeChange, 
  hiddenEmployees, 
  onToggleEmployee,
  selectedChart 
}: EmployeeControlsProps) {
  return (
    <div className="space-y-2">
      {Object.entries(employeeColors).map(([employee, color]) => (
        <Button
          key={employee}
          variant="ghost"
          onClick={() => onToggleEmployee(employee)}
          className={`w-full justify-start h-8 text-xs transition-all duration-200 ${
            hiddenEmployees.has(employee) 
              ? 'opacity-50 text-muted-foreground hover:opacity-70' 
              : 'text-foreground hover:bg-primary/10 hover:text-primary hover:scale-[1.02]'
          }`}
          size="sm"
        >
          {hiddenEmployees.has(employee) ? (
            <EyeOff className="h-3 w-3 mr-2" />
          ) : (
            <Eye className="h-3 w-3 mr-2" />
          )}
          <div 
            className="w-2 h-2 rounded-full mr-2 flex-shrink-0" 
            style={{ backgroundColor: color }}
          />
          <span className="truncate">{employee}</span>
        </Button>
      ))}
      <p className="text-xs text-muted-foreground pt-1">
        Clique para mostrar/ocultar
      </p>
    </div>
  );
}