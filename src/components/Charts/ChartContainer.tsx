import { ReactNode } from "react";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";

interface ChartContainerProps {
  title: string;
  icon: ReactNode;
  children: ReactNode;
  className?: string;
}

export function ChartContainer({ title, icon, children, className = "" }: ChartContainerProps) {
  return (
    <Card className={`bg-gradient-card shadow-card border-border ${className}`}>
      <CardHeader>
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-3">
            {icon}
            <CardTitle className="text-foreground">{title}</CardTitle>
          </div>
        </div>
      </CardHeader>
      <CardContent className="pt-0 pb-4">
        <div className="h-[380px]">
          {children}
        </div>
      </CardContent>
    </Card>
  );
}