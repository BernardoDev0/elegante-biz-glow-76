import { TrendingUp } from "lucide-react";

export function Header() {
  return (
    <div className="text-center mb-8">
      <div className="inline-flex items-center gap-3 mb-4">
        <div className="p-2 rounded-lg bg-gradient-primary shadow-glow">
          <TrendingUp className="h-8 w-8 text-primary-foreground" />
        </div>
        <h1 className="text-3xl lg:text-4xl font-bold bg-gradient-primary bg-clip-text text-transparent">
          Dashboard CEO
        </h1>
      </div>
      <p className="text-lg text-muted-foreground">
        Visão geral da performance da equipe
      </p>
    </div>
  );
}