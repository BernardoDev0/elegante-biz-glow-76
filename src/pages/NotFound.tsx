import { useLocation } from "react-router-dom";
import { useEffect } from "react";
import { Button } from "@/components/ui/button";
import { Home, AlertCircle } from "lucide-react";

const NotFound = () => {
  const location = useLocation();

  useEffect(() => {
    console.error(
      "404 Error: User attempted to access non-existent route:",
      location.pathname
    );
  }, [location.pathname]);

  return (
    <div className="min-h-screen flex items-center justify-center bg-background">
      <div className="text-center space-y-6">
        <div className="flex justify-center">
          <div className="p-4 rounded-full bg-dashboard-danger/10 border border-dashboard-danger/20">
            <AlertCircle className="h-16 w-16 text-dashboard-danger" />
          </div>
        </div>
        
        <div className="space-y-2">
          <h1 className="text-6xl font-bold text-foreground">404</h1>
          <h2 className="text-2xl font-semibold text-foreground">Página não encontrada</h2>
          <p className="text-muted-foreground max-w-md">
            Oops! A página que você está procurando não existe ou foi movida.
          </p>
        </div>
        
        <Button 
          asChild 
          variant="dashboard-active" 
          className="gap-2"
        >
          <a href="/">
            <Home className="h-4 w-4" />
            Voltar ao Dashboard
          </a>
        </Button>
      </div>
    </div>
  );
};

export default NotFound;
