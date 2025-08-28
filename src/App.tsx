import { Toaster } from "@/components/ui/toaster";
import { Toaster as Sonner } from "@/components/ui/sonner";
import { TooltipProvider } from "@/components/ui/tooltip";
import { SidebarProvider, SidebarTrigger } from "@/components/ui/sidebar";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import { BrowserRouter, Routes, Route } from "react-router-dom";
import { AppSidebar } from "@/components/AppSidebar";
import Index from "./pages/Index";
import Graficos from "./pages/Graficos";
import Registros from "./pages/Registros";

import NotFound from "./pages/NotFound";
import Login from "./pages/Login";
import Dashboard from "./pages/Dashboard";

const queryClient = new QueryClient();

const App = () => (
  <QueryClientProvider client={queryClient}>
    <TooltipProvider>
      <Toaster />
      <Sonner />
      <BrowserRouter>
        <Routes>
          {/* Rotas sem sidebar - Login e Dashboard de funcionário */}
          <Route path="/login" element={<Login />} />
          <Route path="/dashboard" element={<Dashboard />} />
          
          {/* Rotas com sidebar - CEO Dashboard */}
          <Route path="/admin/*" element={
            <SidebarProvider>
              <div className="min-h-screen flex w-full bg-background">
                <AppSidebar />
                <main className="flex-1">
                  <header className="h-12 flex items-center border-b border-border bg-card/50 backdrop-blur-sm">
                    <SidebarTrigger className="ml-4" />
                  </header>
                  <div className="p-6">
                    <Routes>
                      <Route path="/" element={<Index />} />
                      <Route path="/graficos" element={<Graficos />} />
                      <Route path="/registros" element={<Registros />} />
                      
                    </Routes>
                  </div>
                </main>
              </div>
            </SidebarProvider>
          } />
          
          {/* Redirect root to login */}
          <Route path="/" element={<Login />} />
          <Route path="*" element={<NotFound />} />
        </Routes>
      </BrowserRouter>
    </TooltipProvider>
  </QueryClientProvider>
);

export default App;
