import { Users, BarChart3, Table, TrendingUp } from "lucide-react";
import { useLocation, useNavigate } from "react-router-dom";

import {
  Sidebar,
  SidebarContent,
  SidebarGroup,
  SidebarGroupContent,
  SidebarGroupLabel,
  SidebarMenu,
  SidebarMenuButton,
  SidebarMenuItem,
  SidebarHeader,
} from "@/components/ui/sidebar";

const items = [
  { title: "Equipe", url: "/admin", icon: Users },
  { title: "Gráficos", url: "/admin/graficos", icon: BarChart3 },
  { title: "Registros", url: "/admin/registros", icon: Table },
  
];

export function AppSidebar() {
  const location = useLocation();
  const navigate = useNavigate();
  const currentPath = location.pathname;

  const isActive = (path: string) => currentPath === path;

  return (
    <Sidebar
      className="bg-card border-border overflow-hidden"
      collapsible="icon"
    >
      <SidebarHeader className="p-2 border-b border-border">
        <div className="flex items-center gap-1">
          <div className="p-1.5 rounded-lg bg-white/10 backdrop-blur-sm border border-white/20">
            <TrendingUp className="h-5 w-5 text-white" />
          </div>
          <div className="group-data-[collapsible=icon]:hidden">
            <h2 className="text-lg font-bold bg-gradient-primary bg-clip-text text-transparent">
              Dashboard CEO
            </h2>
            <p className="text-xs text-muted-foreground">Sistema de Monitoramento</p>
          </div>
        </div>
      </SidebarHeader>

      <SidebarContent>
        <SidebarGroup>
          <SidebarGroupLabel>Navegação</SidebarGroupLabel>
          <SidebarGroupContent>
            <SidebarMenu>
              {items.map((item) => {
                const Icon = item.icon;
                const active = isActive(item.url);
                
                return (
                  <SidebarMenuItem key={item.title}>
                    <SidebarMenuButton
                      onClick={() => navigate(item.url)}
                      className={`
                        ${active 
                          ? 'bg-gradient-primary text-primary-foreground shadow-glow' 
                          : 'hover:bg-secondary/50 text-foreground hover:text-dashboard-primary'
                        }
                        transition-colors h-12
                      `}
                    >
                      <Icon className="h-5 w-5" />
                      <span className="font-medium">{item.title}</span>
                    </SidebarMenuButton>
                  </SidebarMenuItem>
                );
              })}
            </SidebarMenu>
          </SidebarGroupContent>
        </SidebarGroup>
      </SidebarContent>
    </Sidebar>
  );
}