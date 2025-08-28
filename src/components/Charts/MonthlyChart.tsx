import { 
  BarChart, 
  Bar,
  LineChart, 
  Line,
  XAxis, 
  YAxis, 
  CartesianGrid, 
  Tooltip, 
  ResponsiveContainer 
} from 'recharts';
import { ExcelProcessorService } from '@/services/ExcelProcessorService';

interface MonthlyChartProps {
  data: any[];
  hiddenEmployees: Set<string>;
  viewMode: "team" | "individual";
}

const employeeColors = {
  'Rodrigo': '#8b5cf6',
  'Maurício': '#f59e0b', 
  'Matheus': '#10b981',
  'Wesley': '#ef4444'
};

export function MonthlyChart({ data, hiddenEmployees, viewMode }: MonthlyChartProps) {
  if (viewMode === "individual") {
    // Modo individual: gráfico de barras por funcionário
    return (
      <ResponsiveContainer width="100%" height="100%">
        <BarChart data={data} margin={{ top: 20, right: 30, left: 20, bottom: 5 }}>
          <CartesianGrid strokeDasharray="3 3" stroke="hsl(var(--border))" opacity={0.3} />
          <XAxis 
            dataKey="name" 
            stroke="hsl(var(--muted-foreground))"
            fontSize={12}
            axisLine={false}
            tickLine={false}
          />
          <YAxis 
            stroke="hsl(var(--muted-foreground))"
            fontSize={12}
            axisLine={false}
            tickLine={false}
          />
          <Tooltip 
            contentStyle={{ 
              backgroundColor: 'hsl(var(--card))', 
              border: '1px solid hsl(var(--border))',
              borderRadius: '8px',
              boxShadow: '0 4px 12px hsl(var(--shadow) / 0.15)',
              color: 'hsl(var(--card-foreground))',
              fontSize: '14px'
            }}
            formatter={(value: any, name: string) => [
              `${value} pontos (${ExcelProcessorService.formatCurrency(value * 3.25)})`,
              name
            ]}
            cursor={{ fill: 'transparent' }}
            wrapperStyle={{ outline: 'none' }}
          />
          
          {Object.entries(employeeColors).map(([employee, color]) => (
            !hiddenEmployees.has(employee) && (
              <Bar 
                key={employee} 
                dataKey={employee} 
                fill={color} 
                name={employee} 
                radius={[2, 2, 0, 0]} 
              />
            )
          ))}
        </BarChart>
      </ResponsiveContainer>
    );
  } else {
    // Modo equipe: gráfico de linha com total
    const teamProgressData = data.map(month => {
      const totalPoints = Object.keys(employeeColors).reduce((sum, employee) => {
        if (!hiddenEmployees.has(employee)) {
          return sum + (month[employee] || 0);
        }
        return sum;
      }, 0);
      
      return {
        name: month.name,
        pontos: totalPoints,
        lucro: totalPoints * 3.25,
        meta: 29500 // Meta mensal da equipe
      };
    });

    return (
      <ResponsiveContainer width="100%" height="100%">
        <LineChart data={teamProgressData} margin={{ top: 20, right: 30, left: 20, bottom: 5 }}>
          <CartesianGrid strokeDasharray="3 3" stroke="hsl(var(--border))" opacity={0.3} />
          <XAxis 
            dataKey="name" 
            stroke="hsl(var(--muted-foreground))"
            fontSize={12}
            axisLine={false}
            tickLine={false}
          />
          <YAxis 
            stroke="hsl(var(--muted-foreground))"
            fontSize={12}
            axisLine={false}
            tickLine={false}
          />
          <Tooltip 
            contentStyle={{ 
              backgroundColor: 'hsl(var(--card))', 
              border: '1px solid hsl(var(--border))',
              borderRadius: '8px',
              boxShadow: '0 4px 12px hsl(var(--shadow) / 0.15)',
              color: 'hsl(var(--card-foreground))',
              fontSize: '14px'
            }}
            formatter={(value: any, name: string) => {
              if (name === 'pontos') {
                return [`${value} pontos`, 'Pontos Realizados'];
              } else if (name === 'lucro') {
                return [ExcelProcessorService.formatCurrency(value), 'Lucro Total'];
              }
              return [value, name];
            }}
            cursor={{ fill: 'transparent' }}
            wrapperStyle={{ outline: 'none' }}
          />
          <Line 
            type="monotone" 
            dataKey="pontos" 
            stroke="hsl(var(--dashboard-primary))" 
            strokeWidth={3}
            name="pontos"
            dot={{ fill: 'hsl(var(--dashboard-primary))', strokeWidth: 2, r: 4 }}
            activeDot={{ r: 6, stroke: 'hsl(var(--dashboard-primary))', strokeWidth: 2 }}
          />
          <Line 
            type="monotone" 
            dataKey="meta" 
            stroke="hsl(var(--dashboard-warning))" 
            strokeWidth={2}
            strokeDasharray="5 5"
            name="Meta"
            dot={{ fill: 'hsl(var(--dashboard-warning))', strokeWidth: 2, r: 4 }}
            activeDot={{ r: 6, stroke: 'hsl(var(--dashboard-warning))', strokeWidth: 2 }}
          />
        </LineChart>
      </ResponsiveContainer>
    );
  }
}