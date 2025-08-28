import { 
  BarChart, 
  Bar, 
  XAxis, 
  YAxis, 
  CartesianGrid, 
  Tooltip, 
  ResponsiveContainer 
} from 'recharts';
import { DataService } from '@/services/DataService';

interface WeeklyChartProps {
  data: any[];
  hiddenEmployees: Set<string>;
}

const employeeColors = {
  'Rodrigo': '#8b5cf6',
  'Maurício': '#f59e0b', 
  'Matheus': '#10b981',
  'Wesley': '#ef4444'
};

export function WeeklyChart({ data, hiddenEmployees }: WeeklyChartProps) {
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
          domain={[0, 3000]}
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
          cursor={{ fill: 'transparent' }}
          wrapperStyle={{ outline: 'none' }}
        />
        
        {/* Linha de Meta - 2375 pontos */}
        {data && data.length > 0 && (
          <g>
            <line 
              x1="10%" 
              y1={`${((3000 - 2375) / 3000) * 100}%`}
              x2="90%" 
              y2={`${((3000 - 2375) / 3000) * 100}%`}
              stroke="#f59e0b" 
              strokeWidth="2" 
              strokeDasharray="5 5"
              opacity={0.8}
            />
          </g>
        )}
        
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
}