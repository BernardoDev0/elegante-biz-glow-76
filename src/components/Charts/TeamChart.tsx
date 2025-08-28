import { PieChart, Pie, Cell, Tooltip, ResponsiveContainer } from 'recharts';

interface TeamChartProps {
  data: any[];
}

export function TeamChart({ data }: TeamChartProps) {
  return (
    <ResponsiveContainer width="100%" height="100%">
      <PieChart>
        <Pie
          data={data}
          cx="50%"
          cy="50%"
          labelLine={false}
          label={({ name, value, percent }) => `${name}: ${value} (${(percent * 100).toFixed(0)}%)`}
          outerRadius={90}
          fill="#8884d8"
          dataKey="value"
        >
          {data.map((entry, index) => (
            <Cell key={`cell-${index}`} fill={entry.color} />
          ))}
        </Pie>
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
      </PieChart>
    </ResponsiveContainer>
  );
}