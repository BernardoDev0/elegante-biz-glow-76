import { supabase } from '@/integrations/supabase/client';

export interface Employee {
  id: number;
  name: string;
  real_name: string;
  username: string;
  access_key: string;
  role: string;
  weekly_goal: number;
  default_refinery: string;
  created_at?: string;
  updated_at?: string;
}

export interface Entry {
  id: number;
  employee_id: number;
  date: string;
  refinery: string;
  points: number;
  observations: string;
  created_at?: string;
  updated_at?: string;
}

export class EmployeeService {
  // Autenticação por chave de acesso (mantendo lógica original)
  static async authenticateByAccessKey(accessKey: string): Promise<Employee | null> {
    try {
      const { data, error } = await supabase
        .from('employee')
        .select('*')
        .eq('access_key', accessKey)
        .single();

      if (error || !data) {
        console.error('Erro na autenticação:', error);
        return null;
      }

      return data as Employee;
    } catch (error) {
      console.error('Erro na autenticação:', error);
      return null;
    }
  }

  // Buscar funcionário por ID
  static async getEmployeeById(id: number): Promise<Employee | null> {
    try {
      const { data, error } = await supabase
        .from('employee')
        .select('*')
        .eq('id', id)
        .single();

      if (error || !data) {
        console.error('Erro ao buscar funcionário:', error);
        return null;
      }

      return data as Employee;
    } catch (error) {
      console.error('Erro ao buscar funcionário:', error);
      return null;
    }
  }

  // Buscar todos os funcionários
  static async getAllEmployees(): Promise<Employee[]> {
    try {
      const { data, error } = await supabase
        .from('employee')
        .select('*')
        .order('real_name');

      if (error) {
        console.error('Erro ao buscar funcionários:', error);
        return [];
      }

      return data as Employee[];
    } catch (error) {
      console.error('Erro ao buscar funcionários:', error);
      return [];
    }
  }

  // Registrar novo ponto
  static async createEntry(entry: Omit<Entry, 'id' | 'created_at' | 'updated_at'>): Promise<Entry | null> {
    try {
      const { data, error } = await supabase
        .from('entry')
        .insert([entry])
        .select()
        .single();

      if (error || !data) {
        console.error('Erro ao registrar pontos:', error);
        return null;
      }

      return data as Entry;
    } catch (error) {
      console.error('Erro ao registrar pontos:', error);
      return null;
    }
  }

  // Buscar entradas do funcionário
  static async getEmployeeEntries(
    employeeId: number, 
    limit?: number, 
    offset?: number,
    dateFilter?: { start?: string; end?: string }
  ): Promise<Entry[]> {
    try {
      let query = supabase
        .from('entry')
        .select('*')
        .eq('employee_id', employeeId)
        .order('date', { ascending: false });

      if (dateFilter?.start) {
        query = query.gte('date', dateFilter.start);
      }
      if (dateFilter?.end) {
        query = query.lte('date', dateFilter.end);
      }
      if (limit) {
        query = query.limit(limit);
      }
      if (offset) {
        query = query.range(offset, offset + (limit || 10) - 1);
      }

      const { data, error } = await query;

      if (error) {
        console.error('Erro ao buscar entradas:', error);
        return [];
      }

      return data as Entry[];
    } catch (error) {
      console.error('Erro ao buscar entradas:', error);
      return [];
    }
  }

  // Calcular pontos do dia
  static async getTodayPoints(employeeId: number): Promise<number> {
    try {
      const today = new Date().toISOString().split('T')[0];
      
      const { data, error } = await supabase
        .from('entry')
        .select('points')
        .eq('employee_id', employeeId)
        .gte('date', today)
        .lt('date', `${today}T23:59:59`);

      if (error) {
        console.error('Erro ao calcular pontos do dia:', error);
        return 0;
      }

      return data?.reduce((total, entry) => total + (entry.points || 0), 0) || 0;
    } catch (error) {
      console.error('Erro ao calcular pontos do dia:', error);
      return 0;
    }
  }

  // Calcular pontos da semana (baseado na lógica 26→25)
  static async getWeekPoints(employeeId: number, weekDates: { start: string; end: string }): Promise<number> {
    try {
      const { data, error } = await supabase
        .from('entry')
        .select('points')
        .eq('employee_id', employeeId)
        .gte('date', weekDates.start)
        .lte('date', weekDates.end);

      if (error) {
        console.error('Erro ao calcular pontos da semana:', error);
        return 0;
      }

      return data?.reduce((total, entry) => total + (entry.points || 0), 0) || 0;
    } catch (error) {
      console.error('Erro ao calcular pontos da semana:', error);
      return 0;
    }
  }

  // Calcular pontos mensais (baseado na lógica 26→25)
  static async getMonthPoints(employeeId: number, monthDates: { start: string; end: string }): Promise<number> {
    try {
      const { data, error } = await supabase
        .from('entry')
        .select('points')
        .eq('employee_id', employeeId)
        .gte('date', monthDates.start)
        .lte('date', monthDates.end);

      if (error) {
        console.error('Erro ao calcular pontos mensais:', error);
        return 0;
      }

      return data?.reduce((total, entry) => total + (entry.points || 0), 0) || 0;
    } catch (error) {
      console.error('Erro ao calcular pontos mensais:', error);
      return 0;
    }
  }
}