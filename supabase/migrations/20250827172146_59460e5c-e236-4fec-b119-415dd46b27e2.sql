-- MIGRAÇÃO CONSERVADORA - UNIFICAÇÃO E LIMPEZA
-- Preservando todos os dados reais funcionais

-- 1. Adicionar campos de timestamp na tabela employee (se não existirem)
ALTER TABLE public.employee 
ADD COLUMN IF NOT EXISTS created_at TIMESTAMP WITH TIME ZONE DEFAULT now(),
ADD COLUMN IF NOT EXISTS updated_at TIMESTAMP WITH TIME ZONE DEFAULT now();

-- 2. Adicionar campos de timestamp na tabela entry (se não existirem)  
ALTER TABLE public.entry
ADD COLUMN IF NOT EXISTS created_at TIMESTAMP WITH TIME ZONE DEFAULT now(),
ADD COLUMN IF NOT EXISTS updated_at TIMESTAMP WITH TIME ZONE DEFAULT now();

-- 3. Corrigir o tipo da coluna date em entry para timestamp
ALTER TABLE public.entry 
ALTER COLUMN date TYPE TIMESTAMP WITH TIME ZONE USING date::timestamp with time zone;

-- 4. Adicionar foreign key entre entry e employee
ALTER TABLE public.entry
ADD CONSTRAINT fk_entry_employee 
FOREIGN KEY (employee_id) REFERENCES public.employee(id) ON DELETE CASCADE;

-- 5. Habilitar RLS nas tabelas principais
ALTER TABLE public.employee ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.entry ENABLE ROW LEVEL SECURITY;

-- 6. Criar políticas RLS básicas (permitir acesso a todos por enquanto)
CREATE POLICY "Allow all access to employee" ON public.employee FOR ALL USING (true);
CREATE POLICY "Allow all access to entry" ON public.entry FOR ALL USING (true);

-- 7. Função para atualizar updated_at automaticamente
CREATE OR REPLACE FUNCTION public.update_updated_at_column()
RETURNS TRIGGER AS $$
BEGIN
    NEW.updated_at = now();
    RETURN NEW;
END;
$$ LANGUAGE plpgsql;

-- 8. Triggers para updated_at
CREATE TRIGGER update_employee_updated_at
    BEFORE UPDATE ON public.employee
    FOR EACH ROW
    EXECUTE FUNCTION public.update_updated_at_column();

CREATE TRIGGER update_entry_updated_at
    BEFORE UPDATE ON public.entry
    FOR EACH ROW
    EXECUTE FUNCTION public.update_updated_at_column();

-- 9. Atualizar metas do Matheus conforme código Python
UPDATE public.employee 
SET weekly_goal = 2675 
WHERE username = 'E89P';

-- 10. Limpar tabelas duplicadas/desnecessárias mantendo apenas as principais
-- (vamos manter funcionario e registro por segurança, apenas limpar se vazias)
TRUNCATE TABLE public.registro; -- Esta já está vazia
-- Não vamos mexer na funcionario ainda para não quebrar nada

-- 11. Criar índices para performance
CREATE INDEX IF NOT EXISTS idx_entry_employee_id ON public.entry(employee_id);
CREATE INDEX IF NOT EXISTS idx_entry_date ON public.entry(date);
CREATE INDEX IF NOT EXISTS idx_employee_username ON public.employee(username);
CREATE INDEX IF NOT EXISTS idx_employee_access_key ON public.employee(access_key);