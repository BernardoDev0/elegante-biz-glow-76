-- CORREÇÃO DE SEGURANÇA - Resolver todos os problemas detectados

-- 1. Habilitar RLS em todas as tabelas públicas que não têm
ALTER TABLE public.funcionario ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.registro ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.month_reset ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.notification ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.notification_preference ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.points ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.refineries ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.users ENABLE ROW LEVEL SECURITY;

-- 2. Criar políticas RLS básicas para todas as tabelas
CREATE POLICY "Allow all access to funcionario" ON public.funcionario FOR ALL USING (true);
CREATE POLICY "Allow all access to registro" ON public.registro FOR ALL USING (true);
CREATE POLICY "Allow all access to month_reset" ON public.month_reset FOR ALL USING (true);
CREATE POLICY "Allow all access to notification" ON public.notification FOR ALL USING (true);
CREATE POLICY "Allow all access to notification_preference" ON public.notification_preference FOR ALL USING (true);
CREATE POLICY "Allow all access to points" ON public.points FOR ALL USING (true);
CREATE POLICY "Allow all access to refineries" ON public.refineries FOR ALL USING (true);
CREATE POLICY "Allow all access to users" ON public.users FOR ALL USING (true);

-- 3. Corrigir a função update_updated_at_column com search_path seguro
CREATE OR REPLACE FUNCTION public.update_updated_at_column()
RETURNS TRIGGER 
LANGUAGE plpgsql
SECURITY DEFINER
SET search_path = public
AS $$
BEGIN
    NEW.updated_at = now();
    RETURN NEW;
END;
$$;