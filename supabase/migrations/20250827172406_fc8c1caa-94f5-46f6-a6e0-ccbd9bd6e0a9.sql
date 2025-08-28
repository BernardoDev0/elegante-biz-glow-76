-- Corrigir último problema de RLS - tabela alembic_version
ALTER TABLE public.alembic_version ENABLE ROW LEVEL SECURITY;
CREATE POLICY "Allow all access to alembic_version" ON public.alembic_version FOR ALL USING (true);