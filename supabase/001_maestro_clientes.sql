-- Tabla privada para el maestro de clientes usado por el dashboard.
create table if not exists public.maestro_clientes (
    codigo bigint primary key,
    estado text,
    razon_social text,
    direccion text,
    barrio text,
    telefono text,
    localidad text,
    vendedor text,
    zona text,
    lunes text,
    martes text,
    miercoles text,
    jueves text,
    viernes text,
    sabado text,
    domingo text,
    frecuencia text,
    updated_at timestamptz not null default now()
);

create index if not exists idx_maestro_clientes_vendedor
    on public.maestro_clientes(vendedor);

alter table public.maestro_clientes enable row level security;

-- No se crean policies públicas. La app debe usar SUPABASE_SERVICE_ROLE_KEY
-- guardada únicamente en Streamlit Secrets para leer y actualizar este maestro.

-- Bucket privado para persistir el reporte diario y una copia de respaldo del maestro.
insert into storage.buckets (id, name, public)
values ('mi-negocio', 'mi-negocio', false)
on conflict (id) do update set public = excluded.public;
