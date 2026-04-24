# USD Tracker

App para registro profesional de compra y venta de dólares con sincronización en la nube.

## Stack
- React 18
- Supabase (base de datos + realtime sync)
- Vercel (hosting)
- PWA (instalable en celular)

## Setup

### 1. Base de datos Supabase
Ejecutá este SQL en tu proyecto de Supabase:

```sql
create table personas (
  id uuid default gen_random_uuid() primary key,
  nombre text not null unique,
  created_at timestamp default now()
);

create table operaciones (
  id uuid default gen_random_uuid() primary key,
  persona text not null,
  banco text not null,
  usd numeric not null,
  tc numeric not null,
  ars numeric not null,
  fecha date not null,
  nota text,
  tc_venta numeric,
  fecha_venta date,
  usd_vendido numeric,
  lote_id text,
  created_at timestamp default now()
);
```

### 2. Variables de entorno
Creá un archivo `.env` en la raíz con:
```
REACT_APP_SUPABASE_URL=https://jqmdokvsmeundmsjkjsl.supabase.co
REACT_APP_SUPABASE_KEY=sb_publishable_tOx6zCi0idR6BRqiTkXi6w_X8qJL2QC
```

### 3. Instalar y correr local
```bash
npm install
npm start
```

### 4. Deploy en Vercel
1. Subí este repositorio a GitHub
2. Importalo en vercel.com
3. Agregá las variables de entorno en Vercel (Settings → Environment Variables)
4. Deploy

### 5. Instalar como app en el celular
- Android: Chrome → menú → "Agregar a pantalla de inicio"
- iPhone: Safari → compartir → "Agregar a pantalla de inicio"

## Funcionalidades
- Registro de compras por persona y banco
- Venta grupal con TC promedio ponderado automático
- Venta parcial (ej: vender 600 de 3000 USD)
- Sincronización en tiempo real entre celular y PC
- Exportación a Excel (.xlsx)
- Historial filtrable por persona y mes
- Resumen por persona con stock y ganancia
