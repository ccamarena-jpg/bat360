# BAT 360 — Plataforma de gestión de cuenta (piloto: Facturación y Gastos)

Plataforma para gestionar 360° la cuenta **BAT** (y **Palmera**). Mismo stack que tu RRHH:
**1 archivo HTML + Google Apps Script + Google Sheets**, pero en un Sheet/WebApp **nuevo e independiente**.

## Archivos

```
bat/
├── index.html                      ← la plataforma (ábrela en el navegador)
├── gas/Codigo.gs                   ← código para Apps Script
├── data/
│   ├── facturacion_historico.csv   ← 862 facturas limpias (Ene 2025 – Jun 2026)
│   └── limpiar_facturacion.py      ← script que generó el CSV (referencia)
└── SETUP.md                        ← esta guía
```

## Pasos para activar (una sola vez, ~10 min)

1. **Crea un Google Sheet nuevo** y llámalo `BAT - Facturación y Gastos`.
2. En el Sheet: **Extensiones → Apps Script**. Borra lo que haya y pega TODO el contenido de `gas/Codigo.gs`. Guarda.
3. En el editor de Apps Script, selecciona la función **`initSheets`** y dale **Ejecutar** (autoriza permisos la primera vez). Esto crea 3 pestañas: `Facturacion`, `Gastos`, `Maestros` (con los catálogos).
4. **Importa el histórico**: en el Sheet, abre la pestaña `Facturacion` → menú **Archivo → Importar → Subir** `data/facturacion_historico.csv` → opción **"Reemplazar datos en la hoja actual"** y separador **coma**.
5. **Publica la app**: en Apps Script → **Implementar → Nueva implementación** → tipo **Aplicación web** → *Ejecutar como:* **Yo** → *Quién tiene acceso:* **Cualquiera**. Copia la **URL `/exec`**.
6. Abre `index.html`, ve a **⚙️ Configuración**, pega la URL y dale **Guardar y conectar**.

¡Listo! Verás tus 862 facturas, el resumen por mes/servicio y el margen.

## Módulos del piloto

- **Resumen** — KPIs (facturado, gastos, margen, %), facturación por mes y por servicio.
- **Facturación** — tabla con filtros (cliente, servicio, estado, año, mes, búsqueda) + alta/edición/borrado.
- **Gastos** — registras costos por **Cliente + Servicio + Mes**, con grupo **Personas / Otros** y categoría de detalle.
- **Margen por servicio** — Facturado − (Gasto Personas + Otros) = Margen y % , por servicio y por mes.

## Estructura de datos

**Facturacion**: ID, Cliente, Ejecutivo, Tipo de SS (servicio), Responsable de Pago, Servicio/Proyecto, Mes, Importe, OS, Serie, # Factura, Fecha Factura, Vencimiento, Estado, Estado Detalle.

**Gastos**: ID, Cliente, Tipo de SS (servicio), Mes, Grupo (Personas/Otros), Categoría, Monto, Detalle.

**Maestros**: catálogos editables (Cliente, Servicio, Estado, Responsable, CatPersonas, CatOtros).

## Limpieza aplicada al histórico

- 27 filas vacías/basura descartadas → quedaron **862 registros**.
- "DIEGO LAZO" (que estaba como tipo de servicio) reclasificado a **TRANSPORTE**.
- Estados unificados a 5: **Pendiente / Programado / Factura enviada / Facturado / Pagado** (el texto original se conserva en *Estado Detalle*).
- Responsables de pago deduplicados (NICO/NICOLE AGUIRRE, etc.).
- Cliente "DP" se mantuvo tal cual.

## Próximos módulos (cuando valides el piloto)

Promotoría (Oxxo, Tambo, Primax, Repsol, Ava, Rappi · semáforos, cobertura, performance), Merch (OTs, cobertura) y Calendario por responsable.
