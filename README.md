# 📦 Dashboard — Estatus de Back Orders

Dashboard interactivo para el seguimiento de órdenes de compra y back orders.

## 🚀 Cómo subir a GitHub Pages

### Paso 1 — Sube los archivos al repositorio
Sube **ambos archivos** a tu repositorio de GitHub:
```
📁 tu-repositorio/
   ├── index.html              ← renombra dashboard_ordenes.html a index.html
   └── Estatus_de_BO.xlsx      ← tu archivo Excel (mismo nombre, mismo directorio)
```

### Paso 2 — Activa GitHub Pages
1. Ve a tu repositorio → **Settings** → **Pages**
2. En *Source* selecciona **Deploy from a branch**
3. Elige la rama `main` y carpeta `/ (root)`
4. Guarda → en unos minutos tendrás tu URL del tipo `https://tuusuario.github.io/tu-repo/`

### Paso 3 — Actualizar los datos
Cuando tengas un nuevo reporte:
1. Reemplaza `Estatus_de_BO.xlsx` en el repositorio con el nuevo archivo  
2. El dashboard se actualiza automáticamente al recargar la página

---

## 🔄 Cargar Excel manualmente
Si prefieres no usar GitHub Pages, abre `dashboard_ordenes.html` en cualquier navegador
y haz clic en **"Cargar Excel"** para subir el archivo desde tu computadora.

---

## 📊 Funcionalidades
- **Filtros en cascada**: Estatus · Back Order · Comprador · Proveedor · Fecha OC · Fecha Entrega · Folio OC
- **Búsqueda libre** por texto en proveedor u observaciones  
- **KPIs en tiempo real**: Total OC · Pendientes · Aplicados · Monto · Compradores · Proveedores  
- **Detalle desplegable** por OC: clave, descripción, marca, unidad, cantidad pedida, cantidad pendiente, back order, estatus compra, comprador  
- **Ordenamiento** por cualquier columna (clic en encabezado)  
- **Paginación** ajustable (25 / 50 / 100 / 250 filas)  

---

## 🗂 Columnas soportadas del Excel

**Hoja 1 — RP01-23 Ordenes de Compra**  
`Folio · Sucursal · Fecha · Proveedor · Plazo Pago · Subtotal · Importe IVA · Importe · Flete x Cobrar · Transportista · Fecha Entrega · Acepta Parciales · Observaciones · Estatus · Back Order · Referencia Lista · Comprador Capturo`

**Hoja 2 — Productos Incluidos en**  
`FOLIO OC · FECHA OC · PROVEEDOR · CLAVE · DESCRIPCION · MARCA · UNIDAD · CONTENIDO · PZASXUNI · CANTIDAD PED · CANTIDAD PEND · BACKORDER · FECHA ENTREGA · ESTATUS COMPRA · COMPRADOR`
