"""
convert_to_json.py
==================
Convierte Estatus_de_BO.xlsx → data/data.json para el dashboard.
Busca hojas por NOMBRE para ser robusto a cambios de orden.
"""

import pandas as pd
import json
import os
import re
from datetime import datetime

JSON_PATH = os.path.join("data", "data.json")


# ── Validar Excel ─────────────────────────────────────────────────────────
def es_excel_valido(ruta):
    try:
        pd.ExcelFile(ruta, engine='openpyxl')
        return True
    except Exception as e:
        print(f"   ⚠ Archivo inválido ({ruta}): {e}")
        return False


# ── Encontrar Excel ───────────────────────────────────────────────────────
def encontrar_excel():
    print("\n📁 Archivos en el directorio actual:")
    for f in sorted(os.listdir(".")):
        print(f"   {f}")

    carpeta_data = "data"
    if os.path.isdir(carpeta_data):
        print(f"\n📁 Archivos en {carpeta_data}/:")
        for f in sorted(os.listdir(carpeta_data)):
            print(f"   {f}")

    nombres_pref = [
        "Estatus de BO.xlsx", "Estatus_de_BO.xlsx",
        "Estatus BO.xlsx",    "Estatus_BO.xlsx",
    ]
    carpetas = [carpeta_data, "."]
    for carpeta in carpetas:
        for nombre in nombres_pref:
            ruta = os.path.join(carpeta, nombre)
            if os.path.exists(ruta) and es_excel_valido(ruta):
                print(f"\n✅ Excel encontrado: {ruta}")
                return ruta

    # Búsqueda amplia: cualquier .xlsx válido
    for carpeta in carpetas:
        base = carpeta if os.path.isdir(carpeta) else "."
        for f in sorted(os.listdir(base)):
            if f.lower().endswith(".xlsx"):
                ruta = os.path.join(base, f)
                if es_excel_valido(ruta):
                    print(f"\n✅ Excel encontrado (búsqueda amplia): {ruta}")
                    return ruta

    raise FileNotFoundError("No se encontró ningún archivo Excel válido.")


# ── Encontrar hojas por contenido ─────────────────────────────────────────
def encontrar_hoja_ordenes(xl):
    """Hoja principal de órdenes: tiene Folio, Estatus, Back Order."""
    for name in xl.sheet_names:
        try:
            df = pd.read_excel(xl, sheet_name=name, nrows=1, engine='openpyxl')
            cols = [str(c).strip() for c in df.columns]
            if 'Folio' in cols and 'Estatus' in cols and 'Back Order' in cols:
                print(f"   ✅ Hoja Órdenes: \"{name}\"")
                return name
        except Exception:
            continue
    raise ValueError("No se encontró la hoja de Órdenes de Compra.")


def encontrar_hoja_productos(xl):
    """Hoja de productos: tiene FOLIO OC y CLAVE."""
    for name in xl.sheet_names:
        try:
            df = pd.read_excel(xl, sheet_name=name, nrows=1, engine='openpyxl')
            cols = [str(c).strip() for c in df.columns]
            if 'FOLIO OC' in cols and 'CLAVE' in cols:
                print(f"   ✅ Hoja Productos: \"{name}\"")
                return name
        except Exception:
            continue
    raise ValueError("No se encontró la hoja de Productos.")


def encontrar_hoja_estado_oc(xl):
    """Hoja con Estado OC: tiene columna 'Estado OC'."""
    for name in xl.sheet_names:
        try:
            df = pd.read_excel(xl, sheet_name=name, nrows=1, engine='openpyxl')
            cols = [str(c).strip() for c in df.columns]
            if 'Estado OC' in cols and 'Folio' in cols:
                print(f"   ✅ Hoja Estado OC: \"{name}\"")
                return name
        except Exception:
            continue
    print("   ℹ No se encontró hoja con 'Estado OC'")
    return None


# ── Helpers ───────────────────────────────────────────────────────────────
def safe_str(val):
    if val is None or (isinstance(val, float) and str(val) == 'nan'):
        return ''
    return str(val).strip()


def safe_num(val):
    try:
        f = float(val)
        return round(f, 4) if f == f else 0
    except Exception:
        return 0


def clean_date(val):
    """Convierte cualquier formato de fecha a YYYY-MM-DD."""
    if not val or val == '':
        return ''
    s = str(val).strip()
    # Quitar timestamp si existe: "2026-01-02 09:17:57.620000" → "2026-01-02"
    s = re.split(r'[\sT]', s)[0]
    # Validar formato mínimo
    if len(s) >= 10 and s[4] == '-':
        return s[:10]   # YYYY-MM-DD
    return s


# ── Main ──────────────────────────────────────────────────────────────────
def main():
    EXCEL_PATH = encontrar_excel()
    print(f"\n📂 Procesando: {EXCEL_PATH}")

    xl = pd.ExcelFile(EXCEL_PATH, engine='openpyxl')
    print(f"   Hojas: {xl.sheet_names}")

    hoja_oc      = encontrar_hoja_ordenes(xl)
    hoja_prod    = encontrar_hoja_productos(xl)
    hoja_estado  = encontrar_hoja_estado_oc(xl)

    # ── Órdenes ───────────────────────────────────────────────────────────
    df_oc = pd.read_excel(EXCEL_PATH, sheet_name=hoja_oc, dtype=str, engine='openpyxl').fillna('')

    orders = []
    for _, row in df_oc.iterrows():
        folio = safe_str(row.get('Folio', ''))
        if not folio:
            continue
        orders.append({
            'Folio':             folio,
            'Sucursal':          safe_str(row.get('Sucursal', '')),
            'Fecha':             clean_date(row.get('Fecha', '')),
            'Proveedor':         safe_str(row.get('Proveedor', '')),
            'Plazo Pago':        safe_str(row.get('Plazo Pago', '')),
            'Subtotal':          safe_num(row.get('Subtotal', 0)),
            'Importe IVA':       safe_num(row.get('Importe IVA', 0)),
            'Importe':           safe_num(row.get('Importe', 0)),
            'Flete x Cobrar':    safe_str(row.get('Flete x Cobrar', '')),
            'Transportista':     safe_str(row.get('Transportista', '')),
            'Fecha Entrega':     clean_date(row.get('Fecha Entrega', '')),
            'Acepta Parciales':  safe_str(row.get('Acepta Parciales', '')),
            'Observaciones':     safe_str(row.get('Observaciones', '')),
            'Estatus':           safe_str(row.get('Estatus', '')).upper(),
            'Back Order':        safe_str(row.get('Back Order', '')).upper(),
            'Referencia Lista':  safe_str(row.get('Referencia Lista', '')),
            'Comprador Capturo': safe_str(row.get('Comprador Capturo', '')),
            'Estado OC':         'RECIBIDO',   # default, se sobreescribe abajo
        })
    print(f"   ✅ Órdenes: {len(orders)}")

    # ── Estado OC ─────────────────────────────────────────────────────────
    if hoja_estado:
        df_est = pd.read_excel(EXCEL_PATH, sheet_name=hoja_estado, dtype=str, engine='openpyxl').fillna('')
        estado_map = {}
        for _, r in df_est.iterrows():
            f = safe_str(r.get('Folio', ''))
            e = safe_str(r.get('Estado OC', '')).upper()
            if f and e:
                estado_map[f] = e
        for o in orders:
            if o['Folio'] in estado_map:
                o['Estado OC'] = estado_map[o['Folio']]
        print(f"   ✅ Estado OC mapeados: {len(estado_map)}")

    # ── Productos ─────────────────────────────────────────────────────────
    df_prod = pd.read_excel(EXCEL_PATH, sheet_name=hoja_prod, dtype=str, engine='openpyxl').fillna('')

    products = {}
    for _, row in df_prod.iterrows():
        folio = safe_str(row.get('FOLIO OC', ''))
        if not folio:
            continue
        if folio not in products:
            products[folio] = []
        products[folio].append({
            'FOLIO OC':       folio,
            'FECHA OC':       clean_date(row.get('FECHA OC', '')),
            'PROVEEDOR':      safe_str(row.get('PROVEEDOR', '')),
            'CLAVE':          safe_str(row.get('CLAVE', '')),
            'DESCRIPCION':    safe_str(row.get('DESCRIPCION', '')),
            'MARCA':          safe_str(row.get('MARCA', '')),
            'UNIDAD':         safe_str(row.get('UNIDAD', '')),
            'CONTENIDO':      safe_str(row.get('CONTENIDO', '')),
            'PZASXUNI':       safe_num(row.get('PZASXUNI', 0)),
            'CANTIDAD PED':   safe_num(row.get('CANTIDAD PED', 0)),
            'CANTIDAD PEND':  safe_num(row.get('CANTIDAD PEND', 0)),
            'BACKORDER':      safe_num(row.get('BACKORDER', 0)),
            'FECHA ENTREGA':  clean_date(row.get('FECHA ENTREGA', '')),
            'ESTATUS COMPRA': safe_str(row.get('ESTATUS COMPRA', '')),
            'COMPRADOR':      safe_str(row.get('COMPRADOR', '')),
        })

    total_prods = sum(len(v) for v in products.values())
    print(f"   ✅ Claves: {total_prods} en {len(products)} folios")

    # ── Generar JSON ──────────────────────────────────────────────────────
    output = {
        'generated':     datetime.utcnow().strftime('%Y-%m-%dT%H:%M:%SZ'),
        'totalOrders':   len(orders),
        'totalProducts': total_prods,
        'orders':        orders,
        'products':      products,
    }

    os.makedirs('data', exist_ok=True)
    with open(JSON_PATH, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, separators=(',', ':'))

    size_kb = os.path.getsize(JSON_PATH) / 1024
    print(f"\n✅ Generado: {JSON_PATH}  ({size_kb:.1f} KB)")
    print(f"   Timestamp: {output['generated']}")
    print(f"   Órdenes: {len(orders)} | Claves: {total_prods}")


if __name__ == '__main__':
    main()
