"""
convert_to_json.py
==================
Lee Estatus_de_BO.xlsx desde la carpeta data/ y genera data/data.json
para que el dashboard pueda consumirlo sin depender del Excel.

Ejecutado automáticamente por GitHub Actions al subir el Excel.
"""

import pandas as pd
import json
import os
from datetime import datetime

POSIBLES_RUTAS = [
    "Estatus_de_BO.xlsx",           # raíz del repo (más común)
    os.path.join("data", "Estatus_de_BO.xlsx"),  # carpeta data/
]
JSON_PATH = os.path.join("data", "data.json")


def encontrar_excel():
    for ruta in POSIBLES_RUTAS:
        if os.path.exists(ruta):
            print(f"✅ Excel encontrado en: {ruta}")
            return ruta
    raise FileNotFoundError(
        f"No se encontró el Excel. Rutas buscadas: {POSIBLES_RUTAS}\n"
        f"Sube Estatus_de_BO.xlsx a la raíz del repositorio."
    )


def safe_str(val):
    if val is None or (isinstance(val, float) and str(val) == "nan"):
        return ""
    return str(val).strip()


def safe_date(val):
    if pd.isna(val) if not isinstance(val, str) else False:
        return ""
    try:
        if isinstance(val, (pd.Timestamp, datetime)):
            return val.strftime("%Y-%m-%d")
        d = pd.to_datetime(val, errors="coerce")
        return d.strftime("%Y-%m-%d") if not pd.isna(d) else ""
    except Exception:
        return ""


def safe_num(val):
    try:
        f = float(val)
        return round(f, 4) if not (f != f) else 0   # NaN check
    except Exception:
        return 0


def main():
    EXCEL_PATH = encontrar_excel()
    print(f"📂 Leyendo: {EXCEL_PATH}")
    xl = pd.ExcelFile(EXCEL_PATH)
    print(f"   Hojas encontradas: {xl.sheet_names}")

    # ── Hoja 1: Órdenes de Compra ─────────────────────────────────────────
    df1 = pd.read_excel(EXCEL_PATH, sheet_name=0, dtype=str)
    df1 = df1.fillna("")

    orders = []
    for _, row in df1.iterrows():
        orders.append({
            "Folio":              safe_str(row.get("Folio", "")),
            "Sucursal":           safe_str(row.get("Sucursal", "")),
            "Fecha":              safe_str(row.get("Fecha", "")),
            "Proveedor":          safe_str(row.get("Proveedor", "")),
            "Plazo Pago":         safe_str(row.get("Plazo Pago", "")),
            "Subtotal":           safe_num(row.get("Subtotal", 0)),
            "Importe IVA":        safe_num(row.get("Importe IVA", 0)),
            "Importe":            safe_num(row.get("Importe", 0)),
            "Flete x Cobrar":     safe_str(row.get("Flete x Cobrar", "")),
            "Transportista":      safe_str(row.get("Transportista", "")),
            "Fecha Entrega":      safe_str(row.get("Fecha Entrega", "")),
            "Acepta Parciales":   safe_str(row.get("Acepta Parciales", "")),
            "Observaciones":      safe_str(row.get("Observaciones", "")),
            "Estatus":            safe_str(row.get("Estatus", "")),
            "Back Order":         safe_str(row.get("Back Order", "")),
            "Referencia Lista":   safe_str(row.get("Referencia Lista", "")),
            "Comprador Capturo":  safe_str(row.get("Comprador Capturo", "")),
        })

    print(f"   ✅ Órdenes procesadas: {len(orders)}")

    # ── Hoja 2: Productos ─────────────────────────────────────────────────
    df2 = pd.read_excel(EXCEL_PATH, sheet_name=1, dtype=str)
    df2 = df2.fillna("")

    products = {}   # keyed by folio OC
    for _, row in df2.iterrows():
        folio = safe_str(row.get("FOLIO OC", ""))
        if not folio:
            continue
        if folio not in products:
            products[folio] = []
        products[folio].append({
            "FOLIO OC":       folio,
            "FECHA OC":       safe_str(row.get("FECHA OC", "")),
            "PROVEEDOR":      safe_str(row.get("PROVEEDOR", "")),
            "CLAVE":          safe_str(row.get("CLAVE", "")),
            "DESCRIPCION":    safe_str(row.get("DESCRIPCION", "")),
            "MARCA":          safe_str(row.get("MARCA", "")),
            "UNIDAD":         safe_str(row.get("UNIDAD", "")),
            "CONTENIDO":      safe_str(row.get("CONTENIDO", "")),
            "PZASXUNI":       safe_num(row.get("PZASXUNI", 0)),
            "CANTIDAD PED":   safe_num(row.get("CANTIDAD PED", 0)),
            "CANTIDAD PEND":  safe_num(row.get("CANTIDAD PEND", 0)),
            "BACKORDER":      safe_num(row.get("BACKORDER", 0)),
            "FECHA ENTREGA":  safe_str(row.get("FECHA ENTREGA", "")),
            "ESTATUS COMPRA": safe_str(row.get("ESTATUS COMPRA", "")),
            "COMPRADOR":      safe_str(row.get("COMPRADOR", "")),
        })

    total_prods = sum(len(v) for v in products.values())
    print(f"   ✅ Claves procesadas:  {total_prods}  (en {len(products)} folios)")

    # ── Generar JSON ──────────────────────────────────────────────────────
    output = {
        "generated":     datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"),
        "totalOrders":   len(orders),
        "totalProducts": total_prods,
        "orders":        orders,
        "products":      products,
    }

    os.makedirs("data", exist_ok=True)
    with open(JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, separators=(",", ":"))

    size_kb = os.path.getsize(JSON_PATH) / 1024
    print(f"\n✅ Generado: {JSON_PATH}  ({size_kb:.1f} KB)")
    print(f"   Generado: {output['generated']}")


if __name__ == "__main__":
    main()
