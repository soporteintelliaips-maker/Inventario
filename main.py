from flask import Flask, jsonify, send_file
import pandas as pd
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google.oauth2 import service_account
import io, os, tempfile

app = Flask(__name__)

FOLDER_ID = os.environ.get("FOLDER_ID")
CREDS_JSON = os.environ.get("GOOGLE_CREDENTIALS")

def get_drive_service():
    import json
    creds_info = json.loads(CREDS_JSON)
    creds = service_account.Credentials.from_service_account_info(
        creds_info,
        scopes=["https://www.googleapis.com/auth/drive"]
    )
    return build("drive", "v3", credentials=creds)

def download_file(service, file_id):
    request = service.files().get_media(fileId=file_id)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    buf.seek(0)
    return buf

# Detectar columna de cantidad
def detectar_columna_cantidad(df):
    posibles = ["cantidad", "cant", "qty", "stock", "existencia"]
    for col in df.columns:
        nombre = str(col).strip().lower()
        for p in posibles:
            if p in nombre:
                return col
    raise Exception(f"No se encontró columna de cantidad en: {list(df.columns)}")

# Detectar columna SKU
def detectar_columna_sku(df):
    posibles = ["sku", "codigo", "producto", "id"]
    for col in df.columns:
        nombre = str(col).strip().lower()
        for p in posibles:
            if p in nombre:
                return col
    raise Exception(f"No se encontró columna SKU en: {list(df.columns)}")

@app.route("/")
def home():
    return jsonify({"status": "ok"})

@app.route("/comparar", methods=["GET"])
def comparar():
    try:
        service = get_drive_service()

        results = service.files().list(
            q=f"'{FOLDER_ID}' in parents and mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' and trashed=false",
            fields="files(id, name, createdTime)",
            orderBy="createdTime desc"
        ).execute()

        files = results.get("files", [])

        if len(files) < 2:
            return jsonify({"error": "Se necesitan al menos 2 archivos Excel en la carpeta"}), 400

        file1 = files[0]
        file2 = files[1]

        buf1 = download_file(service, file1["id"])
        buf2 = download_file(service, file2["id"])

        liv = pd.read_excel(buf1)
        gym = pd.read_excel(buf2)

        # Detectar columnas
        col_sku_liv = detectar_columna_sku(liv)
        col_sku_gym = detectar_columna_sku(gym)

        col_qty_liv = detectar_columna_cantidad(liv)
        col_qty_gym = detectar_columna_cantidad(gym)

        # Normalizar SKU
        liv["SKU_norm"] = liv[col_sku_liv].astype(str).str.strip().str.upper()
        gym["SKU_norm"] = gym[col_sku_gym].astype(str).str.strip().str.upper()

        # Limpiar
        liv = liv[(liv["SKU_norm"] != "") & (liv["SKU_norm"] != "NAN")]
        gym = gym[(gym["SKU_norm"] != "") & (gym["SKU_norm"] != "NAN")]

        # Cantidades
        liv["_qty"] = pd.to_numeric(liv[col_qty_liv], errors="coerce").fillna(0)
        gym["_qty"] = pd.to_numeric(gym[col_qty_gym], errors="coerce").fillna(0)

        # Agrupar almacén
        gym_agg = gym.groupby("SKU_norm")["_qty"].sum()

        # 🔥 INTERSECCIÓN (solo los que están en ambos)
        skus_comunes = set(liv["SKU_norm"]) & set(gym_agg.index)

        rows = []

        for sku in sorted(skus_comunes):
            qty_liv = liv[liv["SKU_norm"] == sku]["_qty"].iloc[0]
            qty_gym = float(gym_agg.get(sku, 0))

            if qty_liv != qty_gym:
                rows.append({
                    "SKU": sku,
                    "Liverpool": qty_liv,
                    "Almacén": qty_gym,
                    "Diferencia": qty_gym - qty_liv,
                    "Acción": "Subir stock" if qty_gym > qty_liv else "Revisar"
                })

        df_result = pd.DataFrame(rows)

        # Exportar Excel
        tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        df_result.to_excel(tmp.name, index=False)

        return send_file(
            tmp.name,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name="Comparacion_Liverpool_vs_Almacen.xlsx"
        )

    except Exception as e:
        import traceback
        return jsonify({
            "error": str(e),
            "detalle": traceback.format_exc()
        }), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))