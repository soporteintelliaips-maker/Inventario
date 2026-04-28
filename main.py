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

        # Tomar los 2 más recientes
        file1 = files[0]
        file2 = files[1]

        buf1 = download_file(service, file1["id"])
        buf2 = download_file(service, file2["id"])

        liv = pd.read_excel(buf1)
        gym = pd.read_excel(buf2)

        # Normalizar SKU
        liv["SKU_norm"] = liv.iloc[:, 0].astype(str).str.strip().str.upper()
        gym["SKU_norm"] = gym.iloc[:, 0].astype(str).str.strip().str.upper()

        # Limpiar SKUs inválidos
        liv = liv[(liv["SKU_norm"].notna()) & (liv["SKU_norm"] != "NAN") & (liv["SKU_norm"] != "")]
        gym = gym[(gym["SKU_norm"].notna()) & (gym["SKU_norm"] != "NAN") & (gym["SKU_norm"] != "")]

        # Cantidades
        liv["_qty"] = pd.to_numeric(liv.iloc[:, 1], errors="coerce").fillna(0)
        gym["_qty"] = pd.to_numeric(gym.iloc[:, 1], errors="coerce").fillna(0)

        # Índices
        liv_idx = liv.set_index("SKU_norm")
        gym_agg = gym.groupby("SKU_norm")["_qty"].sum()

        # Sets
        liv_skus = set(liv["SKU_norm"])
        gym_skus = set(gym["SKU_norm"])

        en_ambos = liv_skus & gym_skus
        solo_liv = liv_skus - gym_skus
        solo_gym = gym_skus - liv_skus

        rows = []

        # Comparar
        for sku in sorted(en_ambos):
            qty_liv = liv_idx.loc[sku, "_qty"]
            q1 = float(qty_liv.iloc[0] if hasattr(qty_liv, "iloc") else qty_liv)
            q2 = float(gym_agg.get(sku, 0))

            if q1 != q2:
                rows.append({
                    "SKU": sku,
                    "Cantidad Liverpool": q1,
                    "Cantidad Almacén": q2,
                    "Diferencia": q2 - q1,
                    "Tipo": "Cantidad diferente"
                })

        # Solo Liverpool
        for sku in sorted(solo_liv):
            qty = liv_idx.loc[sku, "_qty"]
            q1 = float(qty.iloc[0] if hasattr(qty, "iloc") else qty)

            rows.append({
                "SKU": sku,
                "Cantidad Liverpool": q1,
                "Cantidad Almacén": 0,
                "Diferencia": None,
                "Tipo": "Solo en Liverpool"
            })

        # Solo Almacén
        for sku in sorted(solo_gym):
            q2 = float(gym_agg.get(sku, 0))

            rows.append({
                "SKU": sku,
                "Cantidad Liverpool": 0,
                "Cantidad Almacén": q2,
                "Diferencia": None,
                "Tipo": "Solo en Almacén"
            })

        df_result = pd.DataFrame(rows)

        # Crear archivo temporal
        tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        df_result.to_excel(tmp.name, index=False)

        return send_file(
            tmp.name,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name="Diferencias_Inventario.xlsx"
        )

    except Exception as e:
        import traceback
        return jsonify({
            "error": str(e),
            "detalle": traceback.format_exc()
        }), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))