"""
Vercel serverless function — retorna a escala do dia.
Env vars: SP_EMAIL, SP_PASS, SP_SHARE_URL
"""

import os
import json
import datetime
import base64
import time
import tempfile
from http.server import BaseHTTPRequestHandler
import msal
import requests
import pandas as pd

TENANT_ID = "1fd54b8d-ac86-436d-8e57-0b88a2805609"
CLIENT_ID = "d3590ed6-52b3-4102-aeff-aad2292ab01c"
VALID_SHIFTS = {"SN", "P"}


def get_token():
    email = os.environ.get("SP_EMAIL")
    password = os.environ.get("SP_PASS")
    if not email or not password:
        raise ValueError("SP_EMAIL e SP_PASS nao definidos")

    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app = msal.PublicClientApplication(CLIENT_ID, authority=authority)
    result = app.acquire_token_by_username_password(
        email, password, scopes=["https://graph.microsoft.com/.default"]
    )
    if "access_token" not in result:
        raise ValueError(result.get("error_description", "Falha na autenticacao"))
    return result["access_token"]


def download_xlsx(token):
    share_url = os.environ.get("SP_SHARE_URL", "")
    if not share_url:
        raise ValueError("SP_SHARE_URL nao definida")

    encoded = base64.urlsafe_b64encode(share_url.encode()).decode().rstrip("=")
    share_id = f"u!{encoded}"

    headers = {"Authorization": f"Bearer {token}"}

    for attempt in range(3):
        r = requests.get(
            f"https://graph.microsoft.com/v1.0/shares/{share_id}/driveItem",
            headers=headers,
            timeout=15,
        )
        if r.status_code == 429:
            retry_after = int(r.headers.get("Retry-After", 5))
            time.sleep(retry_after)
            continue
        break

    if r.status_code != 200:
        raise ValueError(f"Erro ao acessar arquivo: {r.status_code} - {r.text[:200]}")

    item = r.json()
    dl_url = item.get("@microsoft.graph.downloadUrl")
    if dl_url:
        resp = requests.get(dl_url, timeout=30)
    else:
        resp = requests.get(
            f"https://graph.microsoft.com/v1.0/shares/{share_id}/driveItem/content",
            headers=headers,
            timeout=30,
        )

    if resp.status_code != 200:
        raise ValueError(f"Erro no download: {resp.status_code}")

    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    tmp.write(resp.content)
    tmp.close()
    return tmp.name


def load_adm():
    adm_path = os.path.join(os.path.dirname(__file__), "colaboradores.xlsx")
    if not os.path.exists(adm_path):
        return []
    df = pd.read_excel(adm_path, header=0)
    df.columns = [c.strip() for c in df.columns]
    sit_col = [c for c in df.columns if "SITUA" in c][0]
    result = []
    for _, row in df.iterrows():
        cargo = str(row["CARGO"]).strip()
        situacao = str(row[sit_col]).strip()
        if situacao != "ATIVO":
            continue
        if "ENFERMAG" in cargo.upper() or "ENFERMEIRO" in cargo.upper():
            continue
        nome = str(row["NOME"]).strip()
        result.append({
            "nome": nome,
            "cargo": cargo,
            "jornada": int(row["JORNADA"]),
        })
    result.sort(key=lambda x: x["nome"])
    return result


def find_table(df):
    hr = None
    for i, row in df.iterrows():
        for val in row:
            if isinstance(val, str) and "MATR" in val.upper():
                hr = i
                break
        if hr is not None:
            break
    if hr is None:
        return None, None
    er = hr + 1
    ec = 0
    for i in range(hr + 1, len(df)):
        nm = df.iloc[i, 2] if pd.notna(df.iloc[i, 2]) else ""
        if isinstance(nm, str) and "ESCALA" in nm.upper():
            break
        if isinstance(nm, str) and len(nm.strip()) > 2:
            er = i + 1
            ec = 0
        else:
            ec += 1
            if ec > 5:
                break
    return hr, er


def extract_schedule(df, day):
    hr, er = find_table(df)
    if hr is None:
        return []
    dc = 6 + day
    result = []
    for i in range(hr + 1, er):
        nm = df.iloc[i, 2] if pd.notna(df.iloc[i, 2]) else ""
        if not isinstance(nm, str) or len(nm.strip()) < 3:
            continue
        nm = nm.strip().replace("\n", "")
        if "COREN:" in nm:
            nm = nm.split("COREN:")[0].strip()
        if dc < len(df.columns):
            cell = df.iloc[i, dc]
            if pd.notna(cell) and isinstance(cell, str):
                sh = cell.strip().upper()
                if sh in VALID_SHIFTS:
                    lot = str(df.iloc[i, 4]).strip() if pd.notna(df.iloc[i, 4]) else ""
                    funcao = str(df.iloc[i, 3]).strip() if pd.notna(df.iloc[i, 3]) else ""
                    result.append({
                        "nome": nm,
                        "plantao": sh,
                        "lotacao": lot,
                        "funcao": funcao,
                    })
    return result


class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        try:
            today = datetime.date.today()
            day = today.day

            token = get_token()
            fp = download_xlsx(token)

            data = {
                "data": today.isoformat(),
                "dia": day,
                "data_formatada": today.strftime("%d/%m/%Y"),
                "setores": {},
            }

            for sheet in ["TECNICOS", "ENFERMEIRAS", "CME"]:
                try:
                    df = pd.read_excel(fp, sheet_name=sheet, header=None)
                    data["setores"][sheet] = extract_schedule(df, day)
                except Exception:
                    data["setores"][sheet] = []

            try:
                os.unlink(fp)
            except Exception:
                pass

            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(json.dumps(data, ensure_ascii=False).encode("utf-8"))

        except Exception as e:
            self.send_response(500)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps({"error": str(e)}).encode("utf-8"))
