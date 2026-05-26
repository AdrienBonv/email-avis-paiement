# webhook.py
import os
import requests
import base64
import json
import fitz
import easyocr
import ollama
import re
import warnings
warnings.filterwarnings("ignore")

from fastapi import FastAPI, Request, Query
from fastapi.responses import PlainTextResponse
from msal import ConfidentialClientApplication
from datetime import datetime, timedelta, timezone
from pdf2image import convert_from_path
import uvicorn
from dotenv import load_dotenv

import asyncio
from concurrent.futures import ThreadPoolExecutor

executor = ThreadPoolExecutor(max_workers=2)

load_dotenv()

app = FastAPI()

CLIENT_ID = os.getenv("GRAPH_CLIENT_ID")
CLIENT_SECRET = os.getenv("GRAPH_CLIENT_SECRET")
TENANT_ID = os.getenv("GRAPH_TENANT_ID")
BOITE = os.getenv("BOITE_EMAIL")
TON_SERVEUR = os.getenv("TON_SERVEUR")

# Charger EasyOCR une seule fois au démarrage
reader = easyocr.Reader(["fr", "en", "de"])

PROMPT = f"""Extrais les données de cet avis de paiement en JSON.

CHAMPS À EXTRAIRE :

"raison_sociale" : nom de l'entreprise cliente (le donneur d'ordre ou payeur).
Ignore les noms de banques (NATIXIS, CIC, SG, LCL...) et ignore [NOM DE VOTRE ENTREPRISE].

"numero_factures" : liste des numéros de facture.
Formats valides uniquement :
- [FORMAT N FACTURE] (2 lettres + 8 chiffres, ex: AA00001234)

"montant_factures" : liste des montants en euros correspondant à chaque facture.
Un montant contient au plus 2 décimales après un point ou une virgule (ex: 1200.00, 824,47).
Les dates (08.01.2026) et numéros de facture ne sont pas des montants.

"montant_ttc" : montant total de l'avis.

RÈGLES :
- Réponds UNIQUEMENT avec le JSON, rien d'autre.
- Si un champ est introuvable, mets null.
- Utilise le point comme séparateur décimal dans tous les montants.

Exemple :
{{"raison_sociale":"Dupont SARL","numero_factures":["2600001234","2600005678","99000001"],"montant_factures":["1200.00","850.00","30.00"],"montant_ttc":"2080.00"}}

Document :
"""


# ---- OCR et extraction ----
def detecter_type_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    texte = ""
    for page in doc:
        texte += page.get_text()
    if len(texte.strip()) > 50:
        return "natif", texte
    return "scan", None

def extraire_texte_scan(pdf_path):
    pages = convert_from_path(pdf_path, dpi=300)
    texte = ""
    for i, page in enumerate(pages):
        print(f"OCR page {i+1}/{len(pages)}...")
        page.save(f"/tmp/page_{i}.png")
        result = reader.readtext(f"/tmp/page_{i}.png", detail=0)
        texte += " ".join(result) + "\n"
    return texte

def parse_llm_json(raw: str):
    # Extraire le bloc JSON (entre les premières { } trouvées)
    match = re.search(r'\{[\s\S]*\}', raw)
    if not match:
        return None
    
    text = match.group()
    
    # Tenter le parse direct
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    # Nettoyage basique : commentaires, virgules en trop
    text = re.sub(r'//.*?$', '', text, flags=re.MULTILINE)  # commentaires
    text = re.sub(r',\s*([}\]])', r'\1', text)  # virgule avant } ou ]
    
    try:
        return json.loads(text)
    except json.JSONDecodeError as e:
        print(f"JSON irrécupérable: {e}")
        print(f"Réponse brute: {text[:500]}")
        return None


def analyser_page(texte_page, numero_page):
    response = ollama.chat(
        model="mistral",
        options={
            "num_ctx": 8192,
            "temperature": 0
        },
        messages=[{
            "role": "user",
            "content": f"{PROMPT}\n\nPage {numero_page} :\n{texte_page}",
        }]
    )
    return parse_llm_json(response["message"]["content"])

def agent_extraction(pdf_path):
    print(f"📄 Analyse de {pdf_path}...")
    type_pdf, texte_natif = detecter_type_pdf(pdf_path)
    print(f"🔍 Type détecté : {type_pdf}")

    if type_pdf == "natif":
        doc = fitz.open(pdf_path)
        textes_pages = [page.get_text() for page in doc]
    else:
        texte_scan = extraire_texte_scan(pdf_path)
        textes_pages = [texte_scan]
        
    # Classification multilingue par Mistral
    if not classifier_document(textes_pages[0]):
        print("⏭️ Pas un avis de paiement, ignoré")
        return None

    raison_sociale = ""
    toutes_factures = []
    tous_montants = []
    montant_ttc = ""

    for i, texte_page in enumerate(textes_pages):
        if len(texte_page.strip()) < 50:
            continue
        print(f"Mistral analyse page {i+1}/{len(textes_pages)}...")
        data = analyser_page(texte_page, i + 1)
        if not data:
            continue
        if not raison_sociale and data.get("raison_sociale"):
            raison_sociale = data["raison_sociale"]
        if data.get("numero_factures"):
            toutes_factures.extend(data["numero_factures"])
        if data.get("montant_factures"):
            tous_montants.extend(data["montant_factures"])
        if data.get("montant_ttc"):
            montant_ttc = data["montant_ttc"]

    return {
        "raison_sociale": raison_sociale,
        "numero_factures": toutes_factures,
        "montant_factures": tous_montants,
        "montant_ttc": montant_ttc
    }

# ---- Graph API ----
def get_token():
    msal_app = ConfidentialClientApplication(
        client_id=CLIENT_ID,
        client_credential=CLIENT_SECRET,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}"
    )
    result = msal_app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )
    return result["access_token"]

def classifier_document(texte):
    response = ollama.chat(
        model="mistral",
        options={"num_ctx": 4096},
        messages=[{
            "role": "user",
            "content": f"""Is this document a payment advice / remittance advice?
            can be written in French, English, German
Answer ONLY with "YES" or "NO". No explanation.

Document (first page):
{texte}"""
        }]
    )
    print(texte[:500])
    return "YES" in response["message"]["content"].upper()

def modifier_email(token, email_id, resultat):
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    
    email = requests.get(
        f"https://graph.microsoft.com/v1.0/users/{BOITE}/messages/{email_id}",
        headers=headers
    ).json()

    sujet = email.get("subject") or ""
    corps_original = email.get("body", {}).get("content", "")

    # Construire le nouveau sujet
    nouveau_sujet = f"{resultat.get('raison_sociale', '')} - {resultat.get('montant_ttc', '')} - {sujet}"
    
    # Numéros séparés par des sauts de ligne pour copie facile
    numeros_colonne = "<br>".join(resultat.get("numero_factures", []))
    montants_colonne = "<br>".join(resultat.get("montant_factures", []))

    nouveau_corps = f"""
    <html><body>
    <div style="background:#f0f7ff;padding:15px;border-left:4px solid #1F4E79;margin-bottom:20px;">
        <h3 style="color:#1F4E79;margin:0 0 10px 0">✅ Données extraites automatiquement</h3>
        <p><b>Client :</b> {resultat.get('raison_sociale', '')}</p>
        <p><b>Factures :</b></p>
        <table style="border-collapse:collapse;width:100%;max-width:500px;">
            <thead>
                <tr>
                    <th style="padding:6px 12px;border:1px solid #ddd;background:#1F4E79;color:white;text-align:left;">N° Factures</th>
                    <th style="padding:6px 12px;border:1px solid #ddd;background:#1F4E79;color:white;text-align:right;">Montants</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td style="padding:6px 12px;border:1px solid #ddd;vertical-align:top;font-family:monospace;">
                        {numeros_colonne}
                    </td>
                    <td style="padding:6px 12px;border:1px solid #ddd;vertical-align:top;text-align:right;font-family:monospace;">
                        {montants_colonne}
                    </td>
                </tr>
            </tbody>
            <tfoot>
                <tr>
                    <td style="padding:6px 12px;border:1px solid #ddd;background:#f5f5f5;font-weight:bold;">Total TTC</td>
                    <td style="padding:6px 12px;border:1px solid #ddd;background:#f5f5f5;font-weight:bold;text-align:right;">{resultat.get('montant_ttc', '')}€</td>
                </tr>
            </tfoot>
        </table>
    </div>
    {corps_original}
    </body></html>
    """

    response = requests.patch(
        f"https://graph.microsoft.com/v1.0/users/{BOITE}/messages/{email_id}",
        headers=headers,
        json={
            "subject": nouveau_sujet,
            "body": {
                "contentType": "HTML",
                "content": nouveau_corps
            }
        }
    )
    print(f"✏️ Email modifié : {response.status_code}")

def creer_subscription():
    token = get_token()
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    expiration = (datetime.now(timezone.utc) + timedelta(minutes=10069)).strftime("%Y-%m-%dT%H:%M:%SZ")
    body = {
        "changeType": "created",
        "notificationUrl": f"{TON_SERVEUR}/webhook",
        "resource": f"/users/{BOITE}/messages",
        "expirationDateTime": expiration,
        "clientState": "SECRET_VALIDATION"
    }
    response = requests.post(
        "https://graph.microsoft.com/v1.0/subscriptions",
        headers=headers,
        json=body
    )
    print("Subscription :", response.json())

# ---- Endpoints FastAPI ----
@app.get("/webhook")
async def valider_webhook(validationToken: str = Query(None)):
    if validationToken:
        print("✅ Webhook validé")
        return PlainTextResponse(validationToken)

emails_traites = set()

@app.post("/webhook")
async def recevoir_notification(request: Request, validationToken: str = Query(None)):
    if validationToken:
        return PlainTextResponse(validationToken)

    try:
        body = await request.json()
    except Exception:
        return PlainTextResponse("ok", status_code=200)

    for notification in body.get("value", []):
        if notification.get("clientState") != "SECRET_VALIDATION":
            continue
        email_id = notification["resourceData"]["id"]
        
        # Ignorer si déjà traité
        if email_id in emails_traites:
            print(f"⏭️ Email {email_id} déjà traité, ignoré")
            continue

        # Marquer comme traité immédiatement
        emails_traites.add(email_id)

        print(f"📧 Nouvel email détecté : {email_id}")
        loop = asyncio.get_event_loop()
        loop.run_in_executor(executor, traiter_email, email_id)

    return {"status": "accepted"}

def traiter_email(email_id):
    token = get_token()
    headers = {"Authorization": f"Bearer {token}"}

    email = requests.get(
        f"https://graph.microsoft.com/v1.0/users/{BOITE}/messages/{email_id}",
        headers=headers
    ).json()
    print(f"Sujet : {email.get('subject')}")

    attachments = requests.get(
        f"https://graph.microsoft.com/v1.0/users/{BOITE}/messages/{email_id}/attachments",
        headers=headers
    ).json()
    
    if not email.get("hasAttachments"):
        print("⏭️ Pas de pièce jointe, ignoré")
        return

    # Vérifier qu'il y a au moins un PDF
    pdfs = [a for a in attachments.get("value", []) if a.get("name", "").lower().endswith(".pdf")]
        
    if not pdfs:
        print("⏭️ Aucun PDF trouvé, ignoré")
        return


    # Traiter chaque PDF
    for attachment in pdfs:
        nom = attachment.get("name", "")
        print(f"📄 PDF trouvé : {nom}")

        contenu = base64.b64decode(attachment["contentBytes"])
        path = f"/tmp/{nom}"
        with open(path, "wb") as f:
            f.write(contenu)

        # Lancer agent
        resultat = agent_extraction(path)
        print(f"✅ Résultat : {resultat}")
        if resultat is None:
            print("⏭️ Document ignoré — pas un avis de paiement")
            return

        modifier_email(token, email_id, resultat)


    # Marquer comme lu
    requests.patch(
        f"https://graph.microsoft.com/v1.0/users/{BOITE}/messages/{email_id}",
        headers=headers,
        json={"isRead": True}
    )

if __name__ == "__main__":
    import threading
    import time

    def lancer_subscription():
        # Attendre que FastAPI soit prêt
        time.sleep(2)
        creer_subscription()

    # Lancer la subscription en arrière-plan
    thread = threading.Thread(target=lancer_subscription)
    thread.daemon = True
    thread.start()

    # Lancer FastAPI (bloquant)
    uvicorn.run(app, host="0.0.0.0", port=8000)