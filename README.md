# Email Avis de Paiement — Agent comptable IA

Agent automatique qui surveille une boîte mail Outlook, détecte les avis de paiement en pièce jointe (PDF), en extrait les données clés via IA, et enrichit l'email directement dans Outlook.

## Ce que ça fait

1. **Écoute** les nouveaux emails via un webhook Microsoft Graph (API Outlook 365)
2. **Détecte** si le PDF joint est un avis de paiement (Mistral classe le document)
3. **Extrait** automatiquement :
   - Raison sociale du client
   - Numéros de factures
   - Montants par facture
   - Montant total TTC
4. **Modifie l'email** dans Outlook : nouveau sujet `[Client] - [Montant] - [Sujet original]` + tableau récapitulatif inséré en haut du corps

Fonctionne sur les PDFs natifs (texte extractible) **et** les scans (OCR via EasyOCR).

---

## Stack technique

| Composant | Rôle |
|-----------|------|
| **FastAPI** | Serveur webhook qui reçoit les notifications Outlook |
| **Microsoft Graph API** | Lecture des emails et pièces jointes, modification des emails |
| **Ollama + Mistral** | Classification du document + extraction JSON des données |
| **PyMuPDF** | Extraction de texte des PDFs natifs |
| **EasyOCR** | OCR pour les PDFs scannés (fr, en, de) |
| **Docker Compose** | Orchestration des deux services (agent + Ollama) |

---

## Prérequis

- Docker et Docker Compose installés
- Une application Azure AD avec les permissions Microsoft Graph :
  - `Mail.Read`
  - `Mail.ReadWrite`
- Un serveur accessible publiquement (pour que Microsoft puisse envoyer les webhooks)

---

## Configuration

Copier `.env-example` en `.env` et remplir les valeurs :

```env
GRAPH_CLIENT_ID=<ID de l'application Azure AD>
GRAPH_TENANT_ID=<ID du tenant Azure>
GRAPH_CLIENT_SECRET=<Secret de l'application>
BOITE_EMAIL=<adresse@domaine.com>
TON_SERVEUR=https://votre-serveur.com  # URL publique du serveur
```

### Créer l'application Azure AD

1. Azure Portal → **Azure Active Directory** → **Inscriptions d'applications** → Nouvelle inscription
2. Ajouter les permissions API : `Mail.Read`, `Mail.ReadWrite` (permissions d'application, pas déléguées)
3. Créer un secret client → copier la valeur dans `GRAPH_CLIENT_SECRET`

---

## Lancement

```bash
# 1. Configurer l'environnement
cp .env-example .env
# Remplir .env avec vos valeurs

# 2. Démarrer les services
docker compose up -d

# 3. Télécharger le modèle Mistral (première fois seulement)
docker compose exec ollama ollama pull mistral
```

L'agent démarre sur le port `8000` et crée automatiquement la subscription Microsoft Graph au démarrage.

> **Note :** La subscription Outlook expire après ~7 jours. Redémarrer le service la renouvelle automatiquement.

---

## Architecture

```
Email reçu (Outlook 365)
        │
        ▼
  Webhook /POST        ← Microsoft Graph notifie le serveur
        │
        ▼
  Pièce jointe PDF ?
        │ oui
        ▼
  Mistral : avis de paiement ?
        │ oui
        ▼
  PDF natif ──── PyMuPDF ────┐
  PDF scanné ─── EasyOCR ───┤
                              ▼
                     Mistral : extraction JSON
                    (client, factures, montants)
                              │
                              ▼
                     Modification de l'email
                    (sujet + tableau HTML)
```

---

## Formats de factures reconnus

Le prompt est configuré pour extraire les numéros au format `XX00000000` (2 lettres + 8 chiffres). Adapter la ligne suivante dans `agent-compta.py` si votre format est différent :

```python
# ligne ~48
- [FORMAT N FACTURE] (2 lettres + 8 chiffres, ex: AA00001234)
```

---

## Licence

Apache 2.0