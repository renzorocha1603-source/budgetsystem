# allison_extractor.py - Allison the AI Extraction Agent
import json
import re
import requests

# ============================================================================
# MISTRAL CONFIGURATION
# ============================================================================
MISTRAL_API_KEY = "em5oqjSdA1Nus9iUpa1MNAJtQA4YfCtK"
MISTRAL_URL = "https://api.mistral.ai/v1/chat/completions"
MISTRAL_MODEL = "mistral-small-latest"

# ============================================================================
# THE EXTRACTION PROMPT
# ============================================================================

SYSTEM_PROMPT = """You are a financial data extraction specialist. Extract P&L data from structured table text.

THE TABLE HAS 9 COLUMNS IN THIS EXACT ORDER (separated by |):
Col 0: Account Name | Col 1: Mois Courant | Col 2: Budget | Col 3: Écart | Col 4: An Préc | Col 5: Cumulatif | Col 6: Cumul budget | Col 7: Écart cumul | Col 8: An Préc cumul

WE ONLY WANT COLUMN 1 (MOIS COURANT).

Each row is formatted like: Account Name | number | number | number | ...

The FIRST number after the account name IS Mois Courant. Take that number.

If Mois Courant is 0.00 or nan/empty, the value is 0.

ACCOUNTS TO EXTRACT (template_row: account_name):
12: Revenus Journaliers (or Revenus horaires)
13: Revenus mensuels
14: Revenus Lave-Auto
17: Divers
20: Gratuités - mensuels
29: Salaires Stationnement
32: Uniformes
35: Entretien réparation - Nettoyage
36: Entretien réparation - Général
37: Entretien réparation - Equipement
41: Fourn. de stationnement
49: Frais de bureau
50: Télécommunication
53: Frais de cartes de crédit
56: Réclamations
57: Assurances Cautionnement
58: Taxes et permis
63: Honoraires de gestion

VALIDATION (extract but put in "validation" section):
- TOTAL REVENUS
- Total des frais d'exploitation → key: "TOTAL_EXPENSES"
- BÉNÉFICE NET → key: "BENEFICE_NET"

CRITICAL: Take the FIRST number after the account name. That's Mois Courant. No exceptions.

Return ONLY valid JSON:
{"template_data": {"12": 71064.17, "13": 43585.46, "14": 206.12, "17": 0, "20": -1206.86, "29": 12886.70, "32": 174.00, "35": 2527.56, "36": 3117.96, "37": 1160.00, "41": 0, "49": 150.00, "50": 380.09, "53": 5414.74, "56": 230.00, "57": -2622.50, "58": 0, "63": 1.84}, "validation": {"TOTAL_REVENUS": 113648.89, "TOTAL_EXPENSES": 19176.99, "BENEFICE_NET": 94470.06}}"""

def call_allison(structured_text, debug_updates=None):
    """Send structured table text to Allison and get data back."""
    
    if debug_updates is not None:
        debug_updates.append("🤖 Asking Allison to extract from structured table...")
    
    messages = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": f"Extract P&L data from this structured table. Each row has columns separated by |. The FIRST number after the account name is Mois Courant. Return ONLY JSON:\n\n{structured_text[:8000]}"}
    ]
    
    try:
        resp = requests.post(
            MISTRAL_URL,
            headers={
                "Authorization": f"Bearer {MISTRAL_API_KEY}",
                "Content-Type": "application/json",
            },
            json={
                "model": MISTRAL_MODEL,
                "messages": messages,
                "temperature": 0.1,
            },
            timeout=45,
        )
        
        if resp.status_code == 200:
            response_text = resp.json()["choices"][0]["message"]["content"]
            if debug_updates is not None:
                debug_updates.append(f"📝 Allison: {response_text[:200]}...")
            
            result = parse_response(response_text)
            if result:
                if debug_updates is not None:
                    td = result.get("template_data", {})
                    vd = result.get("validation", {})
                    debug_updates.append(f"✅ Allison: {len(td)} template + {len(vd)} validation accounts")
                return result
            else:
                if debug_updates is not None:
                    debug_updates.append("⚠️ Couldn't parse Allison response")
                return None
        else:
            if debug_updates is not None:
                debug_updates.append(f"❌ Mistral API error: {resp.status_code}")
            return None
            
    except Exception as e:
        if debug_updates is not None:
            debug_updates.append(f"❌ Allison error: {e}")
        return None

def parse_response(text):
    """Extract JSON from Allison's response."""
    text = re.sub(r'```json\s*', '', text)
    text = re.sub(r'```\s*', '', text)
    
    json_match = re.search(r'\{.*\}', text, re.DOTALL)
    if json_match:
        try:
            return json.loads(json_match.group())
        except json.JSONDecodeError:
            pass
    
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass
    
    return None
