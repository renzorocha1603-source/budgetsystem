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

SYSTEM_PROMPT = """You are a financial data extraction specialist. Extract P&L data from PDF text.

THE TABLE HAS 9 COLUMNS:
Col 0: Account Name | Col 1: Mois Courant (WHAT WE NEED) | Col 2: Budget | Col 3: Écart | Col 4: An Préc | Col 5: Cumulatif | Col 6: Cumul budget | Col 7: Écart cumul | Col 8: An Préc cumul

WE ONLY WANT COLUMN 1 (MOIS COURANT) - the FIRST number after each account name.

CANADIAN FRENCH FORMAT: "43 585,46 $" = 43585.46 | "(1 206,86)" = -1206.86 | Empty cell = 0.00

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

VALIDATION (extract but don't write to template):
- TOTAL REVENUS
- Total des frais d'exploitation (TOTAL_EXPENSES)
- BÉNÉFICE NET (BENEFICE_NET)

CRITICAL RULES:
1. Take ONLY the FIRST number after the account name (Mois Courant)
2. If no number within ~30 characters, value is 0
3. DO NOT take numbers from Budget/YTD columns
4. Keep negative signs for Gratuités, Assurances, etc.
5. Return ONLY valid JSON, nothing else

Return EXACTLY this JSON structure:
{"template_data": {"12": 71064.17, "13": 43585.46, "14": 206.12, "17": 0, "20": -1206.86, "29": 12886.70, "32": 174.00, "35": 2527.56, "36": 3117.96, "37": 1160.00, "41": 0, "49": 150.00, "50": 380.09, "53": 5414.74, "56": 230.00, "57": -2622.50, "58": 0, "63": 1.84}, "validation": {"TOTAL_REVENUS": 113648.89, "TOTAL_EXPENSES": 19176.99, "BENEFICE_NET": 94470.06}}"""

def call_allison(pdf_text, debug_updates=None):
    """Send PDF text to Allison and get structured data back."""
    
    if debug_updates is not None:
        debug_updates.append("🤖 Asking Allison to extract data...")
    
    messages = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": f"Extract P&L data from this PDF text. Return ONLY JSON:\n\n{pdf_text[:10000]}"}
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
                debug_updates.append(f"📝 Allison response: {response_text[:200]}...")
            
            # Parse JSON from response
            result = parse_response(response_text)
            if result:
                if debug_updates is not None:
                    td = result.get("template_data", {})
                    vd = result.get("validation", {})
                    debug_updates.append(f"✅ Allison extracted {len(td)} template + {len(vd)} validation accounts")
                return result
            else:
                if debug_updates is not None:
                    debug_updates.append("⚠️ Allison response couldn't be parsed as JSON")
                return None
        else:
            if debug_updates is not None:
                debug_updates.append(f"❌ Mistral API error: {resp.status_code}")
            return None
            
    except requests.exceptions.Timeout:
        if debug_updates is not None:
            debug_updates.append("❌ Allison timed out")
        return None
    except Exception as e:
        if debug_updates is not None:
            debug_updates.append(f"❌ Allison error: {e}")
        return None

def parse_response(text):
    """Extract JSON from Allison's response."""
    # Try to find JSON object in the text
    json_match = re.search(r'\{.*\}', text, re.DOTALL)
    if json_match:
        try:
            return json.loads(json_match.group())
        except json.JSONDecodeError:
            pass
    
    # Try parsing the whole text
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass
    
    return None
