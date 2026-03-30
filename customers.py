import pandas as pd
from jinja2 import Environment, FileSystemLoader
import json

TEMPLATE_FILE = "templates/template_cust.html"


CUST_INPUT_FILE = r"H:\My Drive\3.PDC\Auto-cueillette Haren (Responses).xlsx"
CUST_OUTPUT_FILE = r"H:\My Drive\3.PDC\export\clients.html"

def load_data():
    df = pd.read_excel(CUST_INPUT_FILE)

    # Nettoyage colonnes
    df.columns = [c.strip() for c in df.columns]

    # Supprimer colonnes adresse
    df = df[[c for c in df.columns if "Adresse" not in c.lower()]]
    df = df[[c for c in df.columns if "règles" not in c.lower()]]
    df = df.rename(columns={
        "Comment avez-vous entendu parler du projet d'auto-cueillette de l'Asbl Zinnepot ?": "Source",
        "Comment souhaitez-vous être informé des nouvelles du champ (récoltes disponibles, changement d'horaires, autres informations importantes...) ?": "Comm",
        "Souhaiteriez-vous avoir accès à d'autres produits ?": "Produits",
        "En combien de fois voulez-vous payer l'abonnement?": "Echelonnement",
        "Quels produits vous intéressent le plus ?": "Préférences",
        "Quel formule choisissez-vous ? 2": "Prix",
        "À quelle fréquence pensez-vous venir cueillir ?": "Frequence"
        }
    )
    # ✅ Conversion timestamps / dates
    for col in df.columns:
        if "date" in col.lower() or "time" in col.lower():
            df[col] = pd.to_datetime(df[col], errors="coerce") \
                        .dt.strftime("%Y-%m-%d %H:%M") \
                        .fillna("")

    df["Prix brut adulte x1"] = 0.0
    df["Prix brut enfant tous"] = 0.0
    for index, row in df.iterrows():
        raw_price_adulte = 0.0
        if row["Prix"] == "Prix solidaire":
            raw_price_adulte = 220 / 1.06
        elif row["Prix"] == "Prix juste":
            raw_price_adulte = 250 / 1.06
        elif row["Prix"] == "Prix soutient":
            raw_price_adulte = 280 / 1.06

        raw_price_enfant = 13.0 * row["Total des âges des enfants"] / 1.06
        if row["Echelonnement"] == "En deux fois, en début et milieu de saison":
            raw_price_adulte /= 2
            raw_price_enfant /= 2

        df.iloc[index, df.columns.get_loc("Prix brut adulte x1")] = round(raw_price_adulte, 3)
        df.iloc[index, df.columns.get_loc("Prix brut enfant tous")] = round(raw_price_enfant, 3)

    # Remplacer NaN
    df = df.fillna("")

    return df


def split_fields(record):
    col1 = {}
    col2 = {}

    for k, v in record.items():
        key = k.lower()

        if k == "Nom et prénom":
            continue

        if any(x in key for x in ["date", "time"]):
            pass
        if any(x in key for x in ["nombre", "prix", "adresse", "total", "kot", "echelonnement"]):
            col1[k] = v
        else:
            col2[k] = v

    return col1, col2


def gen_customer_html():
    df = load_data()
    items = df.to_dict(orient="records")
    processed = []

    for item in items:
        col1, col2 = split_fields(item)
        processed.append({
            "name": item.get("Nom et prénom", ""),
            "col1": col1,
            "col2": col2
        })

    env = Environment(loader=FileSystemLoader("."))
    template = env.get_template(TEMPLATE_FILE)
    html = template.render(items=processed, json_data=json.dumps(processed, ensure_ascii=False))

    with open(CUST_OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(html)
    print("HTML customer generated :", CUST_OUTPUT_FILE)
