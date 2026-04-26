
import pandas as pd
from jinja2 import Environment, FileSystemLoader
import json
from shared import ROOT_FOLDER

TEMPLATE_FILE = "templates/template_cust.html"


CUST_INPUT_FILE = ROOT_FOLDER / "Auto-cueillette Haren (Responses).xlsx"
CUST_DATA_FILE = ROOT_FOLDER / "cueilleurs.xlsx"
CUST_OUTPUT_FILE = ROOT_FOLDER / "export" / "clients.html"


def load_data():
    print(f" ### Loading file {CUST_INPUT_FILE.name}")
    df = pd.read_excel(CUST_INPUT_FILE, skiprows=[1], index_col=11)
    df.index = df.index.str.strip()
    df.drop(df.tail(2).index, inplace=True)  # drop last 2 rows
    df.columns = [c.strip() for c in df.columns]

    df = df.rename(columns={
        "Comment avez-vous entendu parler du projet d'auto-cueillette de l'Asbl Zinnepot ?": "Source",
        "Comment souhaitez-vous être informé des nouvelles du champ (récoltes disponibles, changement d'horaires, autres informations importantes...) ?": "Comm",
        "Souhaiteriez-vous avoir accès à d'autres produits ?": "Produits",
        "En combien de fois voulez-vous payer l'abonnement?": "Abonnement",
        "Commentaires, allergies ou informations importantes à nous communiquer": "Commentaires",
        "Quels produits vous intéressent le plus ?": "Préférences",
        "Quel formule choisissez-vous ? 2": "Prix",
        "Infos supplémentaires, si étudiants en kot précisez les âges": "Infos",
        "À quelle fréquence pensez-vous venir cueillir ?": "Frequence",
        "Nombre d'adultes du foyer": "Nb adultes",
        "Jours préférés pour la cueillette": "Jours préférés",
        "Comment voulez-vous payer l'abonnement ?": "Mode paiement",
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

        raw_price_enfant = 13.0 * row["Ages des enfants"] / 1.06
        if row["Abonnement"] == "En deux fois, en début et milieu de saison":
            raw_price_adulte /= 2
            raw_price_enfant /= 2
            abonnement = "Mi-saison"
        else:
            abonnement = "Complète"

        df.loc[index, "Abonnement"] = abonnement
        df.loc[index, "Prix brut adulte x1"] = round(raw_price_adulte, 3)
        df.loc[index, "Prix brut enfant tous"] = round(raw_price_enfant, 3)

    # Remplacer NaN
    df = df.fillna("")

    print(f" ### Loading data file {CUST_DATA_FILE}")
    df_data = pd.read_excel(CUST_DATA_FILE, index_col=0)
    df_data.index = df_data.index.str.strip()
    set_index = set([x.strip() for x in set(df.index)])
    set_data_index = set([x.strip() for x in set(df_data.index)])
    if set_index != set_data_index:
        raise ValueError(
            f"Missing in reference: {set_data_index - set_index}\n"
            f"Missing in database: {set_index - set_data_index}"
        )

    print(" ### Somme des adultes: " + str(int(df["Nb adultes"].sum())))
    print(" ### Somme enfants: " + ", ".join(df["Ages des enfants"].dropna().astype(str)))

    return df.join(df_data)


def split_fields(record):
    col1 = {}
    col2 = {}

    for k, v in record.items():
        key = k.lower()

        if any(x in key for x in ["date", "time", "prénom", "règles", "objets"]):
            pass
        elif any(x in key for x in ["nombre", "nb", "prix", "adresse", "ages", "infos", "abonnement"]):
            col1[k] = v
        else:
            col2[k] = v

    return col1, col2


def gen_customer_html():
    df = load_data()
    items = df.T.to_dict()
    processed = []

    for item in items:
        col1, col2 = split_fields(items[item])
        notes = items[item]["Objets"]
        if isinstance(notes, str):
            notes = notes.split("\n")
        processed.append({
            "name": item,
            "col1": col1,
            "col2": col2,
            "notes": notes
        })

    env = Environment(loader=FileSystemLoader("."))
    template = env.get_template(TEMPLATE_FILE)
    html = template.render(items=processed, json_data=json.dumps(processed, ensure_ascii=False))

    with open(CUST_OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(html)
    print("HTML generated :", CUST_OUTPUT_FILE)
