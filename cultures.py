

import pandas as pd
import json
from pathlib import Path


INPUT_FILE = r"H:\My Drive\3.PDC\PDC.xlsx"
SHEET_NAME = "Crop"

TEMPLATE_FILE = "template.html"
OUTPUT_FILE = r"H:\My Drive\3.PDC\export\cultures.html"


class Culture:

    def __init__(self, **entries):
        for k, v in entries.items():
            setattr(self, k, v)

    def to_dict(self):
        return self.__dict__


def load_cultures():
    df = pd.read_excel(
        INPUT_FILE,
        sheet_name=SHEET_NAME,
        header=3
    )
    df = df.dropna(how="all")
    df = df.sort_values(by="Culture", key=lambda col: col.astype(str).str.lower())
    cultures = []

    for _, row in df.iterrows():
        data = {k: ("" if pd.isna(v) else str(v)) for k, v in row.to_dict().items()}
        crop_value = data.get("Culture", "")
        if crop_value in ["", "-"]:
            continue
        print(crop_value)

        cultures.append(Culture(**data))

    return cultures


def generate_html(cultures):
    template = Path(TEMPLATE_FILE).read_text(encoding="utf8")
    data = [c.to_dict() for c in cultures]
    json_data = json.dumps(data, ensure_ascii=False)
    html = template.replace("__DATA__", json_data)
    Path(OUTPUT_FILE).write_text(html, encoding="utf8")


def main():
    cultures = load_cultures()
    generate_html(cultures)
    print(f"{len(cultures)} crops exported → {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
