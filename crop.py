import operator
import re
from typing import List

import pandas as pd
from itk import parse_itk

MC_INPUT_FILE = r"H:\My Drive\3.PDC\Crop_planning_4.0_ Eng_MC1.xlsx"
MC_SHEET_NAME = "Crop Chart"


class CropTask:
    def __init__(self, days, name):
        self.days = days
        self.name = name
        self.weeks = int(float(self.days) / 7)
        self.week_abs = 0

    def update(self, grow_start):
        self.week_abs = grow_start + self.weeks

    def __str__(self):
        return f'{self.name} {self.week_abs}'


def check_dtm(dtm):
    if dtm is None:
        return None
    return None


class Crop:

    def __init__(self, **entries):
        self._tasks: List[CropTask] = []

        for k, v in entries.items():
            try:
                setattr(self, k, float(v))
            except ValueError:
                setattr(self, k, v)

    def to_dict(self):
        return self.__dict__

    def get_int(self, prop):
        attr_val = getattr(self, prop)
        if isinstance(attr_val, str) and "-" in attr_val:
            return attr_val
        return int(float(attr_val)) if attr_val else 0

    def to_print(self):
        useless_attrs = ["Terreau semis", "category", "Technique semis pep", "Densité graines g/30m",
                         "Jours en pep", "Jours au champ", "DTM", "Variété", "Nb jours avant levée", "Lignes / planche",
                         "Espacement sur la ligne (cm)", 'Calibration', "Fenêtre récolte (jours)",
                         'Type de plateau', 'Plateau transplant', "# graines / cellule", "# cellule / planche 12m",
                         "% marge sécurité pour transplants", "# plateau / planche 12m",
                         "Température germination (C)", "Terreau rempotage", "Row marking"
                         ]

        items_print = {
            "Jours en pep / au champ / DTM / récolte":
                f'{self.get_int("Jours en pep")} / {self.get_int("Jours au champ")} / '
                f'{self.get_int("DTM")} / {self.get_int("Fenêtre récolte (jours)")}',
            "Lignes/planche": f'{self.get_int("Lignes / planche")}r'
                                            f'{self.get_int("Espacement sur la ligne (cm)")}',

        }
        if getattr(self, "Technique") == "TR":
            items_print[f"#graines/cellule, marge, #cellule/12m, #plateau {self.get_int('Type de plateau')}"] = (
                f'{self.get_int("# graines / cellule")} / '
                f'{getattr(self, "% marge sécurité pour transplants")} / '
                f'{self.get_int("# cellule / planche 12m")} / '
                f'{getattr(self, "# plateau / planche 12m")}'
            )
        items_print.update(**self.to_dict().copy())

        for item in useless_attrs:
            if item in items_print:
                items_print.pop(item)
        return items_print

    @property
    def crop(self):
        return getattr(self, "Culture")

    def __str__(self):
        return self.crop

    def prepare_print(self):
        attrs = list(vars(self).keys())
        new_attrs = {}
        to_delete = []

        for key in attrs:
            m = re.match(r"Tâche (\d+)", key)
            if m:
                i = m.group(1)

                task = getattr(self, key)
                days_key = f"# jours {i}"

                if hasattr(self, days_key):
                    days = getattr(self, days_key)
                    if days != "" and task not in ["NS", "TR", "DS", "Crop out", "Harvest starts"]:
                        days_int = int(float(days))
                        new_key = f"Tâche J={days_int}"
                        new_attrs[new_key] = task
                        self._tasks.append(CropTask(days_int, task))

                    to_delete.extend([key, days_key])

        for k, v in new_attrs.items():
            setattr(self, k, v)

        for k in to_delete:
            delattr(self, k)


def load_crops():
    print("Parsing crop database...")
    df = pd.read_excel(
        MC_INPUT_FILE,
        sheet_name=MC_SHEET_NAME,
        header=3
    )
    df = df.dropna(how="all")
    df = df.sort_values(by="Culture", key=lambda col: col.astype(str).str.lower())
    crops = []
    crops_itk = parse_itk()

    for _, row in df.iterrows():
        data = {k: ("" if pd.isna(v) else str(v)) for k, v in row.to_dict().items()}
        crop_value = data.get("Culture", "")
        if crop_value in ["", "-"]:
            continue
        new_crop = Crop(**data)
        filtered_itk = [crop_itk for crop_itk in crops_itk if crop_value == crop_itk["name"]]
        if filtered_itk:
            for k, v in filtered_itk[0].items():
                if k != "name" and v:
                    if k in ["Séries", "itinéraire"] and isinstance(v, list):
                        v = "<br>".join(v)
                    setattr(new_crop, k, v)
        else:
            print(f"  # {crop_value} not found in ITK database")

        new_crop.prepare_print()
        crops.append(new_crop)

    print(f"{len(crops)} crops loaded")

    return crops
