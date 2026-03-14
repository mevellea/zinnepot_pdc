import re
import pandas as pd
from itk import parse_itk

MC_INPUT_FILE = r"H:\My Drive\3.PDC\Crop_planning_4.0_ Eng_MC1.xlsx"
MC_SHEET_NAME = "Crop Chart"


class Crop:

    def __init__(self, **entries):

        for k, v in entries.items():
            setattr(self, k, v)

    def to_dict(self):
        return self.__dict__

    @property
    def crop(self):
        return getattr(self, "Culture")

    def check_dtm(self, dtm):
        if dtm is None:
            return None

    def __str__(self):
        val = f"{self.crop} "
        return val


def transform_tasks(crop_obj):
    attrs = list(vars(crop_obj).keys())
    new_attrs = {}
    to_delete = []

    for key in attrs:
        m = re.match(r"Tâche (\d+)", key)
        if m:
            i = m.group(1)

            task = getattr(crop_obj, key)
            days_key = f"# jours {i}"

            if hasattr(crop_obj, days_key):
                days = getattr(crop_obj, days_key)
                if type(days) is str and days != "":
                    new_key = f"Tâche J={int(float(days))}"
                    new_attrs[new_key] = task

                to_delete.extend([key, days_key])

    for k, v in new_attrs.items():
        setattr(crop_obj, k, v)

    for k in to_delete:
        delattr(crop_obj, k)

    return crop_obj


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
        print(crop_value)

        filtered_itk = [crop_itk for crop_itk in crops_itk if crop_value == crop_itk["name"]]
        if filtered_itk:
            for k, v in filtered_itk[0].items():
                if k != "name" and v:
                    if k in ["Séries", "itinéraire"] and isinstance(v, list):
                        v = "<br>".join(v)
                    setattr(new_crop, k, v)
        else:
            print(f"  # {crop_value} not found in ITK database")

        new_crop = transform_tasks(new_crop)
        crops.append(new_crop)

    print(f"{len(crops)} crops loaded")

    return crops
