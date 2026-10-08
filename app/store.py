"""Stock state: insumos, milanesa stock and the movement log. Kept in one JSON file."""
import json
import os
import threading
import time
from pathlib import Path

from .excel import default_insumo_unit, norm

SEED_NAMES = {
    "pechuga": "Pechuga de pollo", "pechuga de pollo": "Pechuga de pollo", "carne vaca": "Carne de vaca",
    "cerdo": "Carne de cerdo", "rebozdor": "Rebozador", "pimenton": "Pimentón dulce",
    "pimenton dulce": "Pimentón dulce", "pimienta": "Pimienta negra", "pimienta negra": "Pimienta negra",
}


class Store:
    def __init__(self, path):
        self.path = Path(path)
        self.lock = threading.RLock()
        self.data = {"seeded": False, "next": 1, "insumos": [], "links": {}, "mil": {}, "moves": []}
        if self.path.exists():
            self.data.update(json.loads(self.path.read_text(encoding="utf-8")))

    def save(self):
        tmp = self.path.with_suffix(".tmp")
        tmp.write_text(json.dumps(self.data, ensure_ascii=False, indent=1), encoding="utf-8")
        os.replace(tmp, self.path)

    def new_id(self, prefix):
        n = self.data["next"]
        self.data["next"] = n + 1
        return f"{prefix}{n}"

    def insumo(self, iid):
        return next((i for i in self.data["insumos"] if i["id"] == iid), None)

    def resolve(self, name):
        iid = self.data["links"].get(norm(name))
        return self.insumo(iid) if iid else None

    def add_insumo(self, name, unit, price=0.0, qty=0.0, minimum=0.0):
        ins = {"id": self.new_id("i"), "name": name.strip(), "unit": unit, "price": price, "qty": qty, "min": minimum}
        self.data["insumos"].append(ins)
        self.data["links"][norm(name)] = ins["id"]
        return ins

    def mil(self, rid):
        return self.data["mil"].setdefault(rid, {"qty": 0.0, "min": 0.0})

    def add_move(self, **kw):
        mv = {"id": self.new_id("m"), "ts": time.time(), "undone": False, **kw}
        self.data["moves"].append(mv)
        del self.data["moves"][:-5000]
        return mv

    def seed(self, model):
        """Create one insumo for each distinct ingredient name found in the workbook."""
        from .excel import conv
        if self.data["seeded"]:
            return
        rows = list(model["general"]["rows"]) + [r for rec in model["recipes"] for r in rec["rows"]]
        for row in rows:
            if not row["name"]:
                continue
            key = norm(row["name"])
            canon = SEED_NAMES.get(key, row["name"].strip()[:1].upper() + row["name"].strip()[1:])
            ins = self.resolve(canon) or self.resolve(row["name"])
            if not ins:
                ins = self.add_insumo(canon, default_insumo_unit(row["unit"]))
            self.data["links"][key] = ins["id"]
            if not ins["price"] and row["cost"]:
                per = conv(1, ins["unit"], row["unit"])
                if per:
                    ins["price"] = round(row["cost"] * per, 4)
        self.data["seeded"] = True
        self.save()
