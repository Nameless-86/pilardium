"""Business logic. It joins the workbook (recipes, prices) with the stock store."""
import threading
import time

from . import excel as X
from .excel import conv, norm
from .store import Store


class ServiceError(Exception):
    def __init__(self, msg, status=400):
        super().__init__(msg)
        self.status = status


def _level(qty, minimum):
    return 2 if qty <= 0 else (1 if minimum > 0 and qty < minimum else 0)


def _month(ts):
    t = time.localtime(ts)
    return t.tm_year, t.tm_mon


class Service:
    def __init__(self, xlsx, state_path, backup_dir):
        self.lock = threading.RLock()
        self.book = X.ExcelBook(xlsx, backup_dir)
        self.store = Store(state_path)

    # ---------- helpers ----------
    def _resolver(self):
        def resolve(name):
            ins = self.store.resolve(name)
            return (ins["unit"], ins["price"]) if ins else None
        return resolve

    def _commit(self, wb, model):
        """Sync insumo prices into the workbook, then save it when it changed."""
        X.sync_prices(wb, model, self._resolver())
        self.book.save(wb)

    def _load(self):
        wb = self.book.open()
        model = X.read_model(wb)
        self.store.seed(model)
        return wb, model

    def _consumption(self, model, rec, kg):
        """List the insumos that a production of kg of raw meat uses."""
        need, skipped = {}, []
        rows = model["general"]["rows"] + rec["rows"]
        for r in rows:
            if not r["name"] or not r["qty"]:
                continue
            ins = self.store.resolve(r["name"])
            q = conv(r["qty"] * kg, r["unit"], ins["unit"]) if ins else None
            if q is None:
                skipped.append(r["name"])
                continue
            need[ins["id"]] = need.get(ins["id"], 0.0) + q
        lines = []
        for iid, q in need.items():
            ins = self.store.insumo(iid)
            lines.append({"id": iid, "name": ins["name"], "unit": ins["unit"], "need": q, "have": ins["qty"]})
        return lines, skipped

    # ---------- state ----------
    def state(self):
        with self.lock:
            wb, model = self._load()
            st = self.store
            used = {}
            for rec in [{"name": "Costos generales", "rows": model["general"]["rows"]}] + model["recipes"]:
                for r in rec["rows"]:
                    ins = st.resolve(r["name"]) if r["name"] else None
                    if ins:
                        used.setdefault(ins["id"], set()).add(rec["name"])
            insumos = [dict(i, level=_level(i["qty"], i["min"]), value=i["qty"] * i["price"],
                            used_in=sorted(used.get(i["id"], []))) for i in st.data["insumos"]]
            unlinked = sorted({r["name"] for rec in [model["general"]] + model["recipes"]
                               for r in rec["rows"] if r["name"] and not st.resolve(r["name"])})

            def decorate(rows):
                out = []
                for r in rows:
                    if not r["name"] and is_blank(r):
                        continue
                    ins = st.resolve(r["name"]) if r["name"] else None
                    flag = None
                    if r["name"] and not ins:
                        flag = "sin vincular"
                    elif r["name"] and r["qty"] is None:
                        flag = "falta cantidad"
                    elif r["name"] and r["cost"] is None:
                        flag = "falta precio"
                    elif ins and conv(1, ins["unit"], r["unit"]) is None:
                        flag = "unidad incompatible"
                    out.append(dict(r, insumo_id=ins["id"] if ins else None,
                                    line=X.line_cost(r), flag=flag))
                return out

            is_blank = X.is_free
            recipes, mil_value = [], 0.0
            for rec in model["recipes"]:
                if rec["hidden"]:
                    continue
                m = st.mil(rec["id"])
                rows = decorate(rec["rows"])
                warn = [f"{r['name'] or 'fila'}: {r['flag']}" for r in rows if r["flag"]]
                if not rec["units_per_kg"]:
                    warn.append("faltan el peso final del lote y el peso por milanesa")
                if not any(r["name"] for r in rows):
                    warn.append("la receta no tiene ingredientes")
                mil_value += m["qty"] * rec["cost"]
                recipes.append(dict(rec, rows=rows, qty=m["qty"], min=m["min"], level=_level(m["qty"], m["min"]),
                                    value=m["qty"] * rec["cost"], warnings=warn))
            general = dict(model["general"], rows=decorate(model["general"]["rows"]))

            now = _month(time.time())
            buys = sells = sell_cost = sell_units = 0.0
            for mv in st.data["moves"]:
                if mv["undone"] or _month(mv["ts"]) != now:
                    continue
                if mv["type"] == "buy":
                    buys += mv.get("cost") or 0
                elif mv["type"] == "sell":
                    sells += mv["qty"] * (mv.get("price") or 0)
                    sell_cost += mv["qty"] * (mv.get("unit_cost") or 0)
                    sell_units += mv["qty"]
            cap = model["capital"]
            return {
                "insumos": insumos, "unlinked": unlinked, "recipes": recipes, "general": general,
                "params": model["params"], "fixed": model["fixed"], "capital": cap,
                "breakeven": model["breakeven"],
                "month": {"spend": buys, "sales": sells, "sales_cost": sell_cost, "units_sold": sell_units,
                          "gross": sells - sell_cost},
                "stock_value": {"insumos": sum(i["value"] for i in insumos), "milanesas": mil_value},
                "moves": list(reversed(st.data["moves"][-400:])),
                "file": str(self.book.path),
            }

    # ---------- insumos ----------
    def insumo_create(self, p):
        with self.lock:
            name = str(p.get("name", "")).strip()
            if not name:
                raise ServiceError("Falta el nombre del insumo.")
            if self.store.resolve(name):
                raise ServiceError("Ya existe un insumo con ese nombre.")
            self.store.add_insumo(name, p.get("unit") or "kg", float(p.get("price") or 0),
                                  float(p.get("qty") or 0), float(p.get("min") or 0))
            self.store.save()

    def insumo_update(self, iid, p):
        with self.lock:
            ins = self._ins(iid)
            wb, model = self._load()
            if "name" in p and str(p["name"]).strip():
                ins["name"] = str(p["name"]).strip()
                self.store.data["links"][norm(ins["name"])] = iid
            if "unit" in p:
                ins["unit"] = p["unit"]
            if "price" in p:
                ins["price"] = float(p["price"] or 0)
            if "min" in p:
                ins["min"] = float(p["min"] or 0)
            self.store.save()
            self._commit(wb, model)

    def insumo_delete(self, iid):
        with self.lock:
            ins = self._ins(iid)
            wb, model = self._load()
            for rec in [model["general"]] + model["recipes"]:
                for r in rec["rows"]:
                    if r["name"] and self.store.resolve(r["name"]) is ins:
                        raise ServiceError("El insumo se usa en una receta. Sacalo de la receta primero.", 409)
            self.store.data["insumos"].remove(ins)
            self.store.data["links"] = {k: v for k, v in self.store.data["links"].items() if v != iid}
            self.store.save()

    def insumo_link(self, p):
        with self.lock:
            name = str(p.get("name", "")).strip()
            if p.get("insumo_id"):
                ins = self._ins(p["insumo_id"])
                self.store.data["links"][norm(name)] = ins["id"]
            else:
                wb, model = self._load()
                unit = "unidad"
                for rec in [model["general"]] + model["recipes"]:
                    for r in rec["rows"]:
                        if norm(r["name"]) == norm(name):
                            unit = X.default_insumo_unit(r["unit"])
                self.store.add_insumo(name, unit)
            self.store.save()
            wb, model = self._load()
            self._commit(wb, model)

    def _ins(self, iid):
        ins = self.store.insumo(iid)
        if not ins:
            raise ServiceError("Insumo no encontrado.", 404)
        return ins

    def buy(self, iid, p):
        with self.lock:
            ins = self._ins(iid)
            qty = float(p.get("qty") or 0)
            if qty <= 0:
                raise ServiceError("La cantidad debe ser mayor que cero.")
            cost = float(p["cost"]) if p.get("cost") not in (None, "") else round(ins["price"] * qty)
            prev, new = ins["price"], ins["price"]
            ins["qty"] += qty
            if p.get("update_price") and cost > 0:
                new = ins["price"] = round(cost / qty, 4)
            self.store.add_move(type="buy", label=ins["name"], qty=qty, unit=ins["unit"], cost=cost,
                                deltas=[{"k": "i", "id": iid, "d": qty}], price_chg=[iid, prev, new])
            self.store.save()
            if new != prev:
                wb, model = self._load()
                self._commit(wb, model)

    def adjust(self, iid, p):
        with self.lock:
            ins = self._ins(iid)
            val = float(p.get("value") or 0)
            delta = val - ins["qty"] if p.get("mode") == "set" else val
            delta = max(delta, -ins["qty"])
            ins["qty"] += delta
            self.store.add_move(type="adjust", label=ins["name"], qty=delta, unit=ins["unit"],
                                deltas=[{"k": "i", "id": iid, "d": delta}])
            self.store.save()

    # ---------- milanesas ----------
    def preview(self, rid, kg):
        with self.lock:
            wb, model = self._load()
            rec = self._rec(model, rid)
            lines, skipped = self._consumption(model, rec, kg)
            return {"lines": lines, "skipped": skipped, "units": rec["units_per_kg"] * kg}

    def _rec(self, model, rid):
        rec = X.find_block(model, rid)
        if not rec:
            raise ServiceError("Receta no encontrada.", 404)
        return rec

    def produce(self, rid, p):
        with self.lock:
            wb, model = self._load()
            rec = self._rec(model, rid)
            kg = float(p.get("kg") or 0)
            if kg <= 0:
                raise ServiceError("Indicá cuántos kg de carne se procesan.")
            units = float(p["units"]) if p.get("units") not in (None, "") else round(rec["units_per_kg"] * kg)
            if units <= 0:
                raise ServiceError("Indicá cuántas milanesas se obtuvieron.")
            lines, _ = self._consumption(model, rec, kg)
            deltas = []
            for ln in lines:
                ins = self.store.insumo(ln["id"])
                used = min(ins["qty"], ln["need"])
                ins["qty"] -= used
                deltas.append({"k": "i", "id": ins["id"], "d": -used})
            self.store.mil(rid)["qty"] += units
            deltas.append({"k": "m", "id": rid, "d": units})
            self.store.add_move(type="prod", label=rec["name"], qty=units, unit="u", kg=kg,
                                unit_cost=rec["cost"], deltas=deltas)
            self.store.save()

    def sell(self, rid, p):
        with self.lock:
            wb, model = self._load()
            rec = self._rec(model, rid)
            qty = float(p.get("qty") or 0)
            m = self.store.mil(rid)
            if qty <= 0:
                raise ServiceError("La cantidad debe ser mayor que cero.")
            if qty > m["qty"] + 1e-9:
                raise ServiceError(f"Solo hay {m['qty']:g} unidades en stock de {rec['name']}.")
            price = float(p["price"]) if p.get("price") not in (None, "") else (rec["price"] or 0)
            m["qty"] -= qty
            self.store.add_move(type="sell", label=rec["name"], qty=qty, unit="u", price=price,
                                unit_cost=rec["cost"], deltas=[{"k": "m", "id": rid, "d": -qty}])
            self.store.save()

    def mil_min(self, rid, p):
        with self.lock:
            self.store.mil(rid)["min"] = float(p.get("min") or 0)
            self.store.save()

    # ---------- workbook edits ----------
    def recipe_save(self, rid, p):
        with self.lock:
            wb, model = self._load()
            rows = []
            for r in p.get("rows", []):
                name = str(r.get("name", "")).strip()
                if name:
                    rows.append(r)
                    if not self.store.resolve(name):
                        raise ServiceError(f"El ingrediente '{name}' no es un insumo. Creá el insumo primero.")
            try:
                X.write_recipe(wb, model, rid, rows, p.get("peso_final"), p.get("peso_unidad"),
                               p.get("price"), "price" in p, p.get("name"))
            except ValueError as e:
                raise ServiceError(str(e))
            # reread so the cost column matches the new ingredient names, then sync prices
            tmp_model = X.read_model(wb)
            self._commit(wb, tmp_model)

    def recipe_create(self, p):
        with self.lock:
            name = str(p.get("name", "")).strip()
            cat = str(p.get("cat", "")).strip()
            if not name or not cat:
                raise ServiceError("Indicá la carne y el nombre de la receta.")
            wb, model = self._load()
            if any(norm(r["name"]) == norm(name) for r in model["recipes"] if r["cat"] == norm(cat)):
                raise ServiceError("Ya existe una receta con ese nombre en esa carne.")
            try:
                rid = X.create_recipe(wb, model, cat, name)
            except ValueError as e:
                raise ServiceError(str(e))
            self._commit(wb, X.read_model(wb))
            return rid

    def recipe_delete(self, rid):
        with self.lock:
            wb, model = self._load()
            self._rec(model, rid)
            if self.store.mil(rid)["qty"] > 0:
                raise ServiceError("La receta tiene stock de milanesas. Vendé o ajustá el stock primero.", 409)
            X.delete_recipe(wb, model, rid)
            self.store.data["mil"].pop(rid, None)
            self.store.save()
            self.book.save(wb)

    def params_save(self, p):
        with self.lock:
            wb, model = self._load()
            com, des = float(p.get("comision") or 0), float(p.get("desperdicio") or 0)
            if not (0 <= com < 1 and 0 <= des < 1):
                raise ServiceError("Los porcentajes deben estar entre 0 y 100.")
            X.write_params(wb, com, des)
            self.book.save(wb)

    def fixed_save(self, p):
        with self.lock:
            wb, model = self._load()
            X.write_fixed(wb, [r for r in p.get("rows", []) if str(r.get("name", "")).strip()])
            self.book.save(wb)

    # ---------- undo ----------
    def undo(self, mid):
        with self.lock:
            mv = next((m for m in self.store.data["moves"] if m["id"] == mid), None)
            if not mv or mv["undone"]:
                return
            targets = []
            for d in mv["deltas"]:
                obj = self.store.insumo(d["id"]) if d["k"] == "i" else self.store.mil(d["id"])
                if obj is None:
                    raise ServiceError("No se puede deshacer: el insumo ya no existe.")
                if obj["qty"] - d["d"] < -1e-9:
                    raise ServiceError("No se puede deshacer: ya se usó o vendió parte de ese stock.")
                targets.append((obj, d["d"]))
            for obj, d in targets:
                obj["qty"] = max(0.0, obj["qty"] - d)
            mv["undone"] = True
            chg = mv.get("price_chg")
            if chg:
                ins = self.store.insumo(chg[0])
                if ins and ins["price"] == chg[2] and chg[1] != chg[2]:
                    ins["price"] = chg[1]
                    self.store.save()
                    wb, model = self._load()
                    self._commit(wb, model)
            self.store.save()
