import os
from pathlib import Path

from fastapi import Body, FastAPI
from fastapi.responses import FileResponse, JSONResponse

from .excel import ExcelBusy
from .service import Service, ServiceError

ROOT = Path(__file__).resolve().parent.parent
XLSX = Path(os.environ.get("PILARDIUM_XLSX", ROOT / "data" / "Costos.xlsx"))
svc = Service(XLSX, ROOT / "data" / "state.json", ROOT / "data" / "backups")
app = FastAPI(title="Pilardium")


@app.exception_handler(ServiceError)
async def _service_error(_, e: ServiceError):
    return JSONResponse({"detail": str(e)}, status_code=e.status)


@app.exception_handler(ExcelBusy)
async def _busy(_, e: ExcelBusy):
    return JSONResponse({"detail": str(e)}, status_code=423)


def done():
    return svc.state()


@app.get("/api/state")
def state():
    return svc.state()


@app.post("/api/insumos")
def insumo_create(p: dict = Body(...)):
    svc.insumo_create(p)
    return done()


@app.post("/api/insumos/link")
def insumo_link(p: dict = Body(...)):
    svc.insumo_link(p)
    return done()


@app.put("/api/insumos/{iid}")
def insumo_update(iid: str, p: dict = Body(...)):
    svc.insumo_update(iid, p)
    return done()


@app.delete("/api/insumos/{iid}")
def insumo_delete(iid: str):
    svc.insumo_delete(iid)
    return done()


@app.post("/api/insumos/{iid}/buy")
def buy(iid: str, p: dict = Body(...)):
    svc.buy(iid, p)
    return done()


@app.post("/api/insumos/{iid}/adjust")
def adjust(iid: str, p: dict = Body(...)):
    svc.adjust(iid, p)
    return done()


@app.get("/api/recipes/{rid}/preview")
def preview(rid: str, kg: float = 1.0):
    return svc.preview(rid, kg)


@app.post("/api/recipes/{rid}/produce")
def produce(rid: str, p: dict = Body(...)):
    svc.produce(rid, p)
    return done()


@app.post("/api/recipes/{rid}/sell")
def sell(rid: str, p: dict = Body(...)):
    svc.sell(rid, p)
    return done()


@app.put("/api/recipes/{rid}/min")
def mil_min(rid: str, p: dict = Body(...)):
    svc.mil_min(rid, p)
    return done()


@app.post("/api/recipes")
def recipe_create(p: dict = Body(...)):
    rid = svc.recipe_create(p)
    return {**svc.state(), "created": rid}


@app.delete("/api/recipes/{rid}")
def recipe_delete(rid: str):
    svc.recipe_delete(rid)
    return done()


@app.put("/api/recipes/{rid}")
def recipe_save(rid: str, p: dict = Body(...)):
    svc.recipe_save(rid, p)
    return done()


@app.put("/api/params")
def params(p: dict = Body(...)):
    svc.params_save(p)
    return done()


@app.put("/api/fixed-costs")
def fixed(p: dict = Body(...)):
    svc.fixed_save(p)
    return done()


@app.post("/api/moves/{mid}/undo")
def undo(mid: str):
    svc.undo(mid)
    return done()


@app.get("/api/download")
def download():
    return FileResponse(XLSX, filename="Costos.xlsx")


@app.get("/")
def index():
    return FileResponse(ROOT / "static" / "index.html")
