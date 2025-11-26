from fastapi import FastAPI, UploadFile, File, Request, HTTPException
from fastapi.responses import StreamingResponse
import pandas as pd
import tempfile
import os

app = FastAPI(title="XLS to XLSX Converter")

@app.post("/convert")
async def convert(request: Request, file: UploadFile = None):
    # Caso 1: archivo normal
    if file:
        filename = file.filename
        content = await file.read()
    else:
        # Caso 2: Power Automate envia binario directo
        content = await request.body()
        filename = "archivo.xls"

    if not filename.lower().endswith(".xls"):
        raise HTTPException(status_code=400, detail="Debe ser un archivo .xls")

    # Guardar XLS temporal
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xls") as tmp:
        tmp.write(content)
        xls_path = tmp.name

    # Temporal para XLSX
    xlsx_fd, xlsx_path = tempfile.mkstemp(suffix=".xlsx")
    os.close(xlsx_fd)

    try:
        df = pd.read_excel(xls_path, engine="xlrd")
        df.to_excel(xlsx_path, index=False, engine="openpyxl")

        return StreamingResponse(
            open(xlsx_path, "rb"),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": 'attachment; filename="convertido.xlsx"'}
        )

    finally:
        try: os.remove(xls_path)
        except: pass
        try: os.remove(xlsx_path)
        except: pass
