from fastapi import FastAPI, UploadFile, File, Request, HTTPException
from fastapi.responses import StreamingResponse
import pyexcel as p
import tempfile
import os

app = FastAPI(title="XLS to XLSX Converter")

@app.post("/convert")
async def convert(request: Request, file: UploadFile = None):
    """
    Soporta:
    - multipart/form-data (archivo UploadFile)
    - binario directo (Power Automate)
    """

    # Caso 1: archivo normal (multipart)
    if file:
        filename = file.filename
        content = await file.read()

    # Caso 2: binario directo (Power Automate)
    else:
        content = await request.body()
        filename = "archivo.xls"

    if not filename.lower().endswith(".xls"):
        raise HTTPException(status_code=400, detail="Debe ser un archivo .xls")

    # Guardar XLS temporal
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xls") as tmp:
        tmp.write(content)
        xls_path = tmp.name

    # Ruta de salida XLSX
    xlsx_fd, xlsx_path = tempfile.mkstemp(suffix=".xlsx")
    os.close(xlsx_fd)

    try:
        # Leer XLS con pyexcel
        sheet = p.get_sheet(file_name=xls_path)

        # Guardar como XLSX
        sheet.save_as(xlsx_path)

        # Enviar XLSX como respuesta
        return StreamingResponse(
            open(xlsx_path, "rb"),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": 'attachment; filename="convertido.xlsx"'}
        )

    finally:
        # limpiar archivos temporales
        try: os.remove(xls_path)
        except: pass
        try: os.remove(xlsx_path)
        except: pass
