from fastapi import FastAPI, File, UploadFile, Form
from pydantic import BaseModel
import rpa_script  # Importa o script de RPA
import shutil

app = FastAPI()

class RPARequest(BaseModel):
    email: str
    senha: str

@app.post("/run-rpa")
async def run_rpa(
    email: str = Form(...),
    senha: str = Form(...),
    file: UploadFile = File(...)
):
    try:
        # Salvar o arquivo enviado
        file_location = f"/tmp/{file.filename}"
        with open(file_location, "wb") as f:
            shutil.copyfileobj(file.file, f)
        
        rpa_script.run_rpa_script(file_location, email, senha)  # Chama a função que executa o script de RPA
        return {"message": "RPA script executed successfully!"}
    except Exception as e:
        return {"error": str(e)}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000, reload=True)
