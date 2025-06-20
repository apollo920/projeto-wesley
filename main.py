from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from bot_tendencias import executar_bot

app = FastAPI()

# Permitir acesso do frontend
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"]
)

@app.get("/")
def home():
    return {"message": "API funcionando"}

@app.post("/enviar-relatorio")
def enviar():
    try:
        executar_bot()
        return {"status": "Relatório enviado com sucesso"}
    except Exception as e:
        return {"status": f"Erro: {e}"}
