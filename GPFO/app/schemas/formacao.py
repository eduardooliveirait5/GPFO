from pydantic import BaseModel

class Formacao(BaseModel):
    codigo: str
    nome: str
    criada_em: str

class FormacaoResumo(BaseModel):
    codigo: str
    nome: str