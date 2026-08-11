from pydantic import BaseModel

class Presenca(BaseModel):
    id: int
    sessao_id: int
    email: str
    nome: str
    funcao: str
    duracao_min: float
    num_reconexoes: int