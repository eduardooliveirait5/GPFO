from dataclasses import dataclass

@dataclass
class Presenca:
    id: int
    sessao_id: int
    email: str
    nome: str
    funcao: str
    duracao_min: float
    num_reconexoes: int