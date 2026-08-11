from dataclasses import dataclass

@dataclass
class Sessao:
    id: int
    formacao_codigo: str
    titulo: str
    data: str
    hora_inicio_janela: str | None
    hora_fim_janela: str | None
    duracao_sessao_min: float
    ficheiro_original: str | None
    importada_em: str