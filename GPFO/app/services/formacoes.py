# ═══════════════════════════════════════════════════════════
#  BASE DE DADOS DAS FORMAÇÕES
#  Serviços relacionados às formações, sessões e presenças.
# ═══════════════════════════════════════════════════════════

from app.database import get_conn


def listar_formacoes():

    with get_conn() as conn:

        linhas = conn.execute(
            """
            SELECT *

            FROM formacoes

            ORDER BY codigo
            """
        ).fetchall()

    resultado = []

    for linha in linhas:

        resultado.append(dict(linha))

    return resultado


def obter_formacao(codigo):

    with get_conn() as conn:

        linha = conn.execute(
            """
            SELECT *

            FROM formacoes

            WHERE codigo=?
            """,

            (codigo,)
        ).fetchone()

    if linha is None:

        return None

    return dict(linha)


def listar_sessoes(codigo):

    with get_conn() as conn:

        linhas = conn.execute(
            """
            SELECT *

            FROM sessoes

            WHERE formacao_codigo=?

            ORDER BY data
            """,

            (codigo,)
        ).fetchall()

    resultado = []

    for linha in linhas:

        resultado.append(dict(linha))

    return resultado


def apagar_sessao(sessao_id):

    with get_conn() as conn:

        conn.execute(

            """
            DELETE FROM presencas

            WHERE sessao_id=?

            """,

            (sessao_id,)
        )

        conn.execute(

            """
            DELETE FROM sessoes

            WHERE id=?

            """,

            (sessao_id,)
        )

        conn.commit()


def apagar_formacao(codigo):

    with get_conn() as conn:

        sessoes = conn.execute(

            """
            SELECT id

            FROM sessoes

            WHERE formacao_codigo=?

            """,

            (codigo,)
        ).fetchall()

        for sessao in sessoes:

            conn.execute(

                """
                DELETE FROM presencas

                WHERE sessao_id=?

                """,

                (sessao["id"],)
            )

        conn.execute(

            """
            DELETE FROM sessoes

            WHERE formacao_codigo=?

            """,

            (codigo,)
        )

        conn.execute(

            """
            DELETE FROM formacoes

            WHERE codigo=?

            """,

            (codigo,)
        )

        conn.commit()