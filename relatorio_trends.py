"""
RELATÓRIO SEMANAL - TENDÊNCIAS DE BUSCA (Google Trends) para produtos Vimacedo
================================================================================

O que este script faz:
1. Lê a lista de categorias de produtos (categorias.csv)
2. Para cada categoria, consulta o Google Trends (últimos 90 dias, Brasil)
   usando a biblioteca não-oficial `pytrends`
3. Calcula a variação entre a média das últimas 4 semanas vs as 4 semanas
   anteriores -> isso indica se a busca está subindo, estável ou caindo
4. Gera um arquivo Excel (.xlsx) ranqueado, com a coluna "Sinal de Compra"
5. Salva o arquivo com a data da semana no nome

Como rodar (na sua máquina, não dentro do Claude):
    pip install pytrends pandas openpyxl
    python relatorio_trends.py

Como automatizar (escolha uma opção):

OPÇÃO A - Cron (Linux/Mac) ou Task Scheduler (Windows):
    Rodar toda segunda-feira às 8h:
    crontab -e
    0 8 * * 1 cd /caminho/do/script && python3 relatorio_trends.py

OPÇÃO B - GitHub Actions (gratuito, na nuvem, não depende do seu PC ligado):
    Crie um repositório privado no GitHub, suba este script + categorias.csv,
    e adicione um workflow (.github/workflows/weekly.yml) com "schedule: cron".
    O resultado pode ser enviado por e-mail automaticamente (ver função
    enviar_email() abaixo, basta preencher as credenciais SMTP).

OPÇÃO C - Mais simples: rodar manualmente toda semana e me enviar o resultado
    aqui no chat para eu te ajudar a interpretar/cruzar com sua tabela de preços.
"""

import time
import pandas as pd
from datetime import datetime

try:
    from pytrends.request import TrendReq
except ImportError:
    raise SystemExit(
        "Faltou instalar a biblioteca. Rode: pip install pytrends pandas openpyxl"
    )

CSV_CATEGORIAS = "categorias.csv"
GEO = "BR"  # Brasil. Use "" para mundial.
TIMEFRAME = "today 3-m"  # últimos 90 dias
PAUSA_ENTRE_REQUISICOES = 30  # segundos - aumentado para evitar bloqueio 429 do Google
TIMEOUT_REQUISICAO = (10, 25)  # (conexão, leitura) em segundos - evita travar pra sempre
LIMITE_CATEGORIAS = None  # defina um número (ex: 10) para testar rápido, ou None para todas
MAX_TENTATIVAS = 3  # quantas vezes tenta de novo se der erro 429
PAUSA_APOS_429 = 60  # segundos de espera extra quando é bloqueado, antes de tentar de novo


def classificar_tendencia(serie):
    """Compara a média das últimas 4 semanas com as 4 semanas anteriores."""
    if len(serie) < 8 or serie.sum() == 0:
        return None, None, "Dados insuficientes"

    volume_medio = round(serie.mean(), 1)
    ultimas_4 = serie[-4:].mean()
    anteriores_4 = serie[-8:-4].mean()

    if anteriores_4 == 0:
        variacao = 100.0 if ultimas_4 > 0 else 0.0
    else:
        variacao = ((ultimas_4 - anteriores_4) / anteriores_4) * 100

    if variacao >= 20:
        obs = "🔼 Forte alta"
    elif variacao >= 5:
        obs = "↗ Leve alta"
    elif variacao <= -20:
        obs = "🔽 Forte queda"
    elif variacao <= -5:
        obs = "↘ Leve queda"
    else:
        obs = "➡ Estável"

    return volume_medio, round(variacao, 1), obs


def gerar_relatorio():
    categorias = pd.read_csv(CSV_CATEGORIAS)
    if LIMITE_CATEGORIAS:
        categorias = categorias.head(LIMITE_CATEGORIAS)

    pytrends = TrendReq(hl="pt-BR", tz=180, timeout=TIMEOUT_REQUISICAO, retries=1, backoff_factor=0.5)

    resultados = []
    total = len(categorias)

    for i, row in categorias.iterrows():
        categoria = row["categoria"]
        termo = row["termo_busca_google"]
        categoria_ml = row.get("categoria_mercado_livre", "Outros")
        print(f"[{i+1}/{total}] Consultando: {termo}")

        tentativa = 0
        sucesso = False

        while tentativa < MAX_TENTATIVAS and not sucesso:
            tentativa += 1
            try:
                pytrends.build_payload([termo], timeframe=TIMEFRAME, geo=GEO)
                dados = pytrends.interest_over_time()

                if dados.empty:
                    resultados.append({
                        "Categoria ML": categoria_ml,
                        "Categoria": categoria,
                        "Termo pesquisado": termo,
                        "Volume de Busca (0-100)": None,
                        "Variação (%)": None,
                        "Observação": "Sem dados no Google Trends",
                    })
                    sucesso = True
                    continue

                serie = dados[termo]
                volume, variacao, obs = classificar_tendencia(serie)

                resultados.append({
                    "Categoria ML": categoria_ml,
                    "Categoria": categoria,
                    "Termo pesquisado": termo,
                    "Volume de Busca (0-100)": volume,
                    "Variação (%)": variacao,
                    "Observação": obs,
                })
                sucesso = True

            except Exception as e:
                if "429" in str(e) and tentativa < MAX_TENTATIVAS:
                    print(f"  Bloqueado (429). Aguardando {PAUSA_APOS_429}s para tentar de novo...")
                    time.sleep(PAUSA_APOS_429)
                else:
                    resultados.append({
                        "Categoria ML": categoria_ml,
                        "Categoria": categoria,
                        "Termo pesquisado": termo,
                        "Volume de Busca (0-100)": None,
                        "Variação (%)": None,
                        "Observação": f"Erro: {e}",
                    })
                    sucesso = True

        time.sleep(PAUSA_ENTRE_REQUISICOES)

    df = pd.DataFrame(resultados)
    # Ordenar por volume de busca (prioridade) e variação como desempate
    df = df.sort_values(
        by=["Volume de Busca (0-100)", "Variação (%)"],
        ascending=[False, False],
        na_position="last"
    )

    nome_arquivo = f"relatorio_trends_{datetime.now().strftime('%Y-%m-%d')}.xlsx"

    with pd.ExcelWriter(nome_arquivo, engine="openpyxl") as writer:
        # Página 1: Resumo geral, todos os produtos ordenados pelo sinal de compra
        df.to_excel(writer, sheet_name="Resumo", index=False)

        # Uma página por categoria do Mercado Livre
        for categoria_ml in df["Categoria ML"].dropna().unique():
            df_categoria = df[df["Categoria ML"] == categoria_ml]
            # Nome da aba não pode ter mais de 31 caracteres nem certos símbolos
            nome_aba = str(categoria_ml)[:31]
            df_categoria.to_excel(writer, sheet_name=nome_aba, index=False)

    print(f"\nRelatório salvo em: {nome_arquivo}")
    return nome_arquivo


def enviar_email(arquivo, destinatario, remetente, senha_app, servidor_smtp="smtp.gmail.com", porta=587):
    """
    Opcional: envia o relatório por e-mail automaticamente.
    Para Gmail, gere uma 'senha de app' em myaccount.google.com/apppasswords
    """
    import smtplib
    from email.message import EmailMessage

    msg = EmailMessage()
    msg["Subject"] = f"Relatório semanal de tendências - {datetime.now().strftime('%d/%m/%Y')}"
    msg["From"] = remetente
    msg["To"] = destinatario
    msg.set_content("Segue em anexo o relatório semanal de tendências de produtos.")

    with open(arquivo, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=arquivo,
        )

    with smtplib.SMTP(servidor_smtp, porta) as smtp:
        smtp.starttls()
        smtp.login(remetente, senha_app)
        smtp.send_message(msg)

    print(f"E-mail enviado para {destinatario}")


if __name__ == "__main__":
    import os

    arquivo_gerado = gerar_relatorio()

    # No GitHub Actions, essas variáveis vêm dos "Secrets" do repositório.
    destinatario = os.environ.get("EMAIL_TO")
    remetente = os.environ.get("EMAIL_FROM")
    senha_app = os.environ.get("EMAIL_APP_PASSWORD")

    if destinatario and remetente and senha_app:
        enviar_email(
            arquivo=arquivo_gerado,
            destinatario=destinatario,
            remetente=remetente,
            senha_app=senha_app,
        )
    else:
        print("Variáveis de e-mail não configuradas — relatório gerado mas não enviado.")
