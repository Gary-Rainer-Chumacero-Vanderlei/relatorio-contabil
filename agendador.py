"""
agendador.py
Executa o fechamento mensal automaticamente todo dia 1º de cada mês às 06:00.

Uso:
    python agendador.py                     # inicia o agendador em loop
    python agendador.py --executar-agora    # força execução imediata para teste

Requer: pip install schedule
"""

from __future__ import annotations

import argparse
import logging
import sys
from datetime import datetime
from pathlib import Path

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger("agendador")

BASE = Path(__file__).parent


def executar_fechamento():
    """Calcula o mês anterior e aciona o gerador de relatórios."""
    from dateutil.relativedelta import relativedelta

    hoje = datetime.now()
    mes_ref = (hoje - relativedelta(months=1)).strftime("%Y-%m")
    log.info("Iniciando fechamento mensal — competência %s", mes_ref)

    try:
        # Importa e executa o motor diretamente (sem subprocess)
        sys.argv = ["gerar_relatorio.py", "--mes", mes_ref, "--formato", "excel,pdf"]
        import importlib.util
        spec = importlib.util.spec_from_file_location(
            "gerar_relatorio",
            BASE / "scripts" / "gerar_relatorio.py",
        )
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        mod.main()
        log.info("Fechamento concluído com sucesso — competência %s", mes_ref)
    except Exception as exc:
        log.error("Falha no fechamento: %s", exc, exc_info=True)


def iniciar_agendador():
    try:
        import schedule
        import time
    except ImportError:
        log.error("Pacote 'schedule' não instalado. Execute: pip install schedule")
        sys.exit(1)

    # Roda todo dia 1º às 06:00 — verifica diariamente
    def tarefa_diaria():
        if datetime.now().day == 1:
            log.info("Dia 1º detectado — disparando fechamento mensal")
            executar_fechamento()

    schedule.every().day.at("06:00").do(tarefa_diaria)
    log.info("Agendador iniciado. Fechamento mensal: todo dia 1º às 06:00.")
    log.info("Aguardando próximo ciclo...")

    while True:
        schedule.run_pending()
        import time
        time.sleep(60)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Agendador de fechamento mensal")
    parser.add_argument(
        "--executar-agora",
        action="store_true",
        help="Executa o fechamento imediatamente (útil para testes)",
    )
    args = parser.parse_args()

    if args.executar_agora:
        log.info("Modo de execução imediata ativado")
        executar_fechamento()
    else:
        iniciar_agendador()
