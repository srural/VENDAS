#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Atalho CLI para execução da rotina de sincronização de banco de dados.
Uso:
    python db_sync.py                                (Modo Dry-Run / Simulação)
    python db_sync.py --apply                        (Aplica as alterações no banco de destino)
    python db_sync.py --source-db=vendas_db --target-db=vendas_prod_db --apply
    python db_sync.py --export-sql=migracao.sql
"""

if __name__ == "__main__":
    from app.db_sync import main
    main()
