"""Compatibilidade para o entrypoint histórico do projeto.

A implementação real da interface foi movida para `ui.interface_executor`,
mas mantemos este arquivo para garantir que o comando `python interface_executor.py`
continue funcionando sem quebrar a organização do projeto.
"""

from ui.interface_executor import AppExecutor, main

if __name__ == "__main__":
    main()

__all__ = ["AppExecutor", "main"]
