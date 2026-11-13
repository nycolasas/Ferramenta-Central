"""Compatibilidade para o entrypoint da interface em pacote raiz.

A implementação funcional foi mantida em `src.ui.interface_executor` para manter a
estrutura de módulos organizada, mas este wrapper preserva a importação antiga e a
execução via `python -m ui.interface_executor`.
"""

from src.ui.interface_executor import AppExecutor, main

__all__ = ["AppExecutor", "main"]

if __name__ == "__main__":
    main()
