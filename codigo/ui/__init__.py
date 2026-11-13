"""Módulo da interface gráfica do projeto.

Mantém a UI em um pacote dedicado para separar a camada de apresentação do
restante da automação.
"""

from .interface_executor import AppExecutor, main

__all__ = ["AppExecutor", "main"]
