"""Compatibilidade para imports legados do projeto.

A lógica real de caminhos e arquivos base foi movida para src.paths para manter
uma estrutura mais organizada e reutilizável, mas preservamos este módulo para
não quebrar scripts antigos ou o empacotamento.
"""

from src.paths import *  # noqa: F401,F403
