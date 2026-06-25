# -*- coding: utf-8 -*-
"""
sources/banagricola.py  —  Gestora de Fondos de Inversión Banagrícola, S.A. (1 fondo)  [PENDIENTE]

Estado: PENDIENTE de habilitar. La página gestorabanagricola.com/tablas-valoracion
pinta la tabla con un componente React; los endpoints visibles en el bundle son
del flujo de pagos (Wompi), no de valoración, y la tabla parece mostrar SOLO el
dato del día (sin serie histórica). Falta inspección de DevTools (Network) para
confirmar si existe un endpoint con histórico; si solo hay dato del día, la serie
se acumularía hacia adelante.

Cuando se resuelva, llenar get_fondos() respetando el contrato común.
"""
from __future__ import annotations

GESTORA = "Gestora Banagrícola"
GESTORA_SLUG = "banagricola"
PENDIENTE = True
MOTIVO = "Tabla cargada por JS (React); sin endpoint histórico confirmado (posible solo dato del día)."

FONDOS_CONOCIDOS = ["FIA Renta Liquidez Banagrícola"]


def get_fondos() -> list[dict]:
    return []
