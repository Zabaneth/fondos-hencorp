# -*- coding: utf-8 -*-
"""
sources/atlantida.py  —  Atlántida Capital, S.A. (8 fondos)  [PENDIENTE]

Estado: PENDIENTE de habilitar. El sitio atlantidacapital.com.sv está detrás de
un WAF (F5 BIG-IP: responde "The requested URL was rejected / support ID …") y
además presenta una cadena TLS incompleta, por lo que un scraper simple
(requests) es rechazado. Falta inspección de DevTools (pestaña Network) en una
página de fondo para identificar el request real de datos y si trae histórico.

Cuando se resuelva, llenar get_fondos() respetando el contrato común
(ver sources/sgb.py) y quitar PENDIENTE.
"""
from __future__ import annotations

GESTORA = "Atlántida Capital"
GESTORA_SLUG = "atlantida"
PENDIENTE = True
MOTIVO = "Sitio tras WAF (F5) + cadena TLS incompleta; pendiente identificar endpoint de datos."

# Catálogo conocido (8 fondos) para tenerlo listo:
FONDOS_CONOCIDOS = [
    "FIA Atlántida de Liquidez a Corto Plazo",
    "FIA Atlántida de Crecimiento a Mediano Plazo",
    "FIA Atlántida Renta Mixta",
    "FIC Inmobiliario Atlántida Progresa+",
    "FIC de Capital de Riesgo Atlántida",
    "FIC de Capital de Riesgo Atlántida Empresarial+",
    "FIC Financiero Atlántida Renta Variable",
    "FIC Financiero Atlántida Renta Variable II",
]


def get_fondos() -> list[dict]:
    # Aún no implementado: el orquestador lo marcará 'pendiente'.
    return []
