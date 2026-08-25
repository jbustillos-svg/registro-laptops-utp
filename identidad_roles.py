"""Reglas puras para separar la validación de identidad por rol."""

ROL_ALUMNO = "ALUMNO"
ROL_DOCENTE = "DOCENTE"


def normalizar_rol(valor):
    return str(valor or ROL_ALUMNO).strip().upper() or ROL_ALUMNO


def seleccionar_flujo_identidad(rol):
    return ROL_DOCENTE if normalizar_rol(rol) == ROL_DOCENTE else ROL_ALUMNO


def rol_requiere_curp(rol):
    return seleccionar_flujo_identidad(rol) == ROL_ALUMNO
