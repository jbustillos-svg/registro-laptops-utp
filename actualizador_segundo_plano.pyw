"""Puente compatible con instalaciones que todavía invoquen el ayudante legado."""

from datetime import datetime
from pathlib import Path


def registrar_log(directorio, mensaje):
    ruta_log = Path(directorio) / "update.log"
    marca_tiempo = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        with ruta_log.open("a", encoding="utf-8") as archivo:
            archivo.write(f"[{marca_tiempo}] [ACTUALIZACION] {mensaje}\n")
    except OSError:
        pass


def main():
    # No espera al proceso principal ni ejecuta Git. Si una versión anterior lo
    # lanzó, convierte ese estado en la señal que atenderá el próximo bootstrap.
    directorio = Path(__file__).resolve().parent
    ruta_marker = directorio / ".actualizacion_pendiente"
    ruta_temporal = directorio / ".actualizacion_pendiente.tmp"
    try:
        ruta_temporal.write_text("pendiente\n", encoding="ascii")
        ruta_temporal.replace(ruta_marker)
        registrar_log(directorio, "ayudante legado convirtió el estado a marker pendiente")
    except OSError as error:
        registrar_log(directorio, f"ayudante legado no pudo crear marker: {error}")


if __name__ == "__main__":
    main()
