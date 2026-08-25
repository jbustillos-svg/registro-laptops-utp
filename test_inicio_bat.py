import os
from pathlib import Path
import shutil
import subprocess
import tempfile
import threading
import time
import unittest


class PruebasInicioBat(unittest.TestCase):
    def setUp(self):
        self.temporal = tempfile.TemporaryDirectory()
        self.raiz = Path(self.temporal.name)
        shutil.copy2(Path(__file__).with_name("inicio_laptop.bat"), self.raiz)
        (self.raiz / "bootstrap_laptop.pyw").touch()
        self.env = os.environ.copy()
        system32 = Path(os.environ["SystemRoot"]) / "System32"
        valores = {
            "PATH": str(system32),
            "LocalAppData": str(self.raiz / "local"),
            "ProgramFiles": str(self.raiz / "pf"),
            "ProgramFiles(x86)": str(self.raiz / "pfx86"),
            "REGISTRO_LAPTOP_LOCALAPPDATA": str(self.raiz / "local"),
            "REGISTRO_LAPTOP_PROGRAMFILES": str(self.raiz / "pf"),
            "REGISTRO_LAPTOP_PROGRAMFILES_X86": str(self.raiz / "pfx86"),
            "REGISTRO_LAPTOP_WINDOWS": str(self.raiz / "windows"),
            "REGISTRO_LAPTOP_SOLO_BUSCAR": "1",
            "REGISTRO_LAPTOP_REINTENTO_SEGUNDOS": "1",
        }
        for nombre, valor in valores.items():
            for existente in list(self.env):
                if existente.casefold() == nombre.casefold():
                    del self.env[existente]
            self.env[nombre] = valor

    def tearDown(self):
        self.temporal.cleanup()

    def crear(self, ruta):
        ruta.parent.mkdir(parents=True, exist_ok=True)
        if ruta.suffix.lower() == ".exe":
            shutil.copy2(Path(os.environ["COMSPEC"]), ruta)
        else:
            ruta.touch()
        return ruta

    def ejecutar(self):
        return subprocess.run(
            [os.environ["COMSPEC"], "/d", "/c", str(self.raiz / "inicio_laptop.bat")],
            cwd=self.raiz,
            env=self.env,
            capture_output=True,
            text=True,
            timeout=15,
            check=False,
        )

    def log(self):
        return (self.raiz / "arranque.log").read_text(encoding="cp1252")

    def comprobar(self, etiqueta):
        resultado = self.ejecutar()
        self.assertEqual(resultado.returncode, 0, resultado.stderr + "\n" + self.log())
        self.assertIn(f"{etiqueta} SI", self.log())

    def test_venv_pythonw(self):
        self.crear(self.raiz / ".venv/Scripts/pythonw.exe")
        self.comprobar("venv_pythonw")

    def test_solo_venv_python(self):
        self.crear(self.raiz / ".venv/Scripts/python.exe")
        self.comprobar("venv_python")

    def test_path_pythonw(self):
        binario = self.crear(self.raiz / "bin/pythonw.exe")
        self.env["PATH"] = f"{binario.parent};{self.env['PATH']}"
        self.comprobar("PATH_pythonw")

    def test_path_python(self):
        binario = self.crear(self.raiz / "bin/python.exe")
        self.env["PATH"] = f"{binario.parent};{self.env['PATH']}"
        self.comprobar("PATH_python")

    def test_path_pyw(self):
        binario = self.crear(self.raiz / "bin/pyw.exe")
        self.env["PATH"] = f"{binario.parent};{self.env['PATH']}"
        self.comprobar("PATH_pyw")
        self.assertIn("-3", self.log())

    def test_path_py(self):
        binario = self.crear(self.raiz / "bin/py.exe")
        self.env["PATH"] = f"{binario.parent};{self.env['PATH']}"
        self.comprobar("PATH_py")
        self.assertIn("-3", self.log())

    def test_launcher_oficial_fuera_de_path(self):
        self.crear(Path(self.env["REGISTRO_LAPTOP_LOCALAPPDATA"]) / "Programs/Python/Launcher/py.exe")
        self.comprobar("launcher_local_py")
        self.assertIn("-3", self.log())

    def test_local_app_data(self):
        self.crear(Path(self.env["LocalAppData"]) / "Programs/Python/Python313/pythonw.exe")
        self.comprobar("localappdata_pythonw")

    def test_program_files(self):
        self.crear(Path(self.env["ProgramFiles"]) / "Python313/python.exe")
        self.comprobar("programfiles_python")

    def test_program_files_x86(self):
        self.crear(Path(self.env["REGISTRO_LAPTOP_PROGRAMFILES_X86"]) / "Python311/pythonw.exe")
        self.comprobar("programfilesx86_pythonw")

    def test_aparece_despues_de_un_reintento(self):
        ruta = self.raiz / ".venv/Scripts/python.exe"

        def crear_despues():
            time.sleep(0.3)
            self.crear(ruta)

        hilo = threading.Thread(target=crear_despues)
        hilo.start()
        resultado = self.ejecutar()
        hilo.join()
        self.assertEqual(resultado.returncode, 0)
        self.assertIn("PYTHON_REINTENTO intento=1", self.log())
        self.assertIn("venv_python SI", self.log())

    def test_python_inexistente_deja_error_claro(self):
        resultado = self.ejecutar()
        self.assertEqual(resultado.returncode, 1)
        log = self.log()
        self.assertIn("ERROR: Python no encontrado tras 3 intentos", log)
        self.assertEqual(log.count("venv_pythonw NO"), 3)


if __name__ == "__main__":
    unittest.main()
