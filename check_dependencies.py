import sys
import subprocess

# Lista de dependencias requeridas
REQUIRED_PACKAGES = [
    ("pandas", "pandas"),
    ("openpyxl", "openpyxl"),
    ("pymupdf", "PyMuPDF")
]

def check_and_install():
    missing = []
    for module_name, pip_name in REQUIRED_PACKAGES:
        try:
            __import__(module_name)
        except ImportError:
            missing.append((module_name, pip_name))
    if not missing:
        return True
    print("\nFaltan las siguientes dependencias para ejecutar el programa:")
    for module_name, pip_name in missing:
        print(f"- {pip_name}")
    resp = input("¿Deseas que intente instalarlas automáticamente? (s/n): ").strip().lower()
    if resp == "s":
        for module_name, pip_name in missing:
            print(f"Instalando {pip_name}...")
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", pip_name])
            except FileNotFoundError:
                print(f"No se pudo encontrar 'pip'. Por favor, instala pip antes de continuar.\nGuía oficial: https://pip.pypa.io/en/stable/installation/")
                sys.exit(1)
            except Exception as e:
                print(f"No se pudo instalar {pip_name}: {e}")
        print("\nIntenta ejecutar el programa nuevamente.")
        sys.exit(0)
    else:
        print("Por favor, instala las dependencias manualmente con:")
        for _, pip_name in missing:
            print(f"pip install {pip_name}")
        sys.exit(1)

if __name__ == "__main__":
    check_and_install()
