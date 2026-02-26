import os
import sys
import subprocess
import venv
from pathlib import Path


def run_command(command):
    result = subprocess.run(command, shell=True)
    if result.returncode != 0:
        print(f"Error while running: {command}")
        sys.exit(1)


def main():
    project_root = Path(__file__).parent
    venv_path = project_root / ".venv"

    print("Creating virtual environment...")
    if not venv_path.exists():
        venv.create(venv_path, with_pip=True)
    else:
        print(".venv already exists.")

    # Determine correct python executable inside venv
    if os.name == "nt":
        python_executable = venv_path / "Scripts" / "python.exe"
    else:
        python_executable = venv_path / "bin" / "python"

    print("Upgrading pip...")
    run_command(f'"{python_executable}" -m pip install --upgrade pip')

    print("Installing requirements...")
    run_command(f'"{python_executable}" -m pip install -r requirements.txt')

    print("\nSetup completed successfully.")
    print("To activate the environment:")
    if os.name == "nt":
        print(r".\.venv\Scripts\activate")
    else:
        print("source .venv/bin/activate")


if __name__ == "__main__":
    main()
