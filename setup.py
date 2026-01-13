from setuptools import setup, find_packages
import os

# Leer el README para usarlo como descripción larga
def read_readme():
    readme_path = os.path.join(os.path.dirname(__file__), "README.md")
    if os.path.exists(readme_path):
        with open(readme_path, "r", encoding="utf-8") as f:
            return f.read()
    return ""

# Leer la versión desde el código o definirla aquí
__version__ = "1.1.1"

setup(
    name="dataframexl",
    version=__version__,
    author="Tu Nombre",  # Cambiar por tu nombre o el del autor
    author_email="tu.email@ejemplo.com",  # Cambiar por tu email
    description="Extensión de pandas.DataFrame para trabajar con Excel y estilos de forma integrada",
    long_description=read_readme(),
    long_description_content_type="text/markdown",
    url="https://github.com/tu-usuario/DataFrameStyle",  # Cambiar por la URL de tu repositorio
    py_modules=["DFXL"],  # Módulo principal
    include_package_data=True,
    install_requires=[
        "pandas>=1.3.0",
        "openpyxl>=3.0.0",
        "numpy>=1.20.0",
    ],
    python_requires=">=3.7",
    classifiers=[
        "Development Status :: 3 - Alpha",
        "Intended Audience :: Developers",
        "Intended Audience :: Science/Research",
        "Topic :: Software Development :: Libraries :: Python Modules",
        "Topic :: Scientific/Engineering",
        "License :: OSI Approved :: MIT License", 
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Operating System :: OS Independent",
    ],
    keywords="pandas excel openpyxl dataframe styling formatting",
    project_urls={
        "Bug Reports": "https://github.com/tu-usuario/DataFrameStyle/issues",
        "Source": "https://github.com/tu-usuario/DataFrameStyle",
        "Documentation": "https://github.com/tu-usuario/DataFrameStyle#readme",
    },
)

