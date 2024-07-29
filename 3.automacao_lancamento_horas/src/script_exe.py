from cx_Freeze import setup, Executable

# Define os detalhes do executável
build_exe_options = {
    "packages": ["pandas", "selenium", "ttkbootstrap", "webdriver_manager"]
}


setup(
    name="Lançamento de Horas de Treinamento",
    version="1.01",
    description="Plataforma para lançamento de horas de treinamento das equipes jurídicas.",
    options={"build_exe": build_exe_options},
    executables=[Executable("app.py")]
)
