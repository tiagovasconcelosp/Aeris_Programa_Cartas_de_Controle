from PyInstaller.utils.hooks import copy_metadata, collect_submodules, collect_data_files

# Metadados (já resolviam o seu erro anterior de PackageNotFoundError)
datas = copy_metadata('streamlit')

# Submódulos dinâmicos do Streamlit
hiddenimports = collect_submodules('streamlit')

# >>> A parte que faz falta agora: arquivos estáticos (HTML/CSS/JS) <<<
datas += collect_data_files('streamlit')