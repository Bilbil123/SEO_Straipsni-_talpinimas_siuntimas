from setuptools import setup

APP = ['email_sender_app.py']
DATA_FILES = ['email_list.txt']  # jei reikia
OPTIONS = {
    'argv_emulation': False,
    'excludes': ['wheel'],  # ← būtina, kad pagaliau atsikratytume šitos nesąmonės
}

setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)
