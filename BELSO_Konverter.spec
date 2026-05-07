name: Build Windows EXE

on:
  push:
    branches: [ main ]
  workflow_dispatch:       # kézi indítás is lehetséges

jobs:
  build-windows:
    runs-on: windows-latest

    steps:
      - name: Kód letöltése
        uses: actions/checkout@v4

      - name: Python 3.11 beállítása
        uses: actions/setup-python@v5
        with:
          python-version: '3.11'

      - name: Függőségek telepítése
        run: |
          pip install pyinstaller openpyxl xlrd tkinterdnd2

      - name: EXE buildelése
        run: |
          pyinstaller BELSO_Konverter.spec

      - name: EXE feltöltése letölthető artifactként
        uses: actions/upload-artifact@v4
        with:
          name: BELSO_Konverter_Windows
          path: dist/BELSO_Konverter.exe
          retention-days: 30
