# Implementar soluciones escalables en la nube seleccionando los servicios de AWS adecuados según las necesidades del negocio. 

# Alumnos: 
Icxell Marcela Arévalo Mesa.
Hugo Ernesto Jovel Hernandez.

Kodigo Full Stack Jr 36 ·
© Hugo Jovel Web (https://hugojovelweb.tech/)


## Archivos entregados
| Archivo | Para qué sirve |
|---|---|
| MediConnect_Reporte_Tecnico_AWS.docx | Reporte editable (entregable) |
| MediConnect_Reporte_Tecnico_AWS.pdf | Misma versión en PDF |
| MediConnect_Diagrama_Arquitectura_AWS.png | Diagrama incrustado en el reporte |
| MediConnect_Arquitectura_AWS.drawio | Diagrama editable (draw.io, íconos oficiales aws4) |
| generar_reporte.js | Regenera el .docx (Node + docx) |
| generar_diagrama_pagina.py | Regenera el PNG del reporte (Python) |
| generar_drawio.py | Regenera el .drawio |

Antes de entregar: poner nombres de integrantes (portada) y pegar el enlace/captura de AWS Pricing Calculator (sección 7, Anexo A).

## A) Debian XFCE 13 (trixie)
```bash
sudo apt update
sudo apt install -y nodejs npm python3 python3-pip python3-venv graphviz libreoffice-writer fonts-liberation git
# draw.io Desktop (opcional; también sirve https://app.diagrams.net)
wget https://github.com/jgraph/drawio-desktop/releases/latest/download/drawio-amd64.deb -O drawio.deb
sudo apt install -y ./drawio.deb

mkdir -p ~/kodigo-full-stack-j36/Modulo-XX/Tarea-aws-mediconnect && cd ~/kodigo-full-stack-j36/Modulo-XX/Tarea-aws-mediconnect
# copiar aquí los archivos descargados, luego:
python3 -m venv .venv && source .venv/bin/activate
pip install matplotlib pillow diagrams
npm init -y && npm install docx

python3 generar_diagrama_pagina.py        # -> PNG
python3 generar_drawio.py                 # -> .drawio
node generar_reporte.js                   # -> .docx (necesita el PNG en la carpeta)
soffice --headless --convert-to pdf MediConnect_Reporte_Tecnico_AWS.docx   # -> PDF
```
Abrir: `libreoffice --writer MediConnect_Reporte_Tecnico_AWS.docx` y `drawio MediConnect_Arquitectura_AWS.drawio`.

Commit sugerido (Conventional Commits):
```bash
git add . && git commit -m "docs(aws): reporte técnico de arquitectura MediConnect" && git push
```
(Cambia `Modulo-XX` por el módulo real.)

## B) Windows 10/11 (PowerShell)
```powershell
winget install OpenJS.NodeJS.LTS
winget install Python.Python.3.12
winget install Graphviz.Graphviz        # reiniciar la terminal para que dot esté en PATH
winget install TheDocumentFoundation.LibreOffice
winget install JGraph.Draw
mkdir C:\kodigo\Tarea-aws-mediconnect ; cd C:\kodigo\Tarea-aws-mediconnect
# copiar aquí los archivos, luego:
python -m venv .venv ; .\.venv\Scripts\Activate.ps1
pip install matplotlib pillow diagrams
npm init -y ; npm install docx
python generar_diagrama_pagina.py ; python generar_drawio.py ; node generar_reporte.js
& "C:\Program Files\LibreOffice\program\soffice.exe" --headless --convert-to pdf MediConnect_Reporte_Tecnico_AWS.docx
```
Si PowerShell bloquea el activate: `Set-ExecutionPolicy -Scope Process Bypass`.

## C) Editar sin código
- Word: abrir el .docx (Word o LibreOffice) y editar texto/tablas. El año del pie se actualiza solo.
- Diagrama: abrir el .drawio, ajustar y *File → Export as → PNG (200 %)*; reemplazar la imagen en Word.
- Si un ícono sale en blanco: *More Shapes → AWS 19* y arrastrar el correcto.

## D) AWS Pricing Calculator
1. https://calculator.aws → Create estimate → región US East (N. Virginia).
2. Ingresar los valores del Anexo A del reporte.
3. Share → copiar el enlace y pegarlo en la sección 7.
