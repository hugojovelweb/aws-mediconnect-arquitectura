# Implementar soluciones escalables en la nube seleccionando los servicios de AWS adecuados según las necesidades del negocio

> **Kodigo Full Stack Jr 36 · Actividad evaluada grupal · Caso MediConnect (telemedicina)**
> Cierre de entrega: **domingo 15 de noviembre de 2026, 23:59**

Repositorio de trabajo del **reporte técnico de arquitectura AWS** para MediConnect, una plataforma salvadoreña de telemedicina que debe migrar de un servidor físico único a la nube. Contiene el entregable (Word/PDF), el diagrama editable y los scripts que lo generan de forma reproducible.

| | |
|---|---|
| **Equipo** | Hugo Jovel (`@hugojovelweb`, responsable del repositorio) · `Icxell Marcela Arévalo Mesa` (`https://github.com/marcelaarevalom-jpg`, revisión) |
| **Región propuesta** | `us-east-1` (N. Virginia), 2 Zonas de Disponibilidad |
| **Estado** | En revisión — ver [Pendientes antes de entregar](#7-pendientes-antes-de-entregar) |

---

## 1. Contexto del caso

MediConnect opera hoy con **un solo servidor físico** que colapsa los lunes por la mañana (hasta 5× el tráfico diario). Componentes a migrar:

| Componente | Tecnología actual |
|---|---|
| Portal del paciente y dashboard del médico | Next.js |
| API principal | Laravel (REST) |
| Base de datos | MySQL (datos clínicos, recetas y pagos) |
| Archivos médicos | Disco local del servidor |
| Notificaciones | Correo electrónico y SMS |
| Reportes semanales | Proceso de 8–15 min que bloquea el servidor |

**Requerimientos obligatorios del negocio**

| ID | Requerimiento |
|---|---|
| R1 | Escalabilidad: hasta **10×** el tráfico normal los lunes sin degradación |
| R2 | Alta disponibilidad **≥ 99.9 %** |
| R3 | BD nunca accesible desde Internet; credenciales fuera del código |
| R4 | Tiempos de carga **< 2 s** en todo el país |
| R5 | Reportes desacoplados: no deben afectar a la API |
| R6 | Archivos médicos con durabilidad extrema |
| R7 | Costos predecibles (estimación mensual clara) |

## 2. Solución en una vista

```
Usuario → Route 53 → CloudFront (+ WAF, ACM) → Internet Gateway → ALB (Multi-AZ)
        → ECS Fargate (Portal · Dashboard · API Laravel) → RDS MySQL 8.4 Multi-AZ
                                   │                          ├─ Réplica de lectura (reportes)
                                   │                          └─ ElastiCache Redis (sesiones/caché)
                                   ├─ S3 archivos médicos (SSE-KMS, versionado, CRR) vía VPC Endpoint
                                   └─ SQS ← EventBridge Scheduler → Workers Fargate → SES · SNS · S3 reportes
Seguridad/observabilidad: IAM · Secrets Manager · KMS · CloudWatch · CloudTrail · Budgets
```

| Capa | Subredes (AZ-a / AZ-b) | Recursos | Salida a Internet |
|---|---|---|---|
| Pública | `10.0.0.0/24` / `10.0.1.0/24` | ALB, NAT Gateway | Directa (IGW) |
| Privada de aplicación | `10.0.10.0/24` / `10.0.11.0/24` | Tareas ECS Fargate | Solo egreso vía NAT; S3 por VPC Endpoint |
| Privada de datos | `10.0.20.0/24` / `10.0.21.0/24` | RDS, ElastiCache | Ninguna |

**Decisiones técnicas clave**

- **ECS + Fargate** en lugar de EC2/EKS: sin administrar servidores, escalado por tarea y menor complejidad para 4 servicios.
- **RDS MySQL 8.4 LTS Multi-AZ**: conserva el motor actual; se evita 8.0 por el fin de su soporte estándar (verificar en la documentación de RDS).
- **SQS + EventBridge Scheduler + workers Fargate** para reportes: Lambda queda descartada por su límite de 15 min.
- **Escalado programado los lunes** (05:30–12:00, UTC−6) + Application Auto Scaling hasta 20 tareas de API (10× la base).
- **Secrets Manager + roles IAM de tarea**: ningún secreto en el repositorio ni en la imagen.

El detalle completo, con tablas de justificación (16 servicios) y estimación de costos (~US$ 1,901.70/mes + 10 % de contingencia), está en el reporte.

## 3. Estructura del repositorio

```
aws-mediconnect-arquitectura/
├── docs/                                      # Entregables (generados)
│   ├── MediConnect_Reporte_Tecnico_AWS.docx   # Reporte editable
│   ├── MediConnect_Reporte_Tecnico_AWS.pdf    # Reporte en PDF
│   ├── MediConnect_Diagrama_Arquitectura_AWS.png
│   └── MediConnect_Arquitectura_AWS.drawio    # Diagrama editable (draw.io)
├── tools/                                     # Generadores reproducibles
│   ├── generar_reporte.js                     # Node + docx → .docx
│   ├── generar_diagrama_pagina.py             # matplotlib + diagrams → .png
│   └── generar_drawio.py                      # Python → .drawio
├── revision/                                  # Hallazgos de la compañera (texto)
├── GUIA_Debian_XFCE13_y_Windows.md            # Instalación y uso por sistema operativo
└── README.md
```

## 4. Reproducir el reporte

**Requisitos:** Node.js 20+, Python 3.10+, Graphviz, LibreOffice Writer (solo para PDF), fuente Liberation/Arial.

### Debian XFCE 13

```bash
sudo apt update
sudo apt install -y nodejs npm python3 python3-pip python3-venv graphviz libreoffice-writer fonts-liberation git

python3 -m venv .venv && source .venv/bin/activate
pip install matplotlib pillow diagrams
npm init -y && npm install docx
```

### Windows 10/11 (PowerShell)

```powershell
winget install OpenJS.NodeJS.LTS
winget install Python.Python.3.12
winget install Graphviz.Graphviz
winget install TheDocumentFoundation.LibreOffice
python -m venv .venv ; .\.venv\Scripts\Activate.ps1
pip install matplotlib pillow diagrams
npm init -y ; npm install docx
```

### Generación (desde la raíz del repo, mismo comando en ambos sistemas)

```bash
cd docs
python3 ../tools/generar_diagrama_pagina.py    # escribe el PNG en docs/
python3 ../tools/generar_drawio.py             # escribe el .drawio en docs/
node ../tools/generar_reporte.js               # lee el PNG y escribe el .docx en docs/
soffice --headless --convert-to pdf MediConnect_Reporte_Tecnico_AWS.docx
```

En Windows usa `python` en lugar de `python3` y la ruta completa `& "C:\Program Files\LibreOffice\program\soffice.exe"` para el PDF.

**Verificación rápida (Debian):**

```bash
pdfinfo MediConnect_Reporte_Tecnico_AWS.pdf | grep Pages
pdftotext MediConnect_Reporte_Tecnico_AWS.pdf - | grep -n -E "\[Nombre|ENLACE DE LA ESTIMACIÓN"   # debe salir vacío al entregar
xmllint --noout MediConnect_Arquitectura_AWS.drawio && echo "drawio válido"
```

## 5. Flujo de trabajo Git

Se trabaja como en un entorno de producción: **`main` está protegida y nada entra sin Pull Request**.

| Rol | Persona | Responsabilidad |
|---|---|---|
| Mantenedor | Hugo | Genera los entregables, revisa y fusiona los PR, hace la entrega final |
| Revisora | Compañera | Verifica precios y requisitos, crea la estimación en AWS Pricing Calculator, reporta hallazgos |

**Reglas**

1. Nadie hace commit directo a `main`. Cada cambio va en una rama y entra por PR.
2. Los archivos de `docs/` (`.docx`, `.pdf`, `.png`, `.drawio`) **solo los modifica el mantenedor**: son binarios o XML grandes y no se pueden fusionar sin perder trabajo.
3. La revisora trabaja únicamente en `revision/REVISION_COMPANERA.md`.
4. Un PR = un cambio con un propósito. Se fusiona con *Merge pull request* y se borra la rama.

**Nombres de rama**

| Prefijo | Uso | Ejemplo |
|---|---|---|
| `docs/` | Documentación y entregables | `docs/reporte-v1` |
| `fix/` | Corrección de errores | `fix/precio-rds` |
| `review/` | Hallazgos de revisión | `review/companera-r1` |

**Mensajes de commit ([Conventional Commits](https://www.conventionalcommits.org/es/v1.0.0/))**

```
docs(aws): agrega diagrama de arquitectura
fix(costos): corrige precio de la instancia RDS
chore(tools): actualiza dependencias de los generadores
```

**Ciclo de trabajo del mantenedor**

```bash
git switch main && git pull
git switch -c docs/mi-cambio
# ...cambios...
git add . && git commit -m "docs(aws): descripción del cambio"
git push -u origin docs/mi-cambio
# abrir Pull Request en GitHub → revisar «Files changed» → Merge → Delete branch
git switch main && git pull
```

**Protección de `main` (una sola vez):** *Settings → Rules → Rulesets → New branch ruleset* → *Include default branch* → activar *Restrict deletions*, *Block force pushes* y *Require a pull request before merging* (aprobaciones requeridas: 0, porque el mantenedor es el único que fusiona).

### Guía para la revisora (sin terminal)

1. Abrir `revision/REVISION_COMPANERA.md` en GitHub y pulsar el lápiz (*Edit this file*).
2. Escribir los hallazgos usando la plantilla del archivo.
3. Pulsar *Commit changes…* → elegir **Create a new branch for this commit and start a pull request** → nombre `review/companera-r1` → *Propose changes* → *Create pull request*.
4. No editar nada fuera de ese archivo. Ante cualquier duda, comentar en el PR.

## 6. Decisiones y supuestos a validar

| Tema | Supuesto / decisión | Acción |
|---|---|---|
| Volumen | 50 000 usuarios activos/mes; 8 000 videoconsultas × 20 min; 800 SMS/mes; 1 TB de archivos | Ajustar con datos reales del caso si el instructor los entrega |
| Precios | Lista de `us-east-1`, USD, bajo demanda, 730 h/mes | Contrastar con **AWS Pricing Calculator** (Anexo A del reporte) |
| RDS | `db.m6g.large` Multi-AZ ≈ US$ 277/mes (varias fuentes difieren) | Confirmar tarifa exacta en la página oficial de RDS |
| SMS | ≈ US$ 0.337 por mensaje a El Salvador | Confirmar en la página oficial de SNS |
| Videoconsulta | Amazon Chime SDK, US$ 0.0017 por minuto-asistente | Confirmar en la página oficial; es el mayor costo variable |
| Equipo | La actividad pide 3 a 5 integrantes y somos 2 | Confirmar con coordinación del bootcamp |

## 7. Pendientes antes de entregar

- [ ] Nombres completos de los integrantes en la portada del reporte
- [ ] Estimación en **AWS Pricing Calculator** con su enlace/captura pegado en la sección 7
- [ ] Verificación de precios por la revisora (`revision/REVISION_COMPANERA.md`)
- [ ] Exportar el diagrama desde draw.io (PNG 200 %) y reemplazarlo en el Word si se editó
- [ ] Regenerar el PDF desde el `.docx` final
- [ ] Revisión visual del `.docx` en Word/LibreOffice (tablas, pie de página, numeración)
- [ ] Probar la reproducción en Windows
- [ ] Entrega en el aula virtual antes del **15-nov-2026 23:59**

## 8. Seguridad del repositorio

- El repositorio es **privado** durante el desarrollo.
- No se versionan secretos. Las políticas IAM y definiciones de tarea del reporte usan `ACCOUNT_ID` y `KEY_ID_*` como marcadores de ejemplo.
- `.gitignore` excluye `node_modules/`, `.venv/`, `__pycache__/`, `*.tmp` y los archivos de bloqueo de LibreOffice (`.~lock.*`).

## 9. Referencias

- [AWS Pricing Calculator](https://calculator.aws)
- [AWS Well-Architected Framework](https://aws.amazon.com/architecture/well-architected/)
- [Iconos oficiales de arquitectura AWS](https://aws.amazon.com/architecture/icons/)
- [draw.io / diagrams.net](https://app.diagrams.net)
- [Conventional Commits](https://www.conventionalcommits.org/es/v1.0.0/)

---

<sub>Kodigo Full Stack Jr 36 · AWS · Next.js 16 · Laravel 12 · MySQL 8.4</sub><br>
<sub>© 2026 <a href="https://hugojovelweb.tech/" target="_blank" rel="noopener noreferrer"><b>Hugo Jovel Web</b></a>. Todos los derechos reservados.</sub>
