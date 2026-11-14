#!/usr/bin/env python3
"""
MediConnect - genera el diagrama EDITABLE para draw.io (diagrams.net)
con los iconos oficiales AWS (libreria mxgraph.aws4 incluida en draw.io).
Uso:  python3 generar_drawio.py   ->  MediConnect_Arquitectura_AWS.drawio
Abrir en https://app.diagrams.net o en draw.io Desktop.
"""
from xml.sax.saxutils import escape

cells = []
_id = [1]
def nid():
    _id[0] += 1
    return f"c{_id[0]}"

COL = {"net": "#8C4FFF", "cmp": "#ED7100", "db": "#C925D1", "sto": "#7AA116",
       "app": "#E7157B", "sec": "#DD344C", "mgmt": "#E7157B", "gen": "#232F3E"}

def group(label, x, y, w, h, stroke, fill="none", dashed=0, icon=None):
    i = nid()
    st = (f"rounded=1;arcSize=2;whiteSpace=wrap;html=1;verticalAlign=top;align=left;spacingLeft=10;"
          f"fontStyle=1;fontSize=12;strokeColor={stroke};fillColor={fill};dashed={dashed};strokeWidth=2;fontColor={stroke};")
    cells.append(f'<mxCell id="{i}" value="{escape(label)}" style="{st}" vertex="1" parent="1">'
                 f'<mxGeometry x="{x}" y="{y}" width="{w}" height="{h}" as="geometry"/></mxCell>')
    return i

def icon(label, res, cat, x, y, size=56):
    i = nid()
    st = (f"sketch=0;outlineConnect=0;fontColor=#232F3E;fillColor={COL[cat]};strokeColor=#ffffff;dashed=0;"
          f"verticalLabelPosition=bottom;verticalAlign=top;align=center;html=1;fontSize=11;fontStyle=0;aspect=fixed;"
          f"shape=mxgraph.aws4.resourceIcon;resIcon=mxgraph.aws4.{res};")
    cells.append(f'<mxCell id="{i}" value="{escape(label)}" style="{st}" vertex="1" parent="1">'
                 f'<mxGeometry x="{x}" y="{y}" width="{size}" height="{size}" as="geometry"/></mxCell>')
    return i

def edge(a, b, label="", color="#232F3E", dashed=0, width=2):
    i = nid()
    st = (f"edgeStyle=orthogonalEdgeStyle;rounded=1;html=1;endArrow=block;endFill=1;strokeColor={color};"
          f"strokeWidth={width};dashed={dashed};fontSize=10;fontStyle=1;labelBackgroundColor=#ffffff;")
    cells.append(f'<mxCell id="{i}" value="{escape(label)}" style="{st}" edge="1" parent="1" source="{a}" target="{b}">'
                 f'<mxGeometry relative="1" as="geometry"/></mxCell>')

def text(label, x, y, w, h, size=11, bold=0):
    i = nid()
    cells.append(f'<mxCell id="{i}" value="{escape(label)}" style="text;html=1;align=left;verticalAlign=top;whiteSpace=wrap;fontSize={size};fontStyle={bold};" vertex="1" parent="1">'
                 f'<mxGeometry x="{x}" y="{y}" width="{w}" height="{h}" as="geometry"/></mxCell>')

text("MediConnect - Arquitectura AWS (us-east-1, 2 AZ)  |  Kodigo Full Stack Jr 36", 20, 10, 900, 24, 16, 1)

# Usuarios + borde
users = icon("Pacientes y médicos\n(El Salvador)", "users", "gen", 30, 330)
group("Borde global", 130, 120, 170, 520, COL["net"], "#F6F1FF")
r53 = icon("Route 53", "route_53", "net", 188, 160)
waf = icon("AWS WAF + Shield", "waf", "sec", 188, 290)
cf = icon("CloudFront (CDN)", "cloudfront", "net", 188, 420)
acm = icon("ACM (TLS)", "certificate_manager", "sec", 188, 550)

# Region / VPC
group("Región us-east-1 (N. Virginia)", 320, 60, 1260, 860, "#00A4A6", "none", 1)
group("VPC mediconnect-prod 10.0.0.0/16", 340, 100, 800, 800, COL["net"], "#FBFAFF")
igw = icon("Internet Gateway", "internet_gateway", "net", 370, 380)

group("Subnets PÚBLICAS 10.0.0.0/24 (AZ-a) · 10.0.1.0/24 (AZ-b)", 470, 130, 220, 640, COL["sto"], "#F5FAEA")
alb = icon("Application Load Balancer\n(Multi-AZ)", "application_load_balancer", "net", 550, 330)
nat = icon("NAT Gateway x2", "nat_gateway", "net", 550, 600)

group("Subnets PRIVADAS app 10.0.10.0/24 · 10.0.11.0/24 — ECS Fargate (Auto Scaling + escalado programado lunes)", 710, 130, 410, 440, COL["cmp"], "#FFF7EE")
portal = icon("Portal Paciente\n(Next.js)", "fargate", "cmp", 740, 200)
dash = icon("Dashboard Médico\n(Next.js)", "fargate", "cmp", 740, 340)
api = icon("API Laravel\n(REST)", "fargate", "cmp", 900, 270)
worker = icon("Workers\nreportes + notif.", "fargate", "cmp", 1030, 440)

group("Subnets PRIVADAS de datos 10.0.20.0/24 · 10.0.21.0/24 (sin ruta a Internet)", 710, 600, 410, 280, COL["db"], "#FDF3FE")
rds = icon("RDS MySQL 8.4\nMulti-AZ", "rds", "db", 740, 660)
rep = icon("Réplica de lectura\n(reportes)", "rds", "db", 880, 660)
redis = icon("ElastiCache\nRedis", "elasticache", "db", 1020, 660)

# Servicios regionales
group("Servicios administrados (fuera de la VPC)", 1160, 100, 400, 520, COL["app"], "#FFF3F8")
s3 = icon("S3 archivos médicos\n(SSE-KMS, versionado, CRR)", "s3", "sto", 1190, 140)
s3r = icon("S3 reportes", "s3", "sto", 1330, 140)
sqs = icon("SQS + DLQ", "sqs", "app", 1190, 270)
evb = icon("EventBridge\nScheduler", "eventbridge", "app", 1330, 270)
ses = icon("SES (correo)", "ses", "app", 1190, 400)
sns = icon("SNS (SMS)", "sns", "app", 1330, 400)
chime = icon("Amazon Chime SDK\n(videoconsulta)", "chime", "app", 1460, 270)
ecr = icon("ECR", "elastic_container_registry", "cmp", 1460, 140)

group("Seguridad y observabilidad", 1160, 640, 400, 240, COL["sec"], "#FFF4F5")
sm = icon("Secrets Manager", "secrets_manager", "sec", 1190, 700)
kms = icon("KMS", "key_management_service", "sec", 1290, 700)
iam = icon("IAM roles", "identity_and_access_management", "sec", 1390, 700)
cw = icon("CloudWatch", "cloudwatch", "mgmt", 1240, 800)
ct = icon("CloudTrail", "cloudtrail", "mgmt", 1360, 800)

edge(users, r53, "1 HTTPS")
edge(r53, cf)
edge(waf, cf, "", COL["sec"], 1, 1)
edge(acm, cf, "", COL["sec"], 1, 1)
edge(cf, igw, "2 origen HTTPS")
edge(igw, alb)
edge(alb, portal, "3", COL["cmp"])
edge(alb, dash, "", COL["cmp"])
edge(alb, api, "3 API", COL["cmp"])
edge(portal, api, "4 REST", COL["cmp"])
edge(dash, api, "", COL["cmp"])
edge(api, rds, "5 MySQL 3306 (SG→SG)", COL["db"])
edge(api, redis, "", COL["db"])
edge(rds, rep, "replicación", COL["db"], 1)
edge(worker, rep, "lee reportes", COL["db"])
edge(api, sqs, "encola", COL["app"])
edge(evb, sqs, "programa", COL["app"], 1)
edge(sqs, worker, "consume", COL["app"])
edge(worker, ses, "", COL["app"])
edge(worker, sns, "", COL["app"])
edge(worker, s3r, "PDF/CSV", COL["sto"])
edge(api, s3, "URL prefirmada", COL["sto"])
edge(api, chime, "sesión", COL["app"], 1)
edge(sm, api, "secrets al arrancar", COL["sec"], 1, 1)
edge(api, cw, "logs/métricas", "#999999", 1, 1)

xml = ('<mxfile host="app.diagrams.net"><diagram name="MediConnect AWS" id="mediconnect-aws">'
       '<mxGraphModel dx="1600" dy="900" grid="1" gridSize="10" guides="1" tooltips="1" connect="1" arrows="1" fold="1" page="1" pageScale="1" pageWidth="1700" pageHeight="960" math="0" shadow="0">'
       '<root><mxCell id="0"/><mxCell id="1" parent="0"/>' + "".join(cells) + '</root></mxGraphModel></diagram></mxfile>')
open("MediConnect_Arquitectura_AWS.drawio", "w", encoding="utf-8").write(xml)
print("OK -> MediConnect_Arquitectura_AWS.drawio")
