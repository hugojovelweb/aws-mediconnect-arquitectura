#!/usr/bin/env python3
"""
MediConnect - Diagrama AWS a escala de página (PNG legible en Word/PDF).
Usa los íconos AWS incluidos en el paquete 'diagrams' (pip install diagrams) y matplotlib.
Uso: python3 generar_diagrama_pagina.py  ->  MediConnect_Diagrama_Arquitectura_AWS.png
"""
import os, diagrams
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, Circle
from matplotlib.offsetbox import OffsetImage, AnnotationBbox
from PIL import Image

RES = os.path.join(os.path.dirname(os.path.dirname(diagrams.__file__)), "resources", "aws")
def ic(path): return Image.open(os.path.join(RES, path)).convert("RGBA")

W, H = 10.0, 6.3
fig = plt.figure(figsize=(W, H), dpi=220)
ax = fig.add_axes([0, 0, 1, 1]); ax.set_xlim(0, W); ax.set_ylim(0, H); ax.axis("off")

def box(x, y, w, h, label, color, fill="white", ls="-", lw=1.4, fs=6.6):
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0,rounding_size=0.06",
                                fc=fill, ec=color, lw=lw, ls=ls, zorder=1))
    ax.text(x + 0.07, y + h - 0.06, label, fontsize=fs, fontweight="bold", color=color, va="top", ha="left", zorder=5)

def icon(path, x, y, label, size=0.46, fs=6.0):
    im = ic(path)
    ab = AnnotationBbox(OffsetImage(im, zoom=size * 72 / im.size[0]), (x, y), frameon=False, zorder=6)
    ax.add_artist(ab)
    ax.text(x, y - size / 2 - 0.04, label, fontsize=fs, ha="center", va="top", color="#232F3E", zorder=6, linespacing=1.05)

def arrow(p, q, color="#232F3E", lw=1.4, ls="-", rad=0.0, z=3):
    ax.annotate("", xy=q, xytext=p, zorder=z,
                arrowprops=dict(arrowstyle="-|>", color=color, lw=lw, ls=ls, shrinkA=0, shrinkB=0,
                                connectionstyle=f"arc3,rad={rad}", mutation_scale=8))

def poly(points, color="#232F3E", lw=1.4, ls="-", z=3):
    xs, ys = zip(*points[:-1]); ax.plot(xs, ys, color=color, lw=lw, ls=ls, zorder=z, solid_capstyle="round")
    arrow(points[-2], points[-1], color, lw, ls, z=z)

def badge(n, x, y):
    ax.add_patch(Circle((x, y), 0.105, fc="#FF9900", ec="white", lw=0.8, zorder=9))
    ax.text(x, y - 0.003, str(n), fontsize=6.4, fontweight="bold", color="white", ha="center", va="center", zorder=10)

def lab(x, y, s, color="#232F3E", fs=5.6, ha="center"):
    ax.text(x, y, s, fontsize=fs, color=color, ha=ha, va="center", zorder=8,
            bbox=dict(fc="white", ec="none", pad=0.6, alpha=0.9))

# ---- Título
ax.text(0.15, 6.14, "MediConnect — Arquitectura AWS (us-east-1, 2 Zonas de Disponibilidad)", fontsize=10.5, fontweight="bold", color="#232F3E", va="center")
ax.text(9.85, 6.14, "Kodigo Full Stack Jr 36", fontsize=6.5, color="#5F6B7A", va="center", ha="right")
ax.add_patch(FancyBboxPatch((0.08, 0.08), 9.84, 5.82, boxstyle="round,pad=0,rounding_size=0.05", fc="none", ec="#232F3E", lw=1.6, zorder=0))
ax.text(0.16, 5.83, "AWS Cloud", fontsize=7, fontweight="bold", color="#232F3E", va="top")

# ---- Usuarios y borde global
icon("general/users.png", 0.55, 4.95, "Pacientes y médicos\n(El Salvador)", 0.5)
box(1.05, 2.55, 1.7, 3.15, "Borde global", "#8C4FFF", "#F6F1FF")
icon("network/route-53.png", 1.9, 5.0, "Route 53\nDNS + health checks")
icon("network/cloudfront.png", 1.9, 3.95, "CloudFront (CDN)\ncaché Next.js")
icon("security/waf.png", 1.45, 2.95, "AWS WAF\n+ Shield Std", 0.4)
icon("security/certificate-manager.png", 2.35, 2.95, "ACM\nTLS gratis", 0.4)

# ---- Seguridad y observabilidad
box(0.15, 0.2, 2.6, 2.2, "Seguridad y observabilidad", "#DD344C", "#FFF4F5")
icon("security/secrets-manager.png", 0.6, 1.7, "Secrets\nManager", 0.38, 5.6)
icon("security/key-management-service.png", 1.45, 1.7, "KMS\n(claves CMK)", 0.38, 5.6)
icon("security/identity-and-access-management-iam.png", 2.3, 1.7, "IAM roles\nmín. privilegio", 0.38, 5.6)
icon("management/cloudwatch.png", 1.0, 0.8, "CloudWatch\nlogs + alarmas", 0.38, 5.6)
icon("management/cloudtrail.png", 1.95, 0.8, "CloudTrail\nauditoría", 0.38, 5.6)

# ---- Región / VPC
box(2.95, 0.2, 4.55, 5.55, "VPC  mediconnect-prod   10.0.0.0/16   ·   Región us-east-1", "#8C4FFF", "#FBFAFF", lw=1.8)
icon("network/internet-gateway.png", 3.15, 3.95, "Internet\nGateway", 0.42, 5.6)

box(3.45, 2.0, 1.25, 3.4, "Subnets públicas\n10.0.0.0/24 (AZ-a)\n10.0.1.0/24 (AZ-b)", "#7AA116", "#F5FAEA", fs=5.6)
icon("network/elb-application-load-balancer.png", 4.08, 3.95, "ALB\nMulti-AZ", 0.46)
icon("network/nat-gateway.png", 4.08, 2.75, "NAT Gateway\nx2 (una por AZ)", 0.42, 5.6)

box(4.9, 1.2, 1.55, 4.2, "Subnets privadas app\n10.0.10.0/24 · 10.0.11.0/24\nECS Fargate + Auto Scaling", "#ED7100", "#FFF7EE", fs=5.4)
icon("compute/fargate.png", 5.68, 4.5, "Portal Paciente\n(Next.js)", 0.42, 5.6)
icon("compute/fargate.png", 5.68, 3.6, "API Laravel\n(REST)", 0.42, 5.6)
icon("compute/fargate.png", 5.68, 2.7, "Dashboard Médico\n(Next.js)", 0.42, 5.6)
icon("compute/fargate.png", 5.68, 1.95, "Workers\nreportes + notificaciones", 0.42, 5.4)

box(6.65, 0.3, 0.77, 5.1, "Subnets\nprivadas datos\n10.0.20.0/24\n10.0.21.0/24", "#C925D1", "#FDF3FE", fs=5.0)
icon("database/rds-mysql-instance.png", 7.03, 4.3, "RDS MySQL 8.4\nMulti-AZ", 0.4, 5.4)
icon("database/rds-mysql-instance.png", 7.03, 3.3, "Réplica de\nlectura", 0.4, 5.4)
icon("database/elasticache-for-redis.png", 7.03, 2.3, "ElastiCache\nRedis", 0.4, 5.4)
icon("network/endpoint.png", 7.03, 1.3, "VPC Endpoint\nS3 (gateway)", 0.38, 5.2)

# ---- Servicios administrados
box(7.7, 0.3, 2.15, 5.45, "Servicios administrados", "#E7157B", "#FFF3F8")
icon("storage/simple-storage-service-s3.png", 8.2, 4.95, "S3 archivos médicos\nSSE-KMS · versionado\n+ CRR a us-east-2", 0.42, 5.2)
icon("storage/simple-storage-service-s3.png", 9.4, 4.95, "S3 reportes\nsemanales", 0.42, 5.6)
icon("integration/simple-queue-service-sqs.png", 8.2, 3.85, "SQS + DLQ\nreportes / notif.", 0.42, 5.6)
icon("integration/eventbridge.png", 9.4, 3.85, "EventBridge\nScheduler", 0.42, 5.6)
icon("engagement/simple-email-service-ses.png", 8.2, 2.75, "SES\ncorreo", 0.42, 5.6)
icon("integration/simple-notification-service-sns.png", 9.4, 2.75, "SNS\nSMS", 0.42, 5.6)
icon("general/multimedia.png", 8.2, 1.65, "Amazon Chime SDK\nvideoconsulta", 0.42, 5.4)
icon("compute/ec2-container-registry.png", 9.4, 1.65, "ECR\nimágenes", 0.42, 5.6)

# ---- Flechas de flujo (numeradas como en la tabla 4.1)
arrow((0.85, 4.95), (1.6, 5.0)); badge(1, 1.15, 5.12)
arrow((1.9, 4.72), (1.9, 4.25))
ax.annotate("", xy=(1.9, 4.25), xytext=(1.9, 3.35), arrowprops=dict(arrowstyle="-", color="#DD344C", lw=1.0, ls=":"), zorder=2)
poly([(2.2, 3.95), (2.95, 3.95)], "#232F3E", 1.8); badge(2, 2.62, 4.1)
poly([(3.36, 3.95), (3.83, 3.95)], "#232F3E", 1.8)
for y in (4.5, 3.6, 2.7):
    arrow((4.33, 3.95), (5.42, y), "#ED7100", 1.3)
badge(3, 4.9, 4.28)
arrow((5.68, 4.25), (5.68, 3.88), "#ED7100", 1.1); badge(4, 5.9, 4.07)
arrow((5.68, 3.05), (5.68, 3.33), "#ED7100", 1.1)
arrow((5.93, 3.6), (6.78, 4.3), "#C925D1", 1.6); badge(5, 6.38, 4.05)
lab(6.2, 4.42, "MySQL 3306", "#C925D1", 5.2)
arrow((7.03, 4.02), (7.03, 3.58), "#C925D1", 1.0, "--")
arrow((5.93, 2.05), (6.78, 3.2), "#C925D1", 1.0, "--")
lab(6.28, 2.75, "lee réplica", "#C925D1", 5.0)
arrow((5.93, 1.85), (6.84, 1.4), "#7AA116", 1.2)
lab(6.35, 1.5, "PDF/CSV", "#7AA116", 5.0)
poly([(5.93, 3.5), (6.5, 3.5), (6.5, 1.5), (6.84, 1.32)], "#7AA116", 1.3)
poly([(7.25, 1.3), (7.55, 1.3), (7.55, 4.95), (7.98, 4.95)], "#7AA116", 1.3); badge(6, 7.55, 2.9)
lab(7.55, 2.2, "URL\nprefirmada", "#7AA116", 5.0)
poly([(5.93, 3.8), (7.98, 3.85)], "#E7157B", 1.3); badge(7, 7.55, 3.98)
arrow((9.18, 3.85), (8.43, 3.85), "#E7157B", 1.0)
ax.text(8.77, 0.62, "Workers consumen SQS,\nleen de la réplica y envían\npor SES / SNS / S3", fontsize=5.4, color="#E7157B", ha="center", va="center", zorder=8)
plt.savefig("MediConnect_Diagrama_Arquitectura_AWS.png", dpi=220, facecolor="white")
print("OK")
