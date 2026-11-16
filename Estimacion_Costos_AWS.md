# Estimación de costos AWS — MediConnect (Sección 7)

Autora: Marcela Arévalo

**Enlace de la estimación (AWS Pricing Calculator, región us-east-1):**
https://calculator.aws/#/estimate?id=bc30fded3d33d4d662b6d66ff9c0b72faf09d0e7

## Totales (USD/mes)

| Concepto | USD/mes |
|---|---|
| Total de la calculadora (todos los servicios cargados) | 1,544.53 |
| SMS (800 mensajes x 0.337, cálculo manual) | 269.60 |
| Transferencia de datos de salida (manual) | 9.00 |
| **Total estimado** | **1,823.13** |
| Reserva del 10 % | 182.31 |
| **Presupuesto recomendado** | **2,005.44** |

## Partidas calculadas a mano

- **SMS:** la calculadora no tiene campo de SMS para El Salvador (SNS y Pinpoint cobran por país y la calculadora lo deriva a una herramienta aparte), por eso se calculó con la tarifa publicada: 800 x 0.337 = 269.60.
- **Transferencia de datos de salida:** cifra manual de 9.00.
- **EventBridge:** costo 0, no se agregó a la calculadora.

## Diferencias contra la tabla 7.1 actual del reporte

| Servicio | Calculadora | Reporte |
|---|---|---|
| RDS primaria | 277.42 | 332.50 |
| RDS réplica | 133.96 | 161.00 |
| S3 | 47.85 | 55.25 |
| CloudWatch | 42.23 | 31.50 |
| ALB / Fargate / Secrets Manager | 39.79 / 178.36 / 2.90 | 39.78 / 178.34 / 2.70 |
| **Total** | **1,823.13** | **1,901.70** |
| **Presupuesto recomendado** | **2,005.44** | **2,091.87** |

Se proponen los valores de la calculadora como oficiales; Hugo actualiza la tabla 7.1, el Anexo A y el enlace del reporte.
