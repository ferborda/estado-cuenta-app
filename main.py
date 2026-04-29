from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import HTMLResponse, FileResponse
import pandas as pd
import shutil
import matplotlib.pyplot as plt

from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
)
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas

app = FastAPI()

LOGO = "logo.jpg"
UPLOAD = "temp.xlsx"
PDF = "reporte.pdf"
GRAF = "graf.png"


# ===============================
# 📥 CARGAR EXCEL
# ===============================
def cargar_excel(path):
    df_raw = pd.read_excel(path, header=None)

    nit = str(df_raw.iloc[0, 1]).strip()

    df = df_raw.iloc[2:].copy()
    df = df.iloc[:, :12]

    df.columns = [
        "documento","tipo","fecha","vencimiento","dias","estado",
        "total","moneda","tasa","cobrado","retenido","pendiente"
    ]

    df["dias"] = pd.to_numeric(df["dias"], errors="coerce").fillna(0)
    df["saldo"] = pd.to_numeric(df["pendiente"], errors="coerce").fillna(0)

    df = df[df["saldo"] > 0]

    if df.empty:
        raise ValueError("No hay facturas con saldo pendiente")

    return df, nit


# ===============================
# 📊 CLASIFICACIÓN
# ===============================
def clasificar(df):

    buckets = {
        "Sin vencer": [],
        "1-30 días": [],
        "31-60 días": [],
        "61-90 días": [],
        "91-180 días": [],
        "181-360 días": [],
        "+1 año": []
    }

    for _, r in df.iterrows():

        saldo = float(r["saldo"])
        dias = int(r["dias"])

        fila = [
            str(r["documento"]),
            str(r["fecha"])[:10],
            str(r["vencimiento"])[:10],
            dias,
            f"${saldo:,.0f}"
        ]

        if dias <= 0:
            buckets["Sin vencer"].append((fila, saldo))
        elif dias <= 30:
            buckets["1-30 días"].append((fila, saldo))
        elif dias <= 60:
            buckets["31-60 días"].append((fila, saldo))
        elif dias <= 90:
            buckets["61-90 días"].append((fila, saldo))
        elif dias <= 180:
            buckets["91-180 días"].append((fila, saldo))
        elif dias <= 360:
            buckets["181-360 días"].append((fila, saldo))
        else:
            buckets["+1 año"].append((fila, saldo))

    return buckets


# ===============================
# 📈 INDICADORES
# ===============================
def indicadores(buckets):

    total = 0
    vencido = 0

    for k, v in buckets.items():
        for _, s in v:
            total += s
            if k != "Sin vencer":
                vencido += s

    riesgo = (vencido / total * 100) if total else 0

    if riesgo < 20:
        txt = "Riesgo bajo: cartera saludable"
        color = colors.lightgreen
    elif riesgo < 40:
        txt = "Riesgo medio: monitoreo"
        color = colors.yellow
    elif riesgo < 60:
        txt = "Riesgo alto: gestionar cobranza"
        color = colors.orange
    else:
        txt = "Riesgo crítico: acción inmediata"
        color = colors.red

    return total, vencido, riesgo, txt, color


# ===============================
# 📊 GRÁFICO
# ===============================
def generar_grafico(buckets):

    labels, valores = [], []

    for k, v in buckets.items():
        s = sum(x for _, x in v)
        if s > 0:
            labels.append(k)
            valores.append(s / 1_000_000)

    plt.figure(figsize=(4,2.5))
    bars = plt.bar(labels, valores, color="#008037")

    plt.xticks(rotation=25, fontsize=7)
    plt.yticks(fontsize=7)

    plt.gca().yaxis.set_major_formatter(
        plt.FuncFormatter(lambda x, _: f"{x:.1f}M")
    )

    for bar in bars:
        y = bar.get_height()
        plt.text(bar.get_x()+bar.get_width()/2, y, f"{y:.1f}",
                 ha='center', va='bottom', fontsize=7)

    plt.tight_layout()
    plt.savefig(GRAF)
    plt.close()


# ===============================
# 📄 NUMERACIÓN
# ===============================
class NumCanvas(canvas.Canvas):
    def __init__(self,*a,**k):
        super().__init__(*a,**k)
        self.pages=[]

    def showPage(self):
        self.pages.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        total=len(self.pages)
        for p in self.pages:
            self.__dict__.update(p)
            self.drawRightString(A4[0]-2*cm,1.5*cm,f"Pág {self._pageNumber} de {total}")
            super().showPage()
        super().save()


# ===============================
# 📄 PDF
# ===============================
def generar_pdf(cliente, nit, buckets, total, vencido, riesgo, txt, color):

    doc = SimpleDocTemplate(PDF, pagesize=A4)
    styles = getSampleStyleSheet()
    elems = []

    # HEADER
    try:
        logo = Image(LOGO, width=6*cm, height=2.5*cm)
    except:
        logo = Paragraph("FLEXOPACK", styles["Title"])

    header = Table([[logo, Paragraph("<b>ESTADO DE CUENTA</b>", styles["Title"])]])
    elems.append(header)
    elems.append(Spacer(1,6))

    # KPI + DASHBOARD
    generar_grafico(buckets)

    kpi_texto = Paragraph(
        f"Cliente: {cliente}<br/>"
        f"NIT: {nit}<br/><br/>"
        f"<b>Total Cartera: ${total:,.0f}</b><br/>"
        f"<b>Total Vencido: ${vencido:,.0f}</b>",
        styles["Normal"]
    )

    tabla_kpi = Table(
        [[kpi_texto, Image(GRAF, width=6*cm, height=2.8*cm)]],
        colWidths=[9*cm,7*cm]
    )

    tabla_kpi.setStyle([
        ('VALIGN',(0,0),(-1,-1),'TOP')
    ])

    elems.append(tabla_kpi)
    elems.append(Spacer(1,6))

    # RIESGO
    riesgo_texto = Paragraph(
        f"<b>Riesgo: {riesgo:.2f}%</b><br/>{txt}",
        styles["Normal"]
    )

    riesgo_box = Table([[riesgo_texto]], colWidths=[16*cm])
    riesgo_box.setStyle([
        ('BOX',(0,0),(-1,-1),1,colors.grey),
        ('BACKGROUND',(0,0),(-1,-1),color),
        ('PADDING',(0,0),(-1,-1),6)
    ])

    elems.append(riesgo_box)
    elems.append(Spacer(1,8))

    # TABLAS (FLUIDAS, SIN CORTES FEOS)
    for k,v in buckets.items():
        if v:
            filas=[f for f,_ in v]

            t=Table(
                [["Número","Fecha","Venc","Días","Saldo"]]+filas,
                repeatRows=1
            )

            t.setStyle([
                ('BACKGROUND',(0,0),(-1,0),colors.lightgrey),
                ('GRID',(0,0),(-1,-1),0.25,colors.grey)
            ])

            elems.append(Paragraph(f"<b>{k}</b>", styles["Heading3"]))
            elems.append(t)
            elems.append(Spacer(1,8))

    doc.build(elems, canvasmaker=NumCanvas)


# ===============================
# 🌐 WEB
# ===============================
@app.get("/", response_class=HTMLResponse)
def home():
    return """
    <h2>Software Cartera</h2>
    <form action="/pdf" method="post" enctype="multipart/form-data">
        Cliente:<br><input name="cliente" required><br><br>
        Excel:<br><input type="file" name="file" required><br><br>
        <button>Generar PDF</button>
    </form>
    """


# ===============================
# 🚀 GENERAR
# ===============================
@app.post("/pdf")
def pdf(cliente: str = Form(...), file: UploadFile = File(...)):
    try:
        with open(UPLOAD, "wb") as b:
            shutil.copyfileobj(file.file, b)

        df, nit = cargar_excel(UPLOAD)
        buckets = clasificar(df)
        total, vencido, riesgo, txt, color = indicadores(buckets)

        generar_pdf(cliente, nit, buckets, total, vencido, riesgo, txt, color)

        return FileResponse(PDF)

    except Exception as e:
        return HTMLResponse(f"<h3>Error:</h3><pre>{str(e)}</pre>")