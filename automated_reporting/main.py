import pandas as pd
import xlsxwriter

# ================= DATOS =================
data = pd.read_csv("tutor_report.csv")
df = pd.DataFrame(data)

# ================= MÉTRICAS =================
df["Earning, USD"] = pd.to_numeric(df["Earning, USD"], errors="coerce")
df["Earning, USD"] = df["Earning, USD"].fillna(0)
df["Type"] = df["Type"].replace({
    "Non-trial lesson": "Regular",
    "Trial": "Prueba"
})

df["Student Location"] = df["Student Location"].replace({
    "Brazil": "Brasil",
    "United States": "Estados Unidos",
    "Philippines": "Filipinas",
    "Hungary": "Hungria",
    "Chile": "Chile",
    "United Kingdom": "Reino Unido",
})
df = df[(df['Type'] == "Prueba") | (df['Type'] == "Regular")].reset_index()
total_ingresos = df['Earning, USD'].sum()
total_clases = df['Service Type'].count()
ingreso_promedio = df['Earning, USD'].mean()
pais_top = df['Student Location'].mode()[0]


# ================= ARCHIVO =================
workbook = xlsxwriter.Workbook('tutor_report.xlsx')

font = 'Segoe UI'

# ================= FORMATOS =================
header = workbook.add_format({
    'bold': True,
    'font_name': font,
    'font_size': 11,
    'bg_color': '#E9EEF7',
    'font_color': '#1F2937',
    'align': 'center',
    'valign': 'vcenter'
})

text = workbook.add_format({
    'font_name': font,
    'font_size': 11,
    'align': 'center',
    'valign': 'vcenter'
})

text_bold = workbook.add_format({
    'font_name': font,
    'font_size': 11,
    'bold': True,
    'align': 'center',
    'valign': 'vcenter'
})

money = workbook.add_format({
    'font_name': font,
    'font_size': 11,
    'num_format': '$#,##0.00',
    'align': 'center',
    'valign': 'vcenter'
})

verde = workbook.add_format({
    'bg_color': '#E8F5E9',
    'font_color': '#2E7D32'
})

rojo = workbook.add_format({
    'bg_color': '#FDECEA',
    'font_color': '#C62828'
})

# ================= HOJA DETALLES =================
ws = workbook.add_worksheet('Detalles')

headers = ['Fecha', 'Alumno', 'País', 'Ingreso ($)', 'Tipo']

for col, h in enumerate(headers):
    ws.write(0, col, h, header)

for i, row in df.iterrows():
    r = i + 1
    ws.write(r, 0, row['Fecha confirmada'], text)
    ws.write(r, 1, row['Estudiante'], text_bold)
    ws.write(r, 2, row['Student Location'], text)
    #ws.write(r, 3, row['Service Type'], text)
    ws.write(r, 3, row['Earning, USD'], money)
    ws.write(r, 4, row['Type'], text)

# Altura (padding visual)
for i in range(len(df) + 1):
    ws.set_row(i, 24)

# Columnas
ws.set_column('A:A', 30)
ws.set_column('B:B', 22)
ws.set_column('C:C', 18)
ws.set_column('D:D', 16)
ws.set_column('E:E', 20)

# UX
last_row = len(df)
ws.freeze_panes(1, 0)
ws.autofilter(0, 0, last_row, 4)

# Condicional
ws.conditional_format(f'D2:D{last_row+1}', {
    'type': 'cell',
    'criteria': '>',
    'value': 0,
    'format': verde
})

ws.conditional_format(f'D2:D{last_row+1}', {
    'type': 'cell',
    'criteria': '==',
    'value': 0,
    'format': rojo
})

# ================= HOJA RESUMEN =================
ws2 = workbook.add_worksheet('Resumen Ejecutivo')

label = workbook.add_format({
    'font_name': font,
    'font_size': 11,
    'bold': True,
    'font_color': '#374151',
    'bg_color': '#EEF2FF',  # 🔥 fondo suave diferenciador
    'align': 'center',
    'valign': 'vcenter'
})

value = workbook.add_format({
    'font_name': font,
    'font_size': 11,
    'bold': True,
    'font_color': '#111827',
    'align': 'center',
    'valign': 'vcenter'
})

# Contenido
ws2.write('A1', 'Total Ingresos', label)
ws2.write('B1', f'${total_ingresos:,.2f}', value)

ws2.write('A2', 'Total Clases', label)
ws2.write('B2', total_clases, value)

ws2.write('A3', 'Ingreso Promedio', label)
ws2.write('B3', f'${ingreso_promedio:,.2f}', value)

ws2.write('A4', 'País Principal', label)
ws2.write('B4', pais_top, value)

# Espaciado
ws2.set_column('A:A', 26)
ws2.set_column('B:B', 18)

for i in range(0, 6):
    ws2.set_row(i, 24)

# ================= GUARDAR =================
workbook.close()

print("🔥 Reporte generado")