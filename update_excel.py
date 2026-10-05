import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def update_excel(file_path):
    wb = openpyxl.Workbook()
    wb.remove(wb.active) # Eliminar hoja por defecto
    
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid") # Dark blue
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    bold_font = Font(name="Calibri", size=11, bold=True)
    regular_font = Font(name="Calibri", size=11)
    money_format = "$#,##0"
    
    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )
    
    # ----------------------------------------------------
    # HOJA 1: Compras & Boleta
    # ----------------------------------------------------
    ws1 = wb.create_sheet(title="Compras")
    
    ws1["A1"] = "Evento"
    ws1["B1"] = "Almuerzo con Quena"
    ws1["A2"] = "Comercio"
    ws1["B2"] = "Cencosud Retail S.A. (Jumbo Quilpué - Freire 2414)"
    ws1["A3"] = "N° Boleta Electrónica"
    ws1["B3"] = "3492491858"
    ws1["A4"] = "Fecha y Hora"
    ws1["B4"] = "04/10/2026 14:02"
    ws1["A5"] = "Pagado por"
    ws1["B5"] = "Ignacio Molina (Tarjeta de Crédito)"
    ws1["A6"] = "Total Compra Boleta"
    ws1["B6"] = 59230
    
    for r in range(1, 7):
        ws1[f"A{r}"].font = bold_font
        if r == 6:
            ws1[f"B{r}"].font = bold_font
            ws1[f"B{r}"].number_format = money_format
        else:
            ws1[f"B{r}"].font = regular_font

    items = [
        ("7802500000053", "Spaghetti 5 Lucchetti (4 un)", 4280, 1280, 3000),
        ("7802300000154", "Salsa Tomate Mallo (6 un)", 5340, 1360, 3980),
        ("7804684740107", "Carne Molida 7% (2 un)", 13980, 0, 13980),
        ("7802920423814", "Queso Colún Reggianito (2 un)", 4420, 640, 3780),
        ("7801610350409", "Bebida Coca-Cola Sin Azúcar 1.5L (2 un)", 4580, 1490, 3090),
        ("7801610175217", "Bebida Sprite Zero", 2590, 0, 2590),
        ("7801365000284", "Champiñón Bandeja", 1690, 0, 1690),
        ("7804684740008", "Carne Molida 10%", 6690, 0, 6690),
        ("7803455005131", "Torta Merengue Lúcuma", 19990, 0, 19990),
        ("7806500241416", "Servilleta Nova", 440, 0, 440)
    ]
    
    header_row_idx = 8
    headers_compras = ["Código", "Producto / Detalle", "Precio Lista ($)", "Descuento ($)", "Total Neto ($)"]
    for c_idx, h in enumerate(headers_compras, start=1):
        cell = ws1.cell(row=header_row_idx, column=c_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
    
    for idx, (code, name, list_p, disc, net_p) in enumerate(items, start=9):
        c_code = ws1.cell(row=idx, column=1, value=code)
        c_code.font = regular_font
        c_code.alignment = Alignment(horizontal="center")
        c_code.border = thin_border
        
        c_name = ws1.cell(row=idx, column=2, value=name)
        c_name.font = regular_font
        c_name.border = thin_border
        
        c_list = ws1.cell(row=idx, column=3, value=list_p)
        c_list.font = regular_font
        c_list.number_format = money_format
        c_list.border = thin_border
        
        c_disc = ws1.cell(row=idx, column=4, value=disc)
        c_disc.font = regular_font
        c_disc.number_format = money_format
        c_disc.border = thin_border
        
        c_net = ws1.cell(row=idx, column=5, value=net_p)
        c_net.font = bold_font
        c_net.number_format = money_format
        c_net.border = thin_border

    tot_row_compras = 9 + len(items)
    ws1.cell(row=tot_row_compras, column=1, value="").border = thin_border
    ws1.cell(row=tot_row_compras, column=2, value="TOTAL GENERAL BOLETA").font = bold_font
    ws1.cell(row=tot_row_compras, column=2).border = thin_border
    
    c_tot_list = ws1.cell(row=tot_row_compras, column=3, value=f"=SUM(C9:C{tot_row_compras-1})")
    c_tot_list.font = bold_font
    c_tot_list.number_format = money_format
    c_tot_list.border = thin_border
    
    c_tot_disc = ws1.cell(row=tot_row_compras, column=4, value=f"=SUM(D9:D{tot_row_compras-1})")
    c_tot_disc.font = bold_font
    c_tot_disc.number_format = money_format
    c_tot_disc.border = thin_border
    
    tot_cell = ws1.cell(row=tot_row_compras, column=5, value=f"=SUM(E9:E{tot_row_compras-1})")
    tot_cell.font = bold_font
    tot_cell.number_format = money_format
    tot_cell.fill = PatternFill(start_color="FEF08A", end_color="FEF08A", fill_type="solid")
    tot_cell.border = thin_border
    
    for col in ws1.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws1.column_dimensions[col_letter].width = max(max_len + 4, 16)

    # ----------------------------------------------------
    # HOJA 2: Porciones y Consumo
    # ----------------------------------------------------
    ws2 = wb.create_sheet(title="Porciones por Integrante")
    
    # 11 porciones en total:
    # Pamela (1), Cote (1), Alejandra (1), Ignacio (1)
    # Mindy (1), Mauricio (1), Gustavo (1) -> grupo Mindy
    # Joaquín (1), Kheissa (1) -> grupo Joaquín
    # Carla (1), Maxi y Martín (1) -> grupo Carla
    participantes = [
        ("Pamela", 1, "Pamela (Grupo Familiar)"),
        ("Cote", 1, "Pamela (Grupo Familiar)"),
        ("Alejandra", 1, "Alejandra"),
        ("Ignacio", 1, "Ignacio"),
        ("Mindy", 1, "Mindy (Grupo Familiar)"),
        ("Mauricio", 1, "Mindy (Grupo Familiar)"),
        ("Gustavo", 1, "Mindy (Grupo Familiar)"),
        ("Joaquín", 1, "Joaquín (Grupo Familiar)"),
        ("Kheissa", 1, "Joaquín (Grupo Familiar)"),
        ("Carla", 1, "Carla (Grupo Familiar)"),
        ("Maxi y Martín (compartido)", 1, "Carla (Grupo Familiar)")
    ]
    
    headers_ws2 = ["Asistente / Porción", "N° Porciones", "Tarjeta de Cobranza Asignada", "Cuota Individual Calculada"]
    for c_idx, h in enumerate(headers_ws2, start=1):
        cell = ws2.cell(row=1, column=c_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
        
    for r_idx, (nombre, porciones, tarjeta) in enumerate(participantes, start=2):
        c1 = ws2.cell(row=r_idx, column=1, value=nombre)
        c1.font = bold_font
        c1.border = thin_border
        
        c2 = ws2.cell(row=r_idx, column=2, value=porciones)
        c2.font = regular_font
        c2.alignment = Alignment(horizontal="center")
        c2.border = thin_border
        
        c3 = ws2.cell(row=r_idx, column=3, value=tarjeta)
        c3.font = regular_font
        c3.border = thin_border
        
        c4 = ws2.cell(row=r_idx, column=4, value=f"=ROUND(Compras!$E${tot_row_compras} / $B$13 * B{r_idx}, 0)")
        c4.font = bold_font
        c4.number_format = money_format
        c4.border = thin_border

    tot_p_row = len(participantes) + 2
    ws2.cell(row=tot_p_row, column=1, value="Total Porciones").font = bold_font
    ws2.cell(row=tot_p_row, column=1).border = thin_border
    
    t_porc = ws2.cell(row=tot_p_row, column=2, value=f"=SUM(B2:B{tot_p_row-1})")
    t_porc.font = bold_font
    t_porc.alignment = Alignment(horizontal="center")
    t_porc.fill = PatternFill(start_color="FEF08A", end_color="FEF08A", fill_type="solid")
    t_porc.border = thin_border
    
    ws2.cell(row=tot_p_row, column=3, value="").border = thin_border
    
    t_cuotas = ws2.cell(row=tot_p_row, column=4, value=f"=SUM(D2:D{tot_p_row-1})")
    t_cuotas.font = bold_font
    t_cuotas.number_format = money_format
    t_cuotas.fill = PatternFill(start_color="FEF08A", end_color="FEF08A", fill_type="solid")
    t_cuotas.border = thin_border

    for col in ws2.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws2.column_dimensions[col_letter].width = max(max_len + 4, 18)

    # ----------------------------------------------------
    # HOJA 3: Tarjetas de Cobranza & Liquidación
    # ----------------------------------------------------
    ws3 = wb.create_sheet(title="División y Tarjetas")
    
    tarjetas_info = [
        ("Pamela", 2, "Pamela (1) + Cote (1)", "Transfiere a Ignacio"),
        ("Alejandra", 1, "Alejandra (1)", "Transfiere a Ignacio"),
        ("Ignacio", 1, "Ignacio (1)", "Organizador (Recauda $53.845)"),
        ("Mindy", 3, "Mindy (1) + Mauricio (1) + Gustavo (1)", "Transfiere a Ignacio"),
        ("Joaquín", 2, "Joaquín (1) + Kheissa (1)", "Transfiere a Ignacio"),
        ("Carla", 2, "Carla (1) + Maxi y Martín (1)", "Transfiere a Ignacio")
    ]
    
    headers_ws3 = [
        "Tarjeta / Responsable", "Porciones", "Personas Incluidas",
        "Total Cuota ($)", "Aportado en Boleta ($)", "Saldo a Liquidar ($)",
        "Acción / Destinatario", "Estado"
    ]
    
    for c_idx, h in enumerate(headers_ws3, start=1):
        cell = ws3.cell(row=1, column=c_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
        
    for r_idx, (nombre, porc, detalle, accion) in enumerate(tarjetas_info, start=2):
        c_nom = ws3.cell(row=r_idx, column=1, value=nombre)
        c_nom.font = bold_font
        c_nom.border = thin_border
        
        c_porc = ws3.cell(row=r_idx, column=2, value=porc)
        c_porc.font = bold_font
        c_porc.alignment = Alignment(horizontal="center")
        c_porc.border = thin_border
        
        c_det = ws3.cell(row=r_idx, column=3, value=detalle)
        c_det.font = regular_font
        c_det.border = thin_border
        
        # Cuota = ROUND(Compras!$E${tot_row_compras} / 'Porciones por Integrante'!$B$13 * B{r_idx}, 0)
        c_cuota = ws3.cell(row=r_idx, column=4, value=f"=ROUND(Compras!$E${tot_row_compras} / 'Porciones por Integrante'!$B$13 * B{r_idx}, 0)")
        c_cuota.font = bold_font
        c_cuota.number_format = money_format
        c_cuota.border = thin_border
        
        # Aportado
        aporte = 59230 if nombre == "Ignacio" else 0
        c_ap = ws3.cell(row=r_idx, column=5, value=aporte)
        c_ap.font = regular_font
        c_ap.number_format = money_format
        c_ap.border = thin_border
        
        # Saldo a liquidar
        if nombre == "Ignacio":
            c_sal = ws3.cell(row=r_idx, column=6, value=f"=E{r_idx}-D{r_idx}")
            c_sal.font = Font(name="Calibri", size=11, bold=True, color="1E3A8A")
            c_st = ws3.cell(row=r_idx, column=8, value="Organizador")
            c_st.fill = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
            c_st.font = Font(name="Calibri", size=11, bold=True, color="1E3A8A")
        else:
            c_sal = ws3.cell(row=r_idx, column=6, value=f"=D{r_idx}-E{r_idx}")
            c_sal.font = Font(name="Calibri", size=11, bold=True, color="047857")
            c_st = ws3.cell(row=r_idx, column=8, value="Pendiente")
            c_st.fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
            c_st.font = Font(name="Calibri", size=11, bold=True, color="991B1B")
            
        c_sal.number_format = money_format
        c_sal.border = thin_border
        
        c_acc = ws3.cell(row=r_idx, column=7, value=accion)
        c_acc.font = regular_font
        c_acc.alignment = Alignment(horizontal="center")
        c_acc.border = thin_border
        
        c_st.alignment = Alignment(horizontal="center")
        c_st.border = thin_border

    tot_tarjetas_row = len(tarjetas_info) + 2
    ws3.cell(row=tot_tarjetas_row, column=1, value="Totales").font = bold_font
    ws3.cell(row=tot_tarjetas_row, column=1).border = thin_border
    
    t_tot_porc = ws3.cell(row=tot_tarjetas_row, column=2, value=f"=SUM(B2:B{tot_tarjetas_row-1})")
    t_tot_porc.font = bold_font
    t_tot_porc.alignment = Alignment(horizontal="center")
    t_tot_porc.border = thin_border
    
    ws3.cell(row=tot_tarjetas_row, column=3, value="").border = thin_border
    
    t_tot_cuota = ws3.cell(row=tot_tarjetas_row, column=4, value=f"=SUM(D2:D{tot_tarjetas_row-1})")
    t_tot_cuota.font = bold_font
    t_tot_cuota.number_format = money_format
    t_tot_cuota.fill = PatternFill(start_color="FEF08A", end_color="FEF08A", fill_type="solid")
    t_tot_cuota.border = thin_border
    
    t_tot_aporte = ws3.cell(row=tot_tarjetas_row, column=5, value=f"=SUM(E2:E{tot_tarjetas_row-1})")
    t_tot_aporte.font = bold_font
    t_tot_aporte.number_format = money_format
    t_tot_aporte.fill = PatternFill(start_color="FEF08A", end_color="FEF08A", fill_type="solid")
    t_tot_aporte.border = thin_border
    
    t_tot_recauda = ws3.cell(row=tot_tarjetas_row, column=6, value=f'=SUMIF(G2:G{tot_tarjetas_row-1}, "Transfiere a Ignacio", F2:F{tot_tarjetas_row-1})')
    t_tot_recauda.font = bold_font
    t_tot_recauda.number_format = money_format
    t_tot_recauda.fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    t_tot_recauda.border = thin_border
    
    ws3.cell(row=tot_tarjetas_row, column=7, value="Total Recaudación Ignacio").font = bold_font
    ws3.cell(row=tot_tarjetas_row, column=7).alignment = Alignment(horizontal="center")
    ws3.cell(row=tot_tarjetas_row, column=7).border = thin_border
    ws3.cell(row=tot_tarjetas_row, column=8, value="").border = thin_border

    for col in ws3.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws3.column_dimensions[col_letter].width = max(max_len + 4, 18)
        
    wb.save(file_path)
    print(f"Excel guardado con éxito en: {file_path}")

if __name__ == "__main__":
    update_excel("Cuentas Almuerzo con Quena.xlsx")
