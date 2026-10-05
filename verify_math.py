import openpyxl

def verify():
    wb = openpyxl.load_workbook("Cuentas Almuerzo con Quena.xlsx", data_only=True)
    
    # Compras
    s1 = wb["Compras"]
    total_boleta = s1["B6"].value
    assert total_boleta == 59230, f"Error en total boleta: {total_boleta}"
    
    items = {
        "Spaghetti 5 Lucchetti": 3000,
        "Salsa Tomate Mallo": 3980,
        "Carne Molida 7%": 13980,
        "Queso Colún Reggianito": 3780,
        "Bebida Coca-Cola Sin Azúcar 1.5L": 3090,
        "Bebida Sprite Zero": 2590,
        "Champiñón Bandeja": 1690,
        "Carne Molida 10%": 6690,
        "Torta Merengue Lúcuma": 19990,
        "Servilleta Nova": 440
    }
    assert sum(items.values()) == 59230, f"La suma de items no da 59.230: {sum(items.values())}"
    
    # Porciones
    porciones_total = 11
    cuota_por_porcion = 59230 / porciones_total # 5384.5454...
    
    # Tarjetas
    tarjetas = {
        "Pamela": 2,    # Pamela + Cote
        "Alejandra": 1,
        "Ignacio": 1,
        "Mindy": 3,   # Mindy + Mauricio + Gustavo
        "Joaquín": 2, # Joaquín + Kheissa
        "Carla": 2    # Carla + (Maxi y Martín juntos = 1)
    }
    assert sum(tarjetas.values()) == 11, f"La suma de porciones debe ser 11: {sum(tarjetas.values())}"
    
    cuotas_tarjetas = {k: round(cuota_por_porcion * v) for k, v in tarjetas.items()}
    
    print(f"Total boleta: ${total_boleta:,}")
    print(f"Total porciones: {porciones_total}")
    print(f"Cuota unitaria por porción: ${cuota_por_porcion:.2f} (~${round(cuota_por_porcion):,})")
    print("Cuotas por tarjeta:")
    for t, c in cuotas_tarjetas.items():
        print(f" - {t} ({tarjetas[t]} porc.): ${c:,}")
        
    total_recaudacion = sum(c for t, c in cuotas_tarjetas.items() if t != "Ignacio")
    print(f"Total recaudado de terceros para Ignacio: ${total_recaudacion:,}")
    print("Verificación matemática completada con éxito.")

if __name__ == "__main__":
    verify()
