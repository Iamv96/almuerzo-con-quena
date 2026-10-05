# 🍽️ División de Cuentas - Almuerzo con Quena

Aplicación web interactiva para la división de los gastos del Almuerzo con Quena.

📱 **Enlace público para celulares y computadores:**  
👉 **[https://iamv96.github.io/almuerzo-con-quena/](https://iamv96.github.io/almuerzo-con-quena/)**

---

## 🌟 Características Principales

- **📱 Diseño 100% Responsivo:** Optimizado para smartphones y computadores.
- **🧾 Respaldo Oficial de la Boleta:** Visualizador integrado (lightbox) de la boleta electrónica de compra de Jumbo Quilpué por $59.230.
- **💳 Tarjetas de Cobranza Agrupadas:** 6 tarjetas de pago organizadas según la responsabilidad de transferencia familiar.
- **💾 Datos Bancarios Integrados:** Acceso rápido para copiar con un clic el RUT, cuenta corriente Bci y correo de Ignacio Molina.

---

## 🧾 Datos de la Compra

- **Comercio:** Cencosud Retail S.A. (Jumbo Quilpué - Av. Ramón Freire 2414)
- **N° Boleta Electrónica:** 3492491858
- **Fecha y Hora:** 04/10/2026 14:02
- **Pagado por:** Ignacio Molina (Tarjeta de Crédito)
- **Total Boleta:** $59.230 (10 productos con descuentos Jumbo Ofertas)

---

## 👥 Porciones y Tarjetas de Pago

El total de $59.230 se divide en **11 porciones** iguales:
$$\text{Cuota por porción} = \frac{\$59.230}{11} \approx \$5.384,55 \longrightarrow \mathbf{\$5.385}$$

### Distribución de las 6 Tarjetas:
1. **Pamela:** 2 porciones (Pamela y Cote) $\rightarrow$ **$10.769**
2. **Alejandra:** 1 porción $\rightarrow$ **$5.385**
3. **Mindy:** 3 porciones (Mindy, Mauricio y Gustavo) $\rightarrow$ **$16.154**
4. **Joaquín:** 2 porciones (Joaquín y Kheissa) $\rightarrow$ **$10.769**
5. **Carla:** 2 porciones (Carla + Maxi y Martín como 1 sola porción compartida) $\rightarrow$ **$10.769**
6. **Ignacio (Organizador):** Pagó con tarjeta $59.230. Le corresponde recaudar **$53.846** en total de los demás familiares.

---

## 📂 Archivos del Proyecto

- `index.html`: Dashboard web interactivo con tarjetas de cobranza y lightbox de la boleta.
- `Cuentas Almuerzo con Quena.xlsx`: Planilla Excel de respaldo con 3 hojas (*Compras*, *Porciones por Integrante*, *División y Tarjetas*).
- `update_excel.py`: Script Python para generar y actualizar la planilla Excel.
- `verify_math.py`: Script para auditoría matemática de cuentas.
- `abrir_dashboard.bat`: Acceso directo para abrir localmente en Windows.
- `comprobantes/boleta_jumbo_almuerzo.jpg`: Imagen de la boleta electrónica.
