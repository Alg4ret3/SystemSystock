import win32print
import win32ui
import win32con
import datetime

def imprimir_ticket_80mm(
    invoice_number,
    client_name,
    client_id,
    client_phone,
    client_address,
    items,
    subtotal,
    delivery_fee,
    descuento,
    total,
    pago,
    payment_method,
    p_unit_label="P.Unit",
    is_credit=False,
    due_date=None
):
    """
    Genera e imprime un ticket de venta en una impresora térmica de 80mm.
    """
    # Configuración inicial de la empresa
    empresa_nombre = "LadyNailShop"
    empresa_direccion = "Pasto, Colombia"
    empresa_telefono = "+57 316-144-44-74"
    max_lines_per_page = 30
    current_line = 0

    # Obtener la fecha actual
    fecha_actual = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    # Formatear valores monetarios
    subtotal_formateado = f"${subtotal:,.2f}"
    total_formateado = f"${total:,.2f}"
    descuento_formateado = f"${descuento:,.2f}"
    
    # Formatear el pago
    if isinstance(pago, str) and "/" in pago:
        try:
            pagos_split = [float(p.replace(".", "").replace(",", ".")) for p in pago.split("/")]
            if len(pagos_split) >= 2:
                pago_formateado = f"Efectivo: ${pagos_split[0]:,.2f}\nTransferencia: ${pagos_split[1]:,.2f}"
            else:
                pago_formateado = f" ${pagos_split[0]:,.2f}"
        except (ValueError, IndexError):
            pago_formateado = f" {pago}"
    else:
        try:
            val_pago = float(str(pago).replace(".", "").replace(",", "."))
            pago_formateado = f" ${val_pago:,.2f}"
        except ValueError:
            pago_formateado = f" {pago}"

    # Formatear el costo de envío
    delivery_fee = float(delivery_fee)
    if delivery_fee.is_integer():
        delivery_fee_formateado = f"${int(delivery_fee):,.0f}"
    else:
        delivery_fee_formateado = f"${delivery_fee:,.2f}"

    # Limitar la dirección del cliente
    direccion = client_address
    direccion_linea1 = direccion[:35]
    direccion_linea2 = direccion[35:] if len(direccion) > 35 else ""

    try:
        # Obtener la impresora predeterminada
        impresora = win32print.GetDefaultPrinter()
        hDC = win32ui.CreateDC()
        hDC.CreatePrinterDC(impresora)

        # Crear un documento de impresión
        hDC.StartDoc("Ticket de Venta")
        hDC.StartPage()

        # Fuentes
        font_encabezado = win32ui.CreateFont({
            "name": "Lucida Console",
            "height": 28,
            "weight": win32con.FW_BOLD
        })
        
        font_size = 18
        line_height = font_size + 10
        font_normal = win32ui.CreateFont({
            "name": "Lucida Console",
            "height": font_size,
            "weight": win32con.FW_BOLD
        })

        # Centrado
        printer_width = hDC.GetDeviceCaps(win32con.HORZRES)
        center_x = printer_width // 2

        # Imprimir encabezado
        hDC.SelectObject(font_encabezado)
        for i, linea in enumerate([empresa_nombre, empresa_direccion, empresa_telefono, fecha_actual]):
            text_size = hDC.GetTextExtent(linea)
            text_width = text_size[0]
            hDC.TextOut(center_x - (text_width // 2), 50 + (i * line_height), linea)
        
        # Posición inicial después del encabezado
        x, y = 2, 50 + (4 * line_height) + line_height
        
        hDC.SelectObject(font_normal)
        separator = "-" * 80 # Línea separadora simple

        # Información de la factura y cliente
        hDC.TextOut(x, y, separator)
        y += line_height
        
        # Título de factura
        ticket_title = "crédito" if is_credit else "TICKET DE VENTA"
        hDC.TextOut(x, y, ticket_title)
        y += line_height
        
        hDC.TextOut(x, y, f"COT No. {invoice_number}")
        y += line_height
        hDC.TextOut(x, y, f"Cliente: {client_name}")
        y += line_height
        hDC.TextOut(x, y, f"Cédula: {client_id}")
        y += line_height
        hDC.TextOut(x, y, f"Teléfono: {client_phone}")
        y += line_height
        hDC.TextOut(x, y, f"Dirección: {direccion_linea1}")
        y += line_height
        if direccion_linea2:
            hDC.TextOut(x, y, direccion_linea2)
            y += line_height

        hDC.TextOut(x, y, separator)
        y += line_height
        
        # Encabezado de tabla productos
        header = "{:<18} {:>6} {:>10} {:>10}".format("Producto", "Cant.", p_unit_label, "Total")
        hDC.TextOut(x, y, header)
        y += line_height

        # Productos
        for item in items:
            nombre_producto = item[0].strip().replace('\n', ' ')[:18].ljust(18)
            cantidad = str(item[1])
            precio_unitario = f"{item[2]:,.0f}".replace(",", ".")
            total_producto = f"{item[3]:,.0f}".replace(",", ".")

            linea = "{:<18} {:>6} {:>10} {:>10}".format(
                nombre_producto, cantidad, precio_unitario, total_producto
            )
            hDC.TextOut(x, y, linea)
            y += line_height
            current_line += 1

            if current_line >= max_lines_per_page:
                hDC.EndPage()
                hDC.StartPage()
                y = 2
                current_line = 0

        # Totales
        hDC.TextOut(x, y, separator)
        y += line_height
        hDC.TextOut(x, y, f"Subtotal: {subtotal_formateado}")
        y += line_height
        hDC.TextOut(x, y, f"Envío: {delivery_fee_formateado}")
        y += line_height
        if descuento > 0:
            hDC.TextOut(x, y, f"Descuento: {descuento_formateado}")
            y += line_height
        hDC.TextOut(x, y, f"Total: {total_formateado}")
        y += line_height

        if is_credit and due_date:
            # Formatear fecha límite si es datetime
            if isinstance(due_date, datetime.datetime):
                due_date_str = due_date.strftime("%d/%m/%Y")
            else:
                due_date_str = str(due_date)
            hDC.TextOut(x, y, f"Fecha Límite: {due_date_str}")
            y += line_height
        
            
        hDC.TextOut(x, y, f"Método de Pago: {payment_method}")
        y += line_height
        hDC.TextOut(x, y, separator)
        y += line_height
        
        msg_thanks = "¡Gracias por tu compra!"
        text_size = hDC.GetTextExtent(msg_thanks)
        hDC.TextOut(center_x - (text_size[0] // 2), y, msg_thanks)
        y += line_height
        hDC.TextOut(x, y, separator)

        # Finalizar
        hDC.EndPage()
        hDC.EndDoc()
        hDC.DeleteDC()
        return True

    except Exception as e:
        print(f"Error al imprimir ticket: {e}")
        if 'hDC' in locals():
            hDC.DeleteDC()
        raise e
