import pyodbc
import pandas as pd  # Necesario para exportar a Excel
import tkinter as tk
from tkinter import messagebox
from tkinter import simpledialog
from tkinter import ttk

# Ruta a tu base de datos Access
database_path = r"Ruta de la base de datos"

def conectar_bd():
    try:
        conn = pyodbc.connect(
            r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=" + database_path + ";"
        )
        print("Conexión exitosa a la base de datos.")
        return conn
    except pyodbc.Error as e:
        messagebox.showerror("Error", f"Error al conectar a la base de datos: {e}")
        return None

def registrar_producto(conn, producto_id, nombre, descripcion, stock_actual, stock_minimo, unidad_medida):
    cursor = conn.cursor()
    query = """
    INSERT INTO Productos (ID, Nombre, Descripcion, StockActual, StockMinimo, UnidadMedida)
    VALUES (?, ?, ?, ?, ?, ?)
    """
    try:
        cursor.execute(query, (producto_id, nombre, descripcion, stock_actual, stock_minimo, unidad_medida))
        conn.commit()
        messagebox.showinfo("Éxito", f"Producto '{nombre}' registrado correctamente.")
    except pyodbc.Error as e:
        messagebox.showerror("Error", f"No se pudo registrar el producto: {e}")

def eliminar_producto(conn, producto_id):
    cursor = conn.cursor()
    query_entradas = "DELETE FROM Entradas WHERE ProductoID = ?"
    query_salidas = "DELETE FROM Salidas WHERE ProductoID = ?"
    query_productos = "DELETE FROM Productos WHERE ID = ?"
    try:
        # Eliminar registros relacionados en Entradas y Salidas
        cursor.execute(query_entradas, (producto_id,))
        cursor.execute(query_salidas, (producto_id,))
        # Eliminar el producto
        cursor.execute(query_productos, (producto_id,))
        conn.commit()
        messagebox.showinfo("Éxito", f"Producto con ID {producto_id} eliminado correctamente.")
    except pyodbc.Error as e:
        messagebox.showerror("Error", f"No se pudo eliminar el producto: {e}")

def registrar_entrada(conn, producto_id, cantidad, proveedor):
    cursor = conn.cursor()
    query = """
    INSERT INTO Entradas (ProductoID, Cantidad, Fecha, Proveedor)
    VALUES (?, ?, NOW(), ?)
    """
    cursor.execute(query, (producto_id, cantidad, proveedor))
    conn.commit()
    actualizar_stock(conn, producto_id, cantidad, "sumar")
    messagebox.showinfo("Éxito", f"Entrada registrada para el producto con ID {producto_id}.")

def registrar_salida(conn, producto_id, cantidad, cliente):
    cursor = conn.cursor()
    query_stock = "SELECT StockActual FROM Productos WHERE ID = ?"
    cursor.execute(query_stock, (producto_id,))
    resultado = cursor.fetchone()

    if resultado is None:
        messagebox.showerror("Error", f"El producto con ID {producto_id} no existe.")
        return

    stock_actual = resultado[0]
    if cantidad > stock_actual:
        messagebox.showerror("Error", f"No se puede realizar la salida. Stock disponible: {stock_actual}.")
        return

    query = """
    INSERT INTO Salidas (ProductoID, Cantidad, Fecha, Cliente)
    VALUES (?, ?, NOW(), ?)
    """
    cursor.execute(query, (producto_id, cantidad, cliente))
    conn.commit()
    actualizar_stock(conn, producto_id, cantidad, "restar")
    messagebox.showinfo("Éxito", f"Salida registrada para el producto con ID {producto_id}.")

def actualizar_stock(conn, producto_id, cantidad, operacion):
    cursor = conn.cursor()
    if operacion == "sumar":
        query = "UPDATE Productos SET StockActual = StockActual + ? WHERE ID = ?"
    elif operacion == "restar":
        query = "UPDATE Productos SET StockActual = StockActual - ? WHERE ID = ?"
    else:
        raise ValueError("Operación no válida: usa 'sumar' o 'restar'.")
    cursor.execute(query, (cantidad, producto_id))
    conn.commit()

def verificar_alarmas(conn):
    cursor = conn.cursor()
    query = "SELECT Nombre FROM Productos WHERE StockActual < StockMinimo"
    cursor.execute(query)
    productos = cursor.fetchall()
    if productos:
        nombres = "\n".join([producto[0] for producto in productos])
        messagebox.showwarning("Alarmas", f"Productos con bajo stock:\n{nombres}")
    else:
        messagebox.showinfo("Todo en orden", "No hay productos con stock crítico.")

def exportar_productos_excel(conn):
    cursor = conn.cursor()
    query = "SELECT * FROM Productos"
    cursor.execute(query)
    productos = cursor.fetchall()

    if not productos:
        messagebox.showinfo("Sin datos", "No hay datos en la tabla Productos para exportar.")
        return

    columns = [desc[0] for desc in cursor.description]
    productos_lista = [list(producto) for producto in productos]
    df = pd.DataFrame(productos_lista, columns=columns)

    try:
        df.to_excel("productos.xlsx", index=False)
        messagebox.showinfo("Éxito", "Reporte exportado a productos.xlsx")
    except Exception as e:
        messagebox.showerror("Error", f"Error al exportar a Excel: {e}")

def cargar_productos(conn, treeview):
    """Carga los productos en el Treeview"""
    cursor = conn.cursor()
    query = "SELECT ID, Nombre, StockActual, UnidadMedida FROM Productos"
    cursor.execute(query)
    productos = cursor.fetchall()
    treeview.delete(*treeview.get_children())  # Limpia el Treeview
    for producto in productos:
        treeview.insert("", tk.END, values=(producto[0], producto[1], producto[2], producto[3]))

def mostrar_menu(conn):
    def on_registrar_producto():
        producto_id = simpledialog.askinteger("Registrar Producto", "ID del Producto:")
        nombre = simpledialog.askstring("Registrar Producto", "Nombre del Producto:")
        descripcion = simpledialog.askstring("Registrar Producto", "Descripción:")
        stock_actual = simpledialog.askinteger("Registrar Producto", "Stock Actual:")
        stock_minimo = simpledialog.askinteger("Registrar Producto", "Stock Mínimo:")
        unidad_medida = simpledialog.askstring("Registrar Producto", "Unidad de Medida:")

        if producto_id and nombre and descripcion and stock_actual is not None and stock_minimo is not None and unidad_medida:
            registrar_producto(conn, producto_id, nombre, descripcion, stock_actual, stock_minimo, unidad_medida)
            cargar_productos(conn, treeview)

    def on_eliminar_producto():
        producto_id = simpledialog.askinteger("Eliminar Producto", "ID del Producto a eliminar:")
        if producto_id:
            eliminar_producto(conn, producto_id)
            cargar_productos(conn, treeview)

    def on_registrar_entrada():
        producto_id = simpledialog.askinteger("Registrar Entrada", "ID del Producto:")
        cantidad = simpledialog.askinteger("Registrar Entrada", "Cantidad:")
        proveedor = simpledialog.askstring("Registrar Entrada", "Proveedor:")
        if producto_id and cantidad and proveedor:
            registrar_entrada(conn, producto_id, cantidad, proveedor)
            cargar_productos(conn, treeview)

    def on_registrar_salida():
        producto_id = simpledialog.askinteger("Registrar Salida", "ID del Producto:")
        cantidad = simpledialog.askinteger("Registrar Salida", "Cantidad:")
        cliente = simpledialog.askstring("Registrar Salida", "Cliente:")
        if producto_id and cantidad and cliente:
            registrar_salida(conn, producto_id, cantidad, cliente)
            cargar_productos(conn, treeview)

    def on_verificar_alarmas():
        verificar_alarmas(conn)

    def on_exportar_excel():
        exportar_productos_excel(conn)

    root = tk.Tk()
    root.title("Control de Inventarios")
    root.geometry("700x600")
    root.configure(bg="#f2f2f2")

    header = tk.Label(root, text="Sistema de Control de Inventarios", bg="#007BFF", fg="white", font=("Arial", 16), pady=10)
    header.pack(fill=tk.X)

    frame = tk.Frame(root, bg="#f2f2f2")
    frame.pack(pady=10)

    tk.Button(frame, text="Registrar Producto", command=on_registrar_producto, width=25, bg="#007bff", fg="white", font=("Arial", 12)).pack(pady=5)
    tk.Button(frame, text="Eliminar Producto", command=on_eliminar_producto, width=25, bg="#dc3545", fg="white", font=("Arial", 12)).pack(pady=5)
    tk.Button(frame, text="Registrar Entrada", command=on_registrar_entrada, width=25, bg="#28a745", fg="white", font=("Arial", 12)).pack(pady=5)
    tk.Button(frame, text="Registrar Salida", command=on_registrar_salida, width=25, bg="#ffc107", fg="black", font=("Arial", 12)).pack(pady=5)
    tk.Button(frame, text="Verificar Alarmas", command=on_verificar_alarmas, width=25, bg="#dc3545", fg="white", font=("Arial", 12)).pack(pady=5)
    tk.Button(frame, text="Exportar a Excel", command=on_exportar_excel, width=25, bg="#17a2b8", fg="white", font=("Arial", 12)).pack(pady=5)

    tree_label = tk.Label(root, text="Lista de Productos", bg="#f2f2f2", font=("Arial", 14))
    tree_label.pack(pady=5)

    treeview = ttk.Treeview(root, columns=("ID", "Nombre", "Stock Actual", "Unidad"), show="headings", height=15)
    treeview.heading("ID", text="ID")
    treeview.heading("Nombre", text="Nombre")
    treeview.heading("Stock Actual", text="Stock Actual")
    treeview.heading("Unidad", text="Unidad")
    treeview.column("ID", width=50, anchor="center")
    treeview.column("Nombre", width=200, anchor="w")
    treeview.column("Stock Actual", width=100, anchor="center")
    treeview.column("Unidad", width=100, anchor="center")
    treeview.pack(pady=10)

    cargar_productos(conn, treeview)

    tk.Button(root, text="Salir", command=root.destroy, width=20, bg="#6c757d", fg="white", font=("Arial", 12)).pack(pady=20)

    root.mainloop()

if __name__ == "__main__":
    conn = conectar_bd()
    if conn:
        mostrar_menu(conn)
        conn.close()
