# Sistema de Control de Inventarios

Este programa es una herramienta para gestionar el inventario de productos en una bodega. Permite realizar un seguimiento detallado de las entradas, salidas y niveles de stock de los productos, todo a través de una interfaz gráfica amigable.

## **Características**

1. **Registrar Productos**:
   - Agrega nuevos productos a la base de datos con detalles como ID, nombre, descripción, stock inicial, stock mínimo y unidad de medida.

2. **Registrar Entradas**:
   - Permite registrar el ingreso de productos al inventario indicando cantidad y proveedor.
   - Actualiza automáticamente el stock del producto.

3. **Registrar Salidas**:
   - Registra las salidas de productos del inventario especificando cantidad y cliente.
   - Asegura que las salidas no dejen el stock en negativo.

4. **Verificar Alarmas**:
   - Identifica productos cuyo stock está por debajo del nivel mínimo configurado.
   - Muestra alertas visuales de productos en estado crítico.

5. **Exportar a Excel**:
   - Genera un informe completo del inventario en formato Excel.

6. **Interfaz Gráfica (GUI)**:
   - Diseñada con `tkinter`.
   - Incluye botones para todas las funcionalidades y un listado dinámico de productos.

## **Requisitos**

Para ejecutar este programa, asegúrate de tener instalados los siguientes paquetes:

- `pyodbc`
- `pandas`
- `openpyxl`
- `tkinter`

## **Instalación**

1. Clona este repositorio:

   ```bash
   git clone <URL-del-repositorio>
   cd <nombre-del-directorio>
   ```

2. Instala las dependencias necesarias:

   ```bash
   pip install -r requirements.txt
   ```

3. Asegúrate de tener una base de datos en Microsoft Access configurada con las tablas necesarias.

## **Uso**

Ejecuta el archivo principal del programa para abrir la interfaz gráfica:

```bash
python ControlDB.py
```

### **Notas**
- La base de datos debe estar ubicada en la ruta especificada dentro del archivo `ControlDB.py`. Cambia la variable `database_path` si es necesario.
- Asegúrate de que tu sistema tenga los controladores ODBC para Access correctamente instalados.

## **Estructura de Tablas**

### Tabla: `Productos`
- `ID`: Identificador único del producto.
- `Nombre`: Nombre del producto.
- `Descripcion`: Descripción del producto.
- `StockActual`: Cantidad actual en inventario.
- `StockMinimo`: Cantidad mínima permitida antes de activar una alarma.
- `UnidadMedida`: Unidad de medida del producto.

### Tabla: `Entradas`
- `ID`: Identificador único de la entrada.
- `ProductoID`: Referencia al producto.
- `Cantidad`: Cantidad ingresada.
- `Fecha`: Fecha de la entrada.
- `Proveedor`: Nombre del proveedor.

### Tabla: `Salidas`
- `ID`: Identificador único de la salida.
- `ProductoID`: Referencia al producto.
- `Cantidad`: Cantidad retirada.
- `Fecha`: Fecha de la salida.
- `Cliente`: Nombre del cliente.

## **Licencia**

Este proyecto está bajo la licencia none. 
