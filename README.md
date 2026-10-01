# SimpleExcel 1.2

Exportación XLSX sencilla para .NET 8.

\# SimpleExcel



SimpleExcel es una biblioteca ligera para .NET enfocada en la generación sencilla de archivos Excel `.xlsx` a partir de colecciones de objetos o `DataTable`.



El objetivo principal es mantener una API simple para escenarios comunes de exportación, por ejemplo, descargar en Excel la información mostrada en tablas, reportes o consultas de una aplicación ASP.NET Core.



SimpleExcel permite generar archivos Excel sin tener Microsoft Excel instalado y sin exponer directamente la complejidad de Open XML al proyecto consumidor.



\## Características



\- Exportación de `IEnumerable<T>` a `.xlsx`.

\- Exportación de `List<T>` a `.xlsx`.

\- Exportación de `DataTable` a `.xlsx`.

\- Tipos de datos nativos de Excel.

\- Encabezados automáticos.

\- Cambio del texto de los encabezados.

\- Exclusión de columnas.

\- Orden personalizado de columnas.

\- Formatos personalizados para fechas y números.

\- Autoajuste del ancho de columnas.

\- Ancho mínimo y máximo.

\- Wrap Text.

\- Encabezados en negritas.

\- Color de fondo opcional en encabezados.

\- Autofiltro.

\- Encabezado congelado.

\- Salida como `byte\[]` o `Stream`.

\- API especialmente sencilla para ASP.NET Core MVC.

\- No requiere Microsoft Excel instalado.



\## Requisitos



\- .NET 8 o superior.

\- DocumentFormat.OpenXml.



\## Instalación



Desde NuGet:



```bash

dotnet add package SimpleExcel

```



O utilizando Package Manager:



```powershell

Install-Package SimpleExcel

```



Si utilizas un servidor NuGet privado, configura previamente el origen correspondiente.



\## Uso básico



Teniendo una colección de objetos:



```csharp

public class Empleado

{

&#x20;   public int NumeroEmpleado { get; set; }



&#x20;   public string Nombre { get; set; } = string.Empty;



&#x20;   public string Departamento { get; set; } = string.Empty;



&#x20;   public DateTime FechaIngreso { get; set; }

}

```



Puedes generar el archivo simplemente con:



```csharp

using SimpleExcel;



var archivo = Excel.Write(empleados);

```



SimpleExcel utilizará automáticamente las propiedades públicas del objeto como columnas.



\## ASP.NET Core MVC



Un caso común es ejecutar una consulta o Stored Procedure, obtener una colección de objetos y descargarla desde un Controller:



```csharp

\[HttpGet]

public IActionResult DescargarExcel()

{

&#x20;   var empleados = \_repository.ObtenerEmpleados();



&#x20;   var archivo = Excel.Write(empleados);



&#x20;   return File(

&#x20;       archivo,

&#x20;       Excel.ContentType,

&#x20;       "Empleados.xlsx");

}

```



No es necesario crear manualmente el `MemoryStream`, definir el MIME type ni trabajar directamente con Open XML.



\## Configuración de columnas



SimpleExcel permite configurar las columnas sin modificar tus DTOs o modelos.



\### Ignorar propiedades



```csharp

var archivo = Excel.Write(perifericos, options =>

{

&#x20;   options.Column(x => x.idCatPeriferico)

&#x20;       .Ignore();



&#x20;   options.Column(x => x.idCatUsuario)

&#x20;       .Ignore();

});

```



Las demás propiedades se exportarán automáticamente.



\## Cambiar encabezados



Puedes cambiar el texto mostrado en Excel mediante `Header()`:



```csharp

var archivo = Excel.Write(perifericos, options =>

{

&#x20;   options.Column(x => x.nombre)

&#x20;       .Header("Nombre");



&#x20;   options.Column(x => x.serie)

&#x20;       .Header("Número de serie");



&#x20;   options.Column(x => x.fechaRegistro)

&#x20;       .Header("Fecha de registro");

});

```



El nombre de la propiedad del objeto no se modifica.



\## Ordenar columnas



A partir de SimpleExcel 1.1 puedes controlar el orden de las columnas mediante `Order()`:



```csharp

var archivo = Excel.Write(perifericos, options =>

{

&#x20;   options.Column(x => x.nombre)

&#x20;       .Header("Nombre")

&#x20;       .Order(1);



&#x20;   options.Column(x => x.categoria)

&#x20;       .Header("Categoría")

&#x20;       .Order(2);



&#x20;   options.Column(x => x.serie)

&#x20;       .Header("Número de serie")

&#x20;       .Order(3);

});

```



Las propiedades que no tengan un `Order()` explícito también serán exportadas y conservarán su orden natural después de las columnas ordenadas explícitamente.



\## Formatos



A partir de SimpleExcel 1.2 puedes aplicar formatos nativos de Excel mediante `Format()`.



\### Fechas



```csharp

options.Column(x => x.fechaRegistro)

&#x20;   .Header("Fecha de registro")

&#x20;   .Order(5)

&#x20;   .Format("dd/MM/yyyy");

```



También puedes incluir la hora:



```csharp

options.Column(x => x.fechaRegistro)

&#x20;   .Format("dd/MM/yyyy HH:mm");

```



La fecha se conserva como un valor de fecha de Excel. `Format()` controla solamente su presentación; no convierte la fecha a texto.



\### Números



```csharp

options.Column(x => x.importe)

&#x20;   .Format("#,##0.00");

```



\### Moneda



```csharp

options.Column(x => x.importe)

&#x20;   .Format("$#,##0.00");

```



\### Porcentaje



```csharp

options.Column(x => x.porcentaje)

&#x20;   .Format("0.00%");

```



\## Ejemplo completo



```csharp

\[HttpGet]

public IActionResult DescargarExcel()

{

&#x20;   var perifericos = \_repository.ObtenerPerifericos();



&#x20;   var archivo = Excel.Write(perifericos, options =>

&#x20;   {

&#x20;       options.SheetName = "Periféricos";



&#x20;       options.Column(x => x.idCatPeriferico)

&#x20;           .Ignore();



&#x20;       options.Column(x => x.idCatUsuario)

&#x20;           .Ignore();



&#x20;       options.Column(x => x.nombre)

&#x20;           .Header("Nombre")

&#x20;           .Order(1);



&#x20;       options.Column(x => x.categoria)

&#x20;           .Header("Categoría")

&#x20;           .Order(2);



&#x20;       options.Column(x => x.subcategoria)

&#x20;           .Header("Subcategoría")

&#x20;           .Order(3);



&#x20;       options.Column(x => x.marca)

&#x20;           .Header("Marca")

&#x20;           .Order(4);



&#x20;       options.Column(x => x.modelo)

&#x20;           .Header("Modelo")

&#x20;           .Order(5);



&#x20;       options.Column(x => x.serie)

&#x20;           .Header("Número de serie")

&#x20;           .Order(6);



&#x20;       options.Column(x => x.fechaRegistro)

&#x20;           .Header("Fecha de registro")

&#x20;           .Order(7)

&#x20;           .Format("dd/MM/yyyy");



&#x20;       options.AutoFitColumns = true;

&#x20;       options.WrapText = true;



&#x20;       options.Header.Bold = true;

&#x20;       options.Header.BackgroundColor = "#D9EAF7";

&#x20;       options.Header.AutoFilter = true;

&#x20;   });



&#x20;   return File(

&#x20;       archivo,

&#x20;       Excel.ContentType,

&#x20;       "Perifericos.xlsx");

}

```



\## Opciones generales



Puedes configurar el nombre de la hoja:



```csharp

options.SheetName = "Periféricos";

```



Autoajustar columnas:



```csharp

options.AutoFitColumns = true;

```



Habilitar Wrap Text:



```csharp

options.WrapText = true;

```



Congelar el encabezado:



```csharp

options.FreezeHeader = true;

```



Mostrar encabezados:



```csharp

options.IncludeHeaders = true;

```



\## Configuración del encabezado



Puedes controlar las características básicas del encabezado:



```csharp

options.Header.Bold = true;



options.Header.BackgroundColor = "#D9EAF7";



options.Header.AutoFilter = true;

```



El color de fondo acepta valores hexadecimales.



Por ejemplo:



```text

\#D9EAF7

```



\## Ancho de columnas



SimpleExcel calcula automáticamente un ancho apropiado con base en el contenido.



```csharp

options.AutoFitColumns = true;

```



También puedes controlar los límites:



```csharp

options.ColumnWidth.Minimum = 8;



options.ColumnWidth.Maximum = 60;

```



Esto evita que una columna que contenga textos demasiado largos genere un ancho excesivo.



\## DataTable



También puedes exportar directamente un `DataTable`:



```csharp

DataTable datos = ObtenerDatos();



var archivo = Excel.Write(datos);

```



Por ejemplo, desde ASP.NET Core:



```csharp

\[HttpGet]

public IActionResult DescargarExcel()

{

&#x20;   DataTable datos = \_repository.ObtenerDatos();



&#x20;   var archivo = Excel.Write(datos);



&#x20;   return File(

&#x20;       archivo,

&#x20;       Excel.ContentType,

&#x20;       "Reporte.xlsx");

}

```



\## Stream



Además de obtener un `byte\[]`, SimpleExcel permite escribir el documento en un `Stream`:



```csharp

using var stream = new MemoryStream();



Excel.Write(

&#x20;   stream,

&#x20;   empleados,

&#x20;   options =>

&#x20;   {

&#x20;       options.SheetName = "Empleados";

&#x20;   });

```



\## Atributos



Aunque se recomienda la configuración fluida para mantener los modelos independientes de la presentación de Excel, SimpleExcel también proporciona atributos.



\### ExcelColumn



```csharp

public class Empleado

{

&#x20;   \[ExcelColumn("Número", Order = 1)]

&#x20;   public int NumeroEmpleado { get; set; }



&#x20;   \[ExcelColumn("Nombre", Order = 2)]

&#x20;   public string Nombre { get; set; } =

