# SimpleExcel 1.2
Exportación XLSX sencilla para .NET 8.

## 1.2
Agrega `.Format()` a la configuración fluida. El valor conserva su tipo nativo de Excel; el formato controla únicamente su presentación.

```csharp
var bytes = Excel.Write(perifericos, options =>
{
    options.Column(x => x.idCatPeriferico).Ignore();
    options.Column(x => x.idCatUsuario).Ignore();
    options.Column(x => x.nombre).Header("Nombre").Order(1);
    options.Column(x => x.fechaRegistro).Header("Fecha de registro").Order(2).Format("dd/MM/yyyy");
});
```

También admite formatos como `#,##0.00`, `$#,##0.00` y `0.00%`.
