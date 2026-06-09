# Manejo de Pedidos — BuscadorExcel

Aplicación de escritorio para buscar y filtrar pedidos almacenados en archivos Excel. Lee múltiples `.xlsx` de una carpeta, muestra cliente, artículos, total y estado de pago, y permite abrir el archivo directamente desde la tabla.

![C#](https://img.shields.io/badge/C%23-239120?style=for-the-badge&logo=c-sharp&logoColor=white)
![.NET](https://img.shields.io/badge/.NET_8-512BD4?style=for-the-badge&logo=dotnet&logoColor=white)
![Windows](https://img.shields.io/badge/Windows-0078D6?style=for-the-badge&logo=windows&logoColor=white)

---

## Funcionalidades

- Seleccionar carpeta de archivos Excel de pedidos
- Buscar pedidos por nombre de cliente
- Filtrar por estado de pago: **Pagados** / **Pendientes**
- Calcular total y comisión por pedido automáticamente
- Ordenar resultados por número de archivo
- Abrir el Excel original con un click desde la tabla

---

## Stack

| Capa | Tecnología |
|------|-----------|
| UI | Windows Forms (.NET 8) |
| Lectura Excel | ClosedXML |
| Plataforma | Windows 10/11 |

---

## Correr localmente

### Requisitos
- Windows 10/11
- .NET 8 SDK o Runtime

```bash
git clone https://github.com/Mickaell22/Manejo_De_Pedidos.git
cd Manejo_De_Pedidos
```

Abrir `Manejo_De_Pedidos.sln` en Visual Studio 2022 y compilar (`Ctrl+Shift+B`), o:

```bash
dotnet run --project BuscadorExcel/BuscadorExcel.csproj
```
