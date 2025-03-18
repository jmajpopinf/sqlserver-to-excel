
using ClosedXML.Excel;
using Dapper;
using System.Data.SqlClient;

string connectionString = "Server=PCMANU\\SQLEXPRESS; Database=Bar; User Id=user1; Password=123456789";
string query = "SELECT * FROM Beer";
//string query = "SELECT Beer.BeerID, Beer.Name AS BeerName, Brand.Name AS BrandName FROM Beer INNER JOIN Brand ON Beer.BrandID = Brand.BrandID";

using (SqlConnection conn = new SqlConnection(connectionString))
{
	conn.Open();
	var data = conn.Query(query).AsList();

	if (data.Count == 0)
	{
		Console.WriteLine("No hay datos para exportar");
		return;
	}

	using (XLWorkbook wb = new XLWorkbook())
	{
		var ws = wb.Worksheets.Add("Datos");

		//encabezados
		var headers = data[0] as IDictionary<string, object>;
		int colIndex = 1;

		foreach (var header in headers.Keys)
		{
			ws.Cell(1, colIndex).Value = header;
			ws.Cell(1, colIndex).Style.Font.Bold = true;
			ws.Cell(1, colIndex).Style.Fill.BackgroundColor = XLColor.LightGray;
			colIndex++;
		}

		//agregar filas de datos
		int rowIndex = 2;

		foreach (var row in data)
		{
			colIndex = 1;
			foreach (var value in (IDictionary<string, object>)row)
			{
				ws.Cell(rowIndex, colIndex).Value = value.Value?.ToString() ?? "";
				colIndex++;
			}
			rowIndex++;
		}

		// Aplicar formato de tabla
		var range = ws.RangeUsed();
		range.CreateTable();
		range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

		// Guardar el archivo
		string filePath = Path.Combine(Directory.GetCurrentDirectory(), "datos.xlsx");
		wb.SaveAs(filePath);
		Console.WriteLine($"Archivo Excel guardado en: {filePath}");
	}
}