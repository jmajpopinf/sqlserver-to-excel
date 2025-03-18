const sql = require('mssql');
const ExcelJS = require('exceljs');

//configuracion de la conexion a la base de datos
const config = {
    user: 'user1',
    password: '123456789',
    server: 'localhost',
    database: 'Bar',
    options: {
        encrypt: false,
        trustServerCertificate: true,
    }
}

//exportar solo una tabla
async function exportToExcel() {
    try {

        // Conectar a SQL Server
        await sql.connect(config);
        const result = await sql.query("SELECT * FROM Beer"); // Ajusta la consulta

        // Crear un nuevo libro de Excel
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet("Datos");

        // Obtener las columnas de la consulta
        const columns = Object.keys(result.recordset[0]);
        worksheet.columns = columns.map((col) => ({ header: col, key: col, width: 20 }));

        // Aplicar formato de tabla
        worksheet.getRow(1).eachCell((cell) => {
            cell.font = { bold: true };
            cell.border = {
                top: { style: 'thin', color: { argb: '000000' } }, // Negro
                left: { style: 'thin', color: { argb: '000000' } },
                bottom: { style: 'thin', color: { argb: '000000' } },
                right: { style: 'thin', color: { argb: '000000' } },
            };
            cell.fill = {
                type: "pattern",
                pattern: "solid",
                fgColor: { argb: "FFFF00" }, // Fondo amarillo
            };
        });

        // Agregar datos a la hoja
        result.recordset.forEach((row) => {
            const newRow = worksheet.addRow(row);
            newRow.eachCell((cell) => {
                cell.border = {
                    top: { style: 'thin', color: { argb: '000000' } },
                    left: { style: 'thin', color: { argb: '000000' } },
                    bottom: { style: 'thin', color: { argb: '000000' } },
                    right: { style: 'thin', color: { argb: '000000' } },
                };
            });
        });

        // Guardar el archivo
        const filePath = "./datos.xlsx";
        await workbook.xlsx.writeFile(filePath);
        console.log(`Archivo Excel guardado en: ${filePath}`);
        
    } catch (error) {
        console.error("Error exportando el excel",error);
    } finally {
        await sql.close();
    }
}

async function fetchData() {
    try {
        // Conectar a SQL Server
        await sql.connect(config);
        console.log("Conectado a SQL Server");

        // Consulta con JOIN entre las tablas Brand y Beer
        const query = `
            SELECT Beer.BeerID, Beer.Name AS BeerName, Brand.Name AS BrandName 
            FROM Beer 
            INNER JOIN Brand ON Beer.BrandID = Brand.BrandID
        `;

        const result = await sql.query(query);
        return result.recordset;
    } catch (err) {
        console.error("Error en la conexión o consulta:", err);
        return [];
    } finally {
        sql.close();
    }
}

// 🔹 Función para exportar a Excel
async function exportToExcel2(data) {
    if (data.length === 0) {
        console.log("No hay datos para exportar.");
        return;
    }

    // Crear un nuevo libro de Excel
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Beers');

    // Agregar encabezados
    const columns = Object.keys(data[0]);
        worksheet.columns = columns.map((col) => ({ header: col, key: col, width: 20 }));


    // Aplicar formato al encabezado
    headerRow = worksheet.getRow(1).eachCell((cell) => {
        cell.font = { bold: true };
        cell.border = {
            top: { style: 'thin', color: { argb: '000000' } }, // Negro
            left: { style: 'thin', color: { argb: '000000' } },
            bottom: { style: 'thin', color: { argb: '000000' } },
            right: { style: 'thin', color: { argb: '000000' } },
        };
        cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFF00" }, // Fondo amarillo
        };
    });

    // Agregar datos a la hoja
    data.forEach((row) => {
        const newRow = worksheet.addRow(row);
        newRow.eachCell((cell) => {
            cell.border = {
                top: { style: 'thin', color: { argb: '000000' } },
                left: { style: 'thin', color: { argb: '000000' } },
                bottom: { style: 'thin', color: { argb: '000000' } },
                right: { style: 'thin', color: { argb: '000000' } },
            };
        });
    });

    // Guardar archivo
    await workbook.xlsx.writeFile('BeersList.xlsx');
    console.log("Archivo 'BeersList.xlsx' generado correctamente.");
}

// 🔹 Ejecutar el proceso
async function main() {
    const data = await fetchData();
    await exportToExcel2(data);
}

main();

// exportToExcel();