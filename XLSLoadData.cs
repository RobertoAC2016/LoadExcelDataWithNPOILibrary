using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Data;
using System.Net.Http.Headers;

namespace LoadExcelDataWithNPOILibrary
{
    public class XLSLoadData
    {
        public XLSLoadData()
        { }
        public DataTable getExcelData()
        {
            //usaremos file stream para leer un archivo desde local
            String fpath = @"C:\Users\Admin\Downloads\horarios-oficiales.xlsx";
            FileStream excelStream = new FileStream(fpath, FileMode.Open);
            //Creamos un datatable para almacenar los datos del excel
            DataTable table = new DataTable();
            //con esta sentencia asignamos el stream a la variable book y cerramos o destruimos el stream
            var book = new XSSFWorkbook(excelStream);
            excelStream.Close();

            //Ahora obtenemos la hoja del libro de excel, el header, el numero de columnas y registros filas
            var sheet = book.GetSheetAt(0);//0 es la primera hoja del libro por numero o indice
            var headerRow = sheet.GetRow(0); //0 es la primera fila de la celda
            var cellCount = headerRow.LastCellNum;//Obtiene el maximo de columnas
            var rowCount = sheet.LastRowNum;//Obtiene el maximo de filas

            //Aqui debemos decidir si la hoja q se esta leyendo tiene headers o cabeceras, sino lo omitimos
            //Leemos las columnas de la primera fila
            for (int i = headerRow.FirstCellNum; i < cellCount; i++)
            {
                var column = new DataColumn(headerRow.GetCell(i).StringCellValue);
                table.Columns.Add(column);
            }

            //Ahora leeremos el resto de las filas
            for (int i = sheet.FirstRowNum + 1; i < rowCount; i++)
            {
                var row = sheet.GetRow(i);
                var dataRow = table.NewRow();
                if (row != null)
                {
                    for (int j = row.FirstCellNum; j < cellCount; j++)
                    {
                        if (row.GetCell(j) != null)
                        {
                            //aqui creamos un funcion q no regrese solo el contenido sino q
                            //trate de interpretarlo ya q puede contener formulas
                            dataRow[j] = GetCellValue(row.GetCell(j));
                        }
                    }
                }
                //Despues de agregar el valor de las celdas a la fila las agregamos a la lista de
                //filas de la tabla
                table.Rows.Add(dataRow);
            }
            return table;
        }
        private object GetCellValue(ICell cell)
        {
            if (cell == null)
            {
                return String.Empty;
            }
            switch (cell?.CellType)
            {
                case CellType.Blank:
                    return String.Empty;
                case CellType.Boolean:
                    return cell.BooleanCellValue.ToString();
                case CellType.Error:
                    return cell.ErrorCellValue.ToString();
                case CellType.Numeric:
                case CellType.Unknown:
                default:
                    return cell?.ToString()??"";
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Formula:
                    try
                    {
                        var e = new HSSFFormulaEvaluator(cell.Sheet.Workbook);
                        e.EvaluateInCell(cell);
                        return cell.ToString()??"";
                    }
                    catch
                    {
                        return cell.StringCellValue;
                    }
            }
        }
    }
}