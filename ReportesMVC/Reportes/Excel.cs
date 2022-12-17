using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace ReportesMVC.Reportes
{
    public class Excel
    {
        public static void TituloHorizontal(ExcelWorksheet ew, string titulo = "Reportes Persona", int posFila = 1, int posInicioCol = 1, int posFinCol = 4, Color? fondo=null, Color? colorTexto = null)
        {
            using (ExcelRange range = ew.Cells[posFila, posInicioCol, posFila, posFinCol])
            {
                range.Merge = true;
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                range.Style.Font.Size = 20;
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                if(fondo == null)
                    range.Style.Fill.BackgroundColor.SetColor(Color.Teal);
                else
                    range.Style.Fill.BackgroundColor.SetColor((Color)fondo);
                if(colorTexto == null)
                    range.Style.Font.Color.SetColor(Color.White);
                else
                    range.Style.Font.Color.SetColor((Color)colorTexto);
            }
            ew.Cells[posFila, posInicioCol].Value = titulo;
        }
    }
}
