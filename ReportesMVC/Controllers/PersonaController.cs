using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using ReportesMVC.Models;
using System.Drawing;

namespace ReportesMVC.Controllers
{
    public class PersonaController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        public string generarReporte()
        {
            using (MemoryStream ms = new MemoryStream())
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelPackage ep = new ExcelPackage ();
                
                ep.Workbook.Worksheets.Add("Hoja de Prueba");

                ExcelWorksheet ew1 = ep.Workbook.Worksheets[0];
                Reportes.Excel.TituloHorizontal(ew1);

                ep.SaveAs(ms);
                byte[] buffer = ms.ToArray();
                return Convert.ToBase64String(buffer);
            }
        }

        public async Task<List<Persona>> listaPersonas()
        {
            BDReportesContext context = new BDReportesContext();
            return await context.Personas.ToListAsync();
        }
    }
}
