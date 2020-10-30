using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using exceldownload.Models;
using exceldownload.ViewModels;
using OfficeOpenXml;
using Microsoft.AspNetCore.Http;

namespace exceldownload.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private UserContext _context;
        public HomeController(ILogger<HomeController> logger, UserContext _context)
        {
            _logger = logger;
            this._context = _context;
        }

        public IActionResult Index()
        {
            List<UserViewModel> list = _context.User.Select(x => new UserViewModel
            {
                Id = x.Id,
                Name = x.Name,
                Lastname = x.Lastname
            }).ToList();

            return View(list);
        }
        public void ExportToExcel()
        {
            List<UserViewModel> list = _context.User.Select(x => new UserViewModel
            {
                Id = x.Id,
                Name = x.Name,
                Lastname = x.Lastname
            }).ToList();

            ExcelPackage pck = new ExcelPackage();
            ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Report");

            ws.Cells["A1"].Value = "Id";
            ws.Cells["B1"].Value = "Name";
            ws.Cells["C1"].Value = "LastName";
            int rowStart = 2;
            foreach (var item in list)
            {
                ws.Cells[string.Format("A{0}", rowStart)].Value = item.Id;
                ws.Cells[string.Format("B{0}", rowStart)].Value = item.Name;
                ws.Cells[string.Format("C{0}", rowStart)].Value = item.Lastname;
                rowStart++;
            }

            ws.Cells["A:AZ"].AutoFitColumns();
            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.Headers.Add("content-disposition", "attachment; filename=ExcelReport.xlsx");
            Response.Body.WriteAsync(pck.GetAsByteArray());
            Response.Body.Flush();


        }
        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
