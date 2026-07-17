using Microsoft.AspNetCore.Mvc;
using MyExcelApp.Models;
using MyExcelApp.Helpers;
using OfficeOpenXml;
using System.IO;
using System.Linq;

namespace MyExcelApp.Controllers
{
    public class HomeController : Controller
    {
        private readonly IWebHostEnvironment _env;

        public HomeController(IWebHostEnvironment env)
        {
            _env = env;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        // ================================================================
        // متد بررسی اتصال به سرور
        // ================================================================
        [HttpGet]
        public IActionResult CheckConnection()
        {
            return Ok();
        }

        // ================================================================
        // صفحه اصلی (فرم)
        // ================================================================
        public IActionResult Index()
        {
            var now = DateTime.Now;
            var persianDayOfWeek = GetPersianDayOfWeek(now);
            
            ViewBag.CurrentDate = now.ToPersianDate();
            ViewBag.CurrentTime = now.ToString("HH:mm:ss");
            ViewBag.CurrentDayOfWeek = persianDayOfWeek;
            
            return View();
        }

        private string GetPersianDayOfWeek(DateTime date)
        {
            var days = new string[] { "یکشنبه", "دوشنبه", "سه‌شنبه", "چهارشنبه", "پنجشنبه", "جمعه", "شنبه" };
            return days[(int)date.DayOfWeek];
        }

        // ================================================================
        // ثبت اطلاعات
        // ================================================================
        [HttpPost]
        public IActionResult Submit(FormData data)
        {
            if (ModelState.IsValid)
            {
                var now = DateTime.Now;
                data.Date = now.ToPersianDate();
                data.Time = now.ToString("HH:mm:ss");
                data.DayOfWeek = GetPersianDayOfWeek(now);
                
                if (data.MoldCount.HasValue && data.MoldWeight.HasValue)
                {
                    data.TotalMoldWeight = data.MoldCount.Value * data.MoldWeight.Value;
                }
                
                data.TotalInputWeight = CalculateTotalMaterials(data);
                
                string filePath = SaveToExcel(data);
                ViewBag.Message = $"✅ ثبت شد. شماره سند: {Path.GetFileNameWithoutExtension(filePath)}";
                
                ViewBag.CurrentDate = now.ToPersianDate();
                ViewBag.CurrentTime = now.ToString("HH:mm:ss");
                ViewBag.CurrentDayOfWeek = GetPersianDayOfWeek(now);
                
                return View("Index", data);
            }
            
            var now2 = DateTime.Now;
            ViewBag.CurrentDate = now2.ToPersianDate();
            ViewBag.CurrentTime = now2.ToString("HH:mm:ss");
            ViewBag.CurrentDayOfWeek = GetPersianDayOfWeek(now2);
            
            return View("Index", data);
        }

        // ================================================================
        // محاسبه جمع وزن مواد اولیه
        // ================================================================
        private decimal CalculateTotalMaterials(FormData data)
        {
            decimal total = 0;
            
            total += data.CheeseFrozen ?? 0;
            total += data.CheeseDefrost ?? 0;
            total += data.Oil ?? 0;
            total += data.Starch ?? 0;
            total += data.Ricotta ?? 0;
            total += data.Phosphate ?? 0;
            total += data.Water ?? 0;
            total += data.Stabilizer ?? 0;
            total += data.Citrate ?? 0;
            total += data.Salt ?? 0;
            total += data.WasteReturn ?? 0;
            total += data.Essence ?? 0;
            total += data.CheeseGolpayegan ?? 0;
            total += data.Sorbate ?? 0;
            total += data.PowderCode24 ?? 0;
            total += data.PowderCodeF ?? 0;
            total += data.ColorAnnatto ?? 0;
            total += data.PlasticUnderLayer ?? 0;
            total += data.RawMaterial1 ?? 0;
            total += data.RawMaterial2 ?? 0;
            total += data.RawMaterial3 ?? 0;
            
            return total;
        }

        // ================================================================
        // ذخیره در فایل Excel
        // ================================================================
        private string SaveToExcel(FormData data)
        {
            string folderPath = Path.Combine(_env.WebRootPath, "ExcelFiles");
            Directory.CreateDirectory(folderPath);

            string persianDate = data.Date?.Replace("/", "") ?? DateTime.Now.ToPersianDate().Replace("/", "");
            string randomNum = new Random().Next(1000, 9999).ToString();
            string docNumber = $"{persianDate}_{randomNum}";
            string filePath = Path.Combine(folderPath, $"{docNumber}.xlsx");

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("گزارش تولید");

                worksheet.Column(1).Width = 30;
                worksheet.Column(2).Width = 25;
                worksheet.Column(3).Width = 25;

                int row = 1;

                // ===== اطلاعات کلی =====
                worksheet.Cells[row, 1].Value = "اطلاعات کلی";
                worksheet.Cells[row, 1, row, 3].Merge = true;
                worksheet.Cells[row, 1].Style.Font.Bold = true;
                worksheet.Cells[row, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                row++;

                worksheet.Cells[row, 1].Value = "تاریخ";
                worksheet.Cells[row, 2].Value = data.Date;
                row++;
                worksheet.Cells[row, 1].Value = "ساعت";
                worksheet.Cells[row, 2].Value = data.Time;
                row++;
                worksheet.Cells[row, 1].Value = "روز هفته";
                worksheet.Cells[row, 2].Value = data.DayOfWeek;
                row++;
                worksheet.Cells[row, 1].Value = "مسئول پخت";
                worksheet.Cells[row, 2].Value = data.ResponsibleCook;
                row++;
                worksheet.Cells[row, 1].Value = "کارشناس فنی تولید";
                worksheet.Cells[row, 2].Value = data.TechnicalExpert;
                row++;
                worksheet.Cells[row, 1].Value = "نوع محصول تولید";
                worksheet.Cells[row, 2].Value = data.ProductType;
                row++;
                worksheet.Cells[row, 1].Value = "تعداد پرسنل حاضر";
                worksheet.Cells[row, 2].Value = data.PersonnelCount;
                row++;

                // ===== مقدار مواد اولیه =====
                row++;
                worksheet.Cells[row, 1].Value = "مقدار مواد اولیه (به گرم)";
                worksheet.Cells[row, 1, row, 3].Merge = true;
                worksheet.Cells[row, 1].Style.Font.Bold = true;
                worksheet.Cells[row, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                row++;

                var materials = new Dictionary<string, decimal?>
                {
                    {"پنیر اولیه منجمد", data.CheeseFrozen},
                    {"پنیر اولیه درفراست", data.CheeseDefrost},
                    {"روغن", data.Oil},
                    {"نشاسته", data.Starch},
                    {"ریکوتا", data.Ricotta},
                    {"فسفات", data.Phosphate},
                    {"آب", data.Water},
                    {"استبلایزر", data.Stabilizer},
                    {"سیترات", data.Citrate},
                    {"نمک", data.Salt},
                    {"ضایعات برگشتی", data.WasteReturn},
                    {"اسانس", data.Essence},
                    {"پنیر گلپایگان", data.CheeseGolpayegan},
                    {"سوربات", data.Sorbate},
                    {"پودر کد 24 رخشان سپهر", data.PowderCode24},
                    {"پودر کد F", data.PowderCodeF},
                    {"رنگ آناتو", data.ColorAnnatto},
                    {"پلاستیک لایه زیر پخت", data.PlasticUnderLayer}
                };

                foreach (var item in materials)
                {
                    worksheet.Cells[row, 1].Value = item.Key;
                    worksheet.Cells[row, 2].Value = item.Value;
                    row++;
                }

                // مواد اولیه متفرقه با نام
                var rawMaterials = new List<(string Name, decimal? Value)>
                {
                    (data.RawMaterial1Name ?? "مواد اولیه 1", data.RawMaterial1),
                    (data.RawMaterial2Name ?? "مواد اولیه 2", data.RawMaterial2),
                    (data.RawMaterial3Name ?? "مواد اولیه 3", data.RawMaterial3)
                };

                foreach (var item in rawMaterials)
                {
                    worksheet.Cells[row, 1].Value = item.Name;
                    worksheet.Cells[row, 2].Value = item.Value;
                    row++;
                }

                // ===== اطلاعات قالب‌ها =====
                row++;
                worksheet.Cells[row, 1].Value = "اطلاعات قالب‌ها";
                worksheet.Cells[row, 1, row, 3].Merge = true;
                worksheet.Cells[row, 1].Style.Font.Bold = true;
                worksheet.Cells[row, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                row++;

                worksheet.Cells[row, 1].Value = "تعداد قالب‌ها";
                worksheet.Cells[row, 2].Value = data.MoldCount;
                row++;
                worksheet.Cells[row, 1].Value = "وزن هر قالب (گرم)";
                worksheet.Cells[row, 2].Value = data.MoldWeight;
                row++;
                worksheet.Cells[row, 1].Value = "وزن کل قالب‌ها (گرم)";
                worksheet.Cells[row, 2].Value = data.TotalMoldWeight;
                row++;

                // ===== اطلاعات وزن‌ها =====
                row++;
                worksheet.Cells[row, 1].Value = "اطلاعات وزن‌ها";
                worksheet.Cells[row, 1, row, 3].Merge = true;
                worksheet.Cells[row, 1].Style.Font.Bold = true;
                worksheet.Cells[row, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                row++;

                worksheet.Cells[row, 1].Value = "جمع وزن مواد ورودی (گرم)";
                worksheet.Cells[row, 2].Value = data.TotalInputWeight;
                row++;
                worksheet.Cells[row, 1].Value = "وزن واقعی کالای تولیدی (گرم)";
                worksheet.Cells[row, 2].Value = data.ActualProductWeight;
                row++;

                // ===== توضیحات نهایی =====
                row++;
                worksheet.Cells[row, 1].Value = "توضیحات نهایی";
                worksheet.Cells[row, 1, row, 3].Merge = true;
                worksheet.Cells[row, 1].Style.Font.Bold = true;
                worksheet.Cells[row, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                row++;

                worksheet.Cells[row, 1].Value = "توضیحات تولید";
                worksheet.Cells[row, 2].Value = data.ProductionDescription;
                row++;
                worksheet.Cells[row, 1].Value = "توضیحات مسئول پخت";
                worksheet.Cells[row, 2].Value = data.CookDescription;
                row++;
                worksheet.Cells[row, 1].Value = "توضیحات کارشناس فنی تولید";
                worksheet.Cells[row, 2].Value = data.ExpertDescription;

                worksheet.Cells.AutoFitColumns();
                package.SaveAs(new FileInfo(filePath));
            }

            return filePath;
        }

        // ================================================================
        // لیست فایل‌ها
        // ================================================================
        public IActionResult FileList()
        {
            string folderPath = Path.Combine(_env.WebRootPath, "ExcelFiles");
            Directory.CreateDirectory(folderPath);

            var files = Directory.GetFiles(folderPath, "*.xlsx")
                .Select(f => new FileInfo(f))
                .OrderByDescending(f => f.CreationTime)
                .Select(f =>
                {
                    var docNumber = Path.GetFileNameWithoutExtension(f.Name);
                    string persianDateStr = f.CreationTime.ToPersianDateTimeShort();
                    
                    string productType = "";
                    try
                    {
                        using (var package = new ExcelPackage(f))
                        {
                            var worksheet = package.Workbook.Worksheets[0];
                            for (int row = 1; row <= worksheet.Dimension.Rows; row++)
                            {
                                var cellValue = worksheet.Cells[row, 1].Value?.ToString();
                                if (cellValue == "نوع محصول تولید")
                                {
                                    productType = worksheet.Cells[row, 2].Value?.ToString() ?? "";
                                    break;
                                }
                            }
                        }
                    }
                    catch
                    {
                        productType = "";
                    }
                    
                    return new
                    {
                        FileName = f.Name,
                        FullPath = f.FullName,
                        DocNumber = docNumber,
                        ProductType = productType,
                        CreatedDate = persianDateStr,
                        FileSize = (f.Length / 1024.0).ToString("F2") + " KB",
                        FilePath = "/ExcelFiles/" + f.Name
                    };
                })
                .ToList();

            return View(files);
        }
    }
}