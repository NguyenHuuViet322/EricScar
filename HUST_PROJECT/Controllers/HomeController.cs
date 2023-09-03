using ClosedXML.Excel;
using DocumentFormat.OpenXml.EMMA;
using HUST_PROJECT.Models;
using IronOcr;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using static IronOcr.OcrResult;
using IHostingEnvironment = Microsoft.AspNetCore.Hosting.IHostingEnvironment;
using Info = HUST_PROJECT.Models.Info;

namespace HUST_PROJECT.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IHostingEnvironment _environment;

        public HomeController(ILogger<HomeController> logger, IHostingEnvironment environment)
        {
            _logger = logger;
            _environment = environment;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Tool()
        {
            return View();
        }

        public string processing(string str)
        {
            if (str.Length >= 2)
            {
                if (str[str.Length - 1] >= '0' && str[str.Length - 1] <= '9' && str[str.Length - 1] != '0')
                {
                    if (str[str.Length - 2] >= '0' && str[str.Length - 2] <= '9')
                    {
                        str = str.Insert(str.Length - 1, ".");
                    }
                }
            }
            return str;
        }

        public IActionResult Result(IFormFile postedFile) {
            List<Info> lstInfo = new List<Info>();

            IronTesseract ocr = new IronTesseract();
            ocr.Language = OcrLanguage.VietnameseBest;
            ocr.Configuration.TesseractVersion = TesseractVersion.Tesseract5;
            
            using (OcrInput input = new OcrInput())
            {
                var stream = new MemoryStream();
                postedFile.CopyTo(stream);
                var img = System.Drawing.Image.FromStream(stream);
                // Add multiple images
                input.AddImage(img);
                input.Scale(150);
                input.Sharpen();
                OcrResult result = ocr.Read(input);
                var text = result.Text;
                String[] spearator = { " ", "\r\n" };
                var tmp = text.Split(spearator, StringSplitOptions.None);
                var stt = 0;

                var n = tmp.Length;
                var i = 0;
                while (i < n)
                {
                    Info info = new Info();



                    try
                    {
                        if (stt == 0)
                        {
                            while (!tmp[i].ToString().Any(char.IsDigit))
                            {
                                i++;
                            }
                        }
                        else
                        {
                            while (tmp[i].ToString().Any(char.IsDigit) || tmp[i].ToString().Length <= 2)
                            {
                                i++;
                            }
                        }
                        if (stt == 0)
                        {
                            i++;
                            stt = 1;
                        }

                        
                        string name = "";
                        while (!tmp[i].ToString().Any(char.IsDigit))
                        {
                            name += tmp[i].ToString();
                            name += " ";
                            i++;
                        }

                        info.HoTen = name;
                        info.Toan = processing(tmp[i].ToString());
                        i++;
                        info.Ly = processing(tmp[i].ToString());
                        i++;
                        info.Hoa = processing(tmp[i].ToString());
                        i++;
                        info.TB = processing(tmp[i].ToString());
                        i++;
                        info.XepLoai = processing(tmp[i].ToString());
                        i++;
                        lstInfo.Add(info);
                    } catch(Exception ex) { }

                    
                }
                

            }
            ViewData["url"] = ExportXlsx(lstInfo);
            return View(lstInfo); 
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        public FileResult DownloadFile(string filename)
        {
            var webRootPath = _environment.WebRootPath;

            byte[] file = System.IO.File.ReadAllBytes(filename);

            return File(file, "application/octet-stream", "BangDiem.xlsx");
        }
        public string ExportXlsx(List<Models.Info> list)
        {
            var webRootPath = _environment.WebRootPath;
            var stt = 1;
            var time = DateTime.Now.ToString("dd_mm_yyyy_hh_mm_ss") + ".xlsx";
            byte[] file;

            using (var workbook = new XLWorkbook())
            {
                
                int rownum = 2;

                var worksheet = workbook.Worksheets.Add("Bang Diem Sheet");

                worksheet.Cell("A1").Value = "STT";
                worksheet.Cell("B1").Value = "Họ và tên";
                worksheet.Cell("C1").Value = "Toán";
                worksheet.Cell("D1").Value = "Lý";
                worksheet.Cell("E1").Value = "Hóa";
                worksheet.Cell("F1").Value = "Trung bình";
                worksheet.Cell("G1").Value = "Xếp loại";

                foreach (var item in list)
                {
                    worksheet.Cell("A" + rownum).Value = stt;
                    worksheet.Cell("B" + rownum).Value = item.HoTen;
                    worksheet.Cell("C" + rownum).Value = item.Toan;
                    worksheet.Cell("D" + rownum).Value = item.Ly;
                    worksheet.Cell("E" + rownum).Value = item.Hoa;
                    worksheet.Cell("F" + rownum).Value = item.TB;
                    worksheet.Cell("G" + rownum).Value = item.XepLoai;
                    rownum++;
                    stt++;
                }

                var workbookBytes = new byte[0];
                using (var ms = new MemoryStream())
                {
                    workbook.SaveAs(ms);
                    workbookBytes = ms.ToArray();
                }

                file = workbookBytes;
            }
            var path = Path.Combine(webRootPath, @"xlsx/", DateTime.Now.ToString("dd_mm_yyyy_hh_mm_ss") + ".xlsx");
            using (var stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None))              //Lỗi dòng này anh ơi
            {
                stream.Write(file, 0, file.Length);
            }

                //using (ExcelEngine excelEngine = new ExcelEngine())
                //         {
                //	IApplication application = excelEngine.Excel;
                //	application.DefaultVersion = ExcelVersion.Xlsx;

                //	IWorkbook workbook = application.Workbooks.Create(1);
                //	IWorksheet worksheet = workbook.Worksheets[0];



                //	worksheet.Range["A" + rownum].Text = "STT";
                //	worksheet.Range["B" + rownum].Text = "Url được chuyển đổi";
                //	worksheet.Range["C" + rownum].Text = "Thời điểm chuyển đổi";
                //	worksheet.Range["D" + rownum].Text = "Trạng thái";



                //             workbook.SaveAs(new FileStream(Path.Combine(webRootPath, @"/xlsx/", time), FileMode.Create, FileAccess.Write,FileShare.None));
                //}


                return path;
        }
    }
}
