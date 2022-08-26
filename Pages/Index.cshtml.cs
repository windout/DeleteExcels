using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Text.RegularExpressions;

namespace DeleteExcels.Pages
{
    public class IndexModel : PageModel
    {
        [Required(ErrorMessage = "Chọn một file")]
        [DataType(DataType.Upload)]
        [Display(Name = "Chọn file upload")]
        [BindProperty]
        public IFormFile? FileUpload { get; set; }
        public void OnGet()
        {
            RedirectToAction("Index");
        }
        
        public async Task<IActionResult> OnPostAsync()
        {
            if (ModelState.IsValid)
            {
                var info = new FileInfo(FileUpload.FileName);
                if (info.Extension == ".xlsx")
                {
                    using (var package = new ExcelPackage(FileUpload.OpenReadStream()))
                    {
                        var t = package.Workbook.Worksheets.Where(x => new Regex(@"thẻ\skho\sbếp").IsMatch(x.Name.ToLower())).ToList();
                        t.AddRange(package.Workbook.Worksheets.Where(x => new Regex(@"thẻ\skho\sbar").IsMatch(x.Name.ToLower())).ToList());
                        int count = 0;
                        if (DateTime.Now.Month == 1)
                        {
                            foreach (var item in t)
                            {
                                try
                                {
                                    if (int.Parse(item.Name.Substring(item.Name.IndexOf('.') + 1, 2)) != DateTime.Now.Month && 
                                        int.Parse(item.Name.Substring(item.Name.IndexOf('.') + 1, 2)) != 12)
                                    {
                                        package.Workbook.Worksheets.Delete(item);
                                        count++;
                                    }
                                }
                                catch
                                {
                                    package.Workbook.Worksheets.Delete(item);
                                    count++;
                                }
                            }
                        }
                        else
                        {
                            foreach (var item in t)
                            {
                                try
                                {
                                    if (int.Parse(item.Name.Substring(item.Name.IndexOf('.') + 1, 2)) != DateTime.Now.Month &&
                                        int.Parse(item.Name.Substring(item.Name.IndexOf('.') + 1, 2)) != DateTime.Now.Month - 1)
                                    {
                                        package.Workbook.Worksheets.Delete(item);
                                        count++;
                                    }
                                }
                                catch
                                {
                                    package.Workbook.Worksheets.Delete(item);
                                    count++;
                                }
                            }
                        }
                        if (count > 0)
                        {
                            var file = new MemoryStream();
                            package.SaveAs(file);
                            return File(file.GetBuffer(),FileUpload.ContentType,FileUpload.FileName);
                        }
                        else ModelState.AddModelError("", "ko có gì để xoá");
                    }
                }
                else ModelState.AddModelError("","Sai định dạng file Excel");
            }
            return null;
        } 
        
    }
}