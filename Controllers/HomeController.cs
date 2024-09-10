using System.Diagnostics;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using xcelHaveImage.Models;

namespace xcelHaveImage.Controllers;

public class HomeController : Controller
{
  private readonly ILogger<HomeController> _logger;

  public HomeController(ILogger<HomeController> logger)
  {
    _logger = logger;
  }

  public IActionResult Index()
  {
    return View();
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

  public IActionResult Upload()
  {
    return View(); // This will look for "Upload.cshtml" in the correct folder by default.
  }

  [HttpPost]
  public async Task<IActionResult> Upload(IFormFile excelFile)
  {
    if (excelFile == null || excelFile.Length == 0)
    {
      ViewBag.Message = "Please select an Excel file.";
      return View();
    }
    string imagesPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "images");
    if (!Directory.Exists(imagesPath))
    {
      Directory.CreateDirectory(imagesPath);
    }
    var records = new List<ExcelRecord>();

    using (var stream = new MemoryStream())
    {
      await excelFile.CopyToAsync(stream);
      using (var package = new ExcelPackage(stream))
      {
        var worksheet = package.Workbook.Worksheets[0]; // Read the first worksheet
        var rowCount = worksheet.Dimension.Rows;

        for (int row = 2; row <= rowCount; row++) // Assuming the first row has headers
        {
          var record = new ExcelRecord
          {
            Id = worksheet.Cells[row, 1].Text,
            Name = worksheet.Cells[row, 2].Text
          };

          // Extract images from the worksheet
          foreach (var drawing in worksheet.Drawings)
          {
            if (drawing is ExcelPicture picture)
            {
              // Check if the image is positioned near the relevant row and column
              if (picture.From.Row + 1 == row && picture.From.Column == 2) // Adjust column index if necessary
              {
                // Save the image to the server
                string imageName = $"image_{row - 1}.png";
                string imagePath = Path.Combine(imagesPath, imageName);

                using (var imageStream = new FileStream(imagePath, FileMode.Create))
                {
                  imageStream.Write(picture.Image.ImageBytes, 0, picture.Image.ImageBytes.Length);
                }

                record.Description = $"/images/{imageName}"; // Set the path for the view
              }
            }
          }

          records.Add(record);
        }
      }
    }

    return View("DisplayRecords", records); // Pass the records to a view named "DisplayRecords"
  }
}


public class ExcelRecord
{
  public string Id { get; set; }
  public string Name { get; set; }
  public string Description { get; set; }
}
