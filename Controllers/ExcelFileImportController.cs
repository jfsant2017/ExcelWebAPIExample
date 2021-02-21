using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.IO;
using OfficeOpenXml;
using System.Text;

namespace ExcelWebApiExample.Controllers
{
    [Route("v1/postfile")]
    public class ExcelFileImportController: ControllerBase
    {
        [HttpPost("upload")]
        public async Task<ActionResult<string>> ImportFile([FromForm] IFormFile excelFile)
        {
            StringBuilder readData = new StringBuilder();
            try
            {
                if (excelFile.Length > 0)
                {
                    MemoryStream ms = new MemoryStream();
                    await excelFile.CopyToAsync(ms);

                    using (var package = new ExcelPackage(ms))
                    {
                        var firstSheet = package.Workbook.Worksheets["Simple interest"];
                        readData.AppendLine(firstSheet.Name);
                        readData.AppendLine(string.Format("Type: {0}", firstSheet.Cells[1, 1].Value.ToString()));
                        readData.AppendLine(string.Format("Client name: {0}", firstSheet.Cells[2, 2].Value.ToString()));
                        readData.AppendLine(string.Format("Principal amount value: {0}", firstSheet.Cells[3, 2].Value.ToString()));
                        readData.AppendLine(string.Format("{0}: {1}", firstSheet.Cells[4, 1].Value.ToString(), firstSheet.Cells[4, 2].Value.ToString()));
                        readData.AppendLine(string.Format("{0}: {1}", firstSheet.Cells[5, 1].Value.ToString(), firstSheet.Cells[5, 2].Value.ToString()));
                        readData.AppendLine(string.Format("Interest value: {0}", firstSheet.Cells[6, 2].Value.ToString()));
                    }
                }
                return Ok(new { message = readData.ToString() });
            }
            catch (System.Exception ex)
            {
                return BadRequest(new { message = ex.Message });
            }

        }
    }
}