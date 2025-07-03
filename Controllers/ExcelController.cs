using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using ClosedXML;
using ToolsCTC.Models;
using ToolsCTC.Services;
using ClosedXML.Excel;

namespace ToolsCTC.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {

        private readonly ExcelProcessor _processor = new();

        [HttpPost("checkctc")]
        public async Task<IActionResult> CheckCTC([FromForm] IFormFile fileCTC, [FromForm] IFormFile fileQAS)
        {
            if (fileCTC == null || fileQAS == null)
                return BadRequest("Thiếu file");

            try
            {
                using var memoryCTC = new MemoryStream();
                await fileCTC.CopyToAsync(memoryCTC);
                memoryCTC.Position = 0;

                using var memoryQAS = new MemoryStream();
                await fileQAS.CopyToAsync(memoryQAS);
                memoryQAS.Position = 0;

                var result = _processor.CheckCTC(memoryCTC, fileCTC.FileName, memoryQAS, fileQAS.FileName);

                return Ok(result);
            }
            catch (Exception ex)
            {
                return StatusCode(500, "Lỗi xử lý file: " + ex.Message);
            }
        }

        [HttpPost("checkpm")]
        public async Task<IActionResult> CheckPM([FromForm] IFormFile fileCTC, [FromForm] IFormFile fileQAS)
        {
            if (fileCTC == null || fileQAS == null)
                return BadRequest("Thiếu file");

            try
            {
                using var memoryCTC = new MemoryStream();
                await fileCTC.CopyToAsync(memoryCTC);
                memoryCTC.Position = 0;

                using var memoryQAS = new MemoryStream();
                await fileQAS.CopyToAsync(memoryQAS);
                memoryQAS.Position = 0;

                var result = _processor.CheckPM(memoryCTC, fileCTC.FileName, memoryQAS, fileQAS.FileName);

                return Ok(result);
            }
            catch (Exception ex)
            {
                return StatusCode(500, "Lỗi xử lý file: " + ex.Message);
            }
        }

        [HttpPost("checkdup")]
        public async Task<IActionResult> CheckDupAsync([FromForm] IFormFile file)
        {
            if (file == null || file.Length == 0)
                return BadRequest("Không có file.");

            try
            {
                using var memoryCTC = new MemoryStream();
                await file.CopyToAsync(memoryCTC);
                memoryCTC.Position = 0;

                var rs = _processor.CheckDup(memoryCTC, file.FileName);

                return Ok(rs);
            }
            catch (Exception ex)
            {
                return StatusCode(500, "Lỗi xử lý file: " + ex.Message);
            }
        }

    }
}
