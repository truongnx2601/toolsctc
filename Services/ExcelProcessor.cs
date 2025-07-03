namespace ToolsCTC.Services
{
    using ClosedXML.Excel;
    using System.Globalization;
    using Models;
    using System.Text.RegularExpressions;
    using System.Text;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.XSSF.UserModel;

    public class ExcelProcessor
    {
        public List<ImportsInjection> CheckCTC(Stream streamCTC, string nameCTC, Stream streamQAS, string nameQAS)
        {
            var lstimp = ReadCTC(streamCTC, nameCTC);
            var inj = ReadQAS(streamQAS, nameQAS);

            // Lọc ra các ca không khớp giữa dữ liệu pm và ctc
            var lstCheck = inj
                .Where(i => !lstimp.Any(x =>
                    NormalizeVietnamese(x.FullName).Equals(NormalizeVietnamese(i.HoTen), StringComparison.OrdinalIgnoreCase)
                    && x.Birthday.Date == i.NgaySinh.Date
                    && x.VaccineDate.Date == i.NgayTiem.Date))
                .ToList();

            return lstCheck;
        }

        private IWorkbook GetWorkbook(Stream stream, string fileName)
        {
            if (fileName.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
                return new HSSFWorkbook(stream);
            if (fileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                return new XSSFWorkbook(stream);
            throw new Exception("Định dạng file không hỗ trợ");
        }

        private List<Imports> ReadCTC(Stream stream, string fileName)
        {
            var result = new List<Imports>();
            var wb = GetWorkbook(stream, fileName);
            var sheet = wb.GetSheetAt(0);

            for (int i = 2; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                if (row == null) continue;

                try
                {
                    result.Add(new Imports
                    {
                        FacID = row.GetCell(1)?.ToString() ?? "",
                        FullName = row.GetCell(2)?.ToString() ?? "",
                        Birthday = DateTime.ParseExact(row.GetCell(3)?.ToString() ?? "", "dd/MM/yyyy", CultureInfo.InvariantCulture),
                        Gen = row.GetCell(4)?.ToString() ?? "",
                        Address = row.GetCell(5)?.ToString() ?? "",
                        VaccineName = row.GetCell(6)?.ToString() ?? "",
                        VaccineDate = DateTime.ParseExact(row.GetCell(7)?.ToString() ?? "", "dd/MM/yyyy", CultureInfo.InvariantCulture),
                        KHT = row.GetCell(8)?.ToString() ?? "",
                        EmpName = row.GetCell(9)?.ToString() ?? ""
                    });
                }
                catch { }
            }

            return result;
        }

        private List<ImportsInjection> ReadQAS(Stream stream, string fileName)
        {
            var result = new List<ImportsInjection>();
            var wb = GetWorkbook(stream, fileName);
            var sheet = wb.GetSheetAt(0);

            for (int i = 6; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                if (row == null || string.IsNullOrWhiteSpace(row.GetCell(0)?.ToString())) continue;

                try
                {
                    result.Add(new ImportsInjection
                    {
                        NgayTiem = DateTime.ParseExact(row.GetCell(0)?.ToString() ?? "", "yyyy-MM-dd", CultureInfo.InvariantCulture),
                        STT = row.GetCell(1)?.ToString() ?? "",
                        SoPCD = row.GetCell(2)?.ToString() ?? "",
                        HoTen = row.GetCell(3)?.ToString() ?? "",
                        MaKH = row.GetCell(4)?.ToString() ?? "",
                        NgaySinh = DateTime.ParseExact(row.GetCell(5)?.ToString() ?? "", "dd/MM/yyyy", CultureInfo.InvariantCulture),
                        DiaChi = row.GetCell(6)?.ToString() ?? "",
                        NguoiLH = row.GetCell(7)?.ToString() ?? "",
                        SDT = row.GetCell(8)?.ToString() ?? "",
                        TenVaccine = row.GetCell(11)?.ToString() ?? ""
                    });
                }
                catch { }
            }

            return result;
        }

        public List<Imports> CheckPM(Stream streamCTC, string nameCTC, Stream streamQAS, string nameQAS)
        {
            var lstimp = ReadCTC(streamCTC, nameCTC);
            var inj = ReadQAS(streamQAS, nameQAS);

            // Kiểm tra người có trong kế hoạch mà không xuất hiện trong thực tế
            var lstCheck = lstimp
                .Where(i => !inj.Any(x =>
                    NormalizeVietnamese(x.HoTen).Equals(NormalizeVietnamese(i.FullName), StringComparison.OrdinalIgnoreCase)
                    && x.NgaySinh.Date == i.Birthday.Date
                    && x.NgayTiem.Date == i.VaccineDate.Date))
                .ToList();

            return lstCheck;
        }

        public List<Imports> CheckDup(Stream streamCTC, string nameCTC)
        {
            var list = ReadCTC(streamCTC, nameCTC);

            // Kiểm tra người có bị trùng không
            var lstCheck = list
                    .GroupBy(x => new
                    {
                        HoTen = NormalizeVietnamese(x.FullName).Trim().ToLower(),
                        NgaySinh = x.Birthday.Date,
                        NgayTiem = x.VaccineDate.Date,
                        TenVaccine = x.VaccineName?.Replace(" ", "").ToLower(),
                        DiaChi = x.Address
                    })
                    .Where(g => g.Count() > 1)
                    .SelectMany(g => g)
                    .ToList();

            return lstCheck;
        }

        public static string NormalizeVaccineName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "";

            // 1. Loại bỏ các từ chứa số và các ký hiệu như "0.5", "1ml", "mũi 2"
            // Cách: giữ lại chỉ chữ cái và khoảng trắng
            var lettersOnly = Regex.Replace(name, @"[^a-zA-Z\s]", "");  // bỏ mọi thứ không phải chữ

            // 2. Loại bỏ từ "mui" nếu còn sót
            //lettersOnly = Regex.Replace(lettersOnly, @"\bmui\b", "", RegexOptions.IgnoreCase);

            // 3. Chuẩn hóa khoảng trắng
            return Regex.Replace(lettersOnly, @"\s+", " ").Trim();
        }

        public static string NormalizeVietnamese(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return "";

            input = input.ToLowerInvariant().Trim();

            // Bỏ dấu tiếng Việt
            string normalized = input.Normalize(NormalizationForm.FormD);
            var sb = new StringBuilder();

            foreach (char c in normalized)
            {
                var unicodeCategory = CharUnicodeInfo.GetUnicodeCategory(c);
                if (unicodeCategory != UnicodeCategory.NonSpacingMark)
                {
                    sb.Append(c);
                }
            }

            // Ghép lại, xóa khoảng trắng dư
            var cleaned = Regex.Replace(sb.ToString(), @"\s+", " ");
            return cleaned.Normalize(NormalizationForm.FormC);
        }

    }
}
