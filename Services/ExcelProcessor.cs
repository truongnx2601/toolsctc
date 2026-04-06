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

    var pmKeys = new HashSet<string>(
        lstimp.Select(x => BuildKey(x.FullName, x.Birthday, x.VaccineDate))
    );

    var result = inj
        .Where(i => !pmKeys.Contains(
            BuildKey(i.HoTen, i.NgaySinh, i.NgayTiem)))
        .ToList();

    return result;
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
                        FacID = GetCellString(row.GetCell(1)) ?? "",
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

        private List<ImportsInjection> ReadPM(Stream stream, string fileName)
        {
            var result = new List<ImportsInjection>();
            var wb = GetWorkbook(stream, fileName);
            var sheet = wb.GetSheetAt(0);

            for (int i = 6; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                if (row == null || string.IsNullOrWhiteSpace(GetCellString(row.GetCell(0)))) continue;

                try
                {
                    result.Add(new ImportsInjection
                    {
                        MaTC = GetCellString(row.GetCell(1)) ?? "",
                        HoTen = row.GetCell(3)?.ToString() ?? "",
                    });
                }
                catch { }
            }

            return result;
        }

        public List<ImportsInjection> CheckPM(Stream streamCTC, string nameCTC, Stream streamQAS, string nameQAS)
        {
            var lstimp = ReadCTC(streamCTC, nameCTC);
            var inj = ReadQAS(streamQAS, nameQAS);

            var injKeys = new HashSet<string>(
                inj.Select(x => BuildKey(x.HoTen, x.NgaySinh, x.NgayTiem))
            );

            var lstCheck = lstimp
                .Where(i => !injKeys.Contains(
                    BuildKey(i.FullName, i.Birthday, i.VaccineDate)))
                .ToList();

            var result = lstCheck.Select(x => new ImportsInjection
            {
                MaTC = x.FacID,
                HoTen = x.FullName,
                NgaySinh = x.Birthday,
                NgayTiem = x.VaccineDate,
                TenVaccine = x.VaccineName,
                DiaChi = x.Address,
                SDT = "",
                NguoiLH = ""
            }).ToList();

            return result;
        }

        public List<ImportsInjection> CheckDup(Stream streamCTC, string nameCTC)
        {
            var list = ReadCTC(streamCTC, nameCTC);

            var lstCheck = list
                .GroupBy(x => new
                {
                    maTC = x.FacID,
                    hoTen = NormalizeVietnamese(x.FullName).Trim().ToLower(),
                    ngaySinh = x.Birthday.Date,
                    ngayTiem = x.VaccineDate.Date,
                    tenVaccine = x.VaccineName?.Replace(" ", "").ToLower(),
                    diaChi = x.Address?.Trim().ToLower()
                })
                .Where(g => g.Count() > 1)
                .SelectMany(g => g)
                .ToList();

            // Mapping sang dạng frontend cần
            var result = lstCheck.Select(x => new ImportsInjection
            {
                MaTC = x.FacID,
                HoTen = x.FullName,
                NgaySinh = x.Birthday,
                NgayTiem = x.VaccineDate,
                TenVaccine = x.VaccineName,
                DiaChi = x.Address,
                SDT = "", // nếu có thì truyền từ x
                NguoiLH = "" // nếu có thì truyền từ x
            }).ToList();

            return result;
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

        public List<ImportsInjection> CheckErr(Stream streamCTC, string nameCTC, Stream streamQAS, string nameQAS)
        {
            var ctcList = ReadCTC(streamCTC, nameCTC);
            var pmList = ReadPM(streamQAS, nameQAS);

            var ctcGroup = ctcList
                .Where(x => !string.IsNullOrWhiteSpace(x.FacID))
                .GroupBy(x => BuildKeyByMaTC(x.FacID))
                .ToDictionary(g => g.Key, g => g.Count());

            var pmGroup = pmList
                .Where(x => !string.IsNullOrWhiteSpace(x.MaTC))
                .GroupBy(x => BuildKeyByMaTC(x.MaTC))
                .ToDictionary(g => g.Key, g => g.Count());

            var allKeys = new HashSet<string>(ctcGroup.Keys);
            allKeys.UnionWith(pmGroup.Keys);

            var result = new List<ImportsInjection>();

            foreach (var key in allKeys)
            {
                if (string.IsNullOrWhiteSpace(key)) continue;

                int countCTC = ctcGroup.ContainsKey(key) ? ctcGroup[key] : 0;
                int countPM = pmGroup.ContainsKey(key) ? pmGroup[key] : 0;

                if (countCTC == countPM) continue; // số lượng bằng nhau → không báo lỗi

                // Chọn sample để lấy tên (HoTen)
                string hoTen = pmList.FirstOrDefault(x => BuildKeyByMaTC(x.MaTC) == key)?.HoTen
                                ?? ctcList.FirstOrDefault(x => BuildKeyByMaTC(x.FacID) == key)?.FullName
                                ?? "";

                result.Add(new ImportsInjection
                {
                    MaTC = key,
                    HoTen = hoTen,
                    CountCTC = countCTC,
                    CountQAS = countPM,
                    Status = countCTC > countPM ? "THIEU_PM" : "DU_PM",
                    Note = "Có sự khác biệt về số lượng mũi tiêm trên CTC/PM"
                });
            }

            return result;
        }

        private string BuildKey(string name, DateTime birth, DateTime injectDate)
        {
            return $"{NormalizeVietnamese(name)}|{birth:yyyyMMdd}|{injectDate:yyyyMMdd}";
        }

        private string BuildKeyByMaTC(string maTC)
        {
            if (string.IsNullOrWhiteSpace(maTC)) return "";

            return Regex.Replace(maTC, @"[^0-9]", "");
        }

        private readonly DataFormatter _formatter = new DataFormatter();
        private string GetCellString(ICell cell)
        {
            if (cell == null) return "";

            return _formatter.FormatCellValue(cell)?.Trim() ?? "";
        }

    }
}
