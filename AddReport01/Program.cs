using System.Globalization;
using ClosedXML.Excel;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;

static string? GetOption(string[] arguments, string optionName)
{
    var index = Array.FindIndex(arguments, value =>
        string.Equals(value, optionName, StringComparison.OrdinalIgnoreCase));
    if (index < 0)
        return null;
    if (index + 1 >= arguments.Length || arguments[index + 1].StartsWith("--", StringComparison.Ordinal))
        throw new ArgumentException($"ต้องระบุค่าหลัง {optionName}");
    return arguments[index + 1];
}

static DateTime ParseCommandDate(string value, string optionName) =>
    DateTime.TryParseExact(value, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var date)
        ? date.Date
        : throw new ArgumentException($"รูปแบบ {optionName} ไม่ถูกต้อง กรุณาใช้ yyyy-MM-dd");

static string NormalizeHeader(string value)
{
    var normalized = new string(value
        .Where(char.IsLetterOrDigit)
        .Select(char.ToLowerInvariant)
        .ToArray());
    return normalized is "visitdate" or "dateapp" ? "date" : normalized;
}

static bool TryReadDateCell(IXLCell cell, out DateTime date)
{
    if (cell.TryGetValue<DateTime>(out var dateTime))
    {
        date = dateTime.Date;
        return true;
    }

    var value = cell.GetString().Trim();
    var formats = new[] { "yyyy-MM-dd", "yyyyMMdd", "dd/MM/yyyy", "MM/dd/yyyy" };
    if (DateTime.TryParseExact(value, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out date) ||
        DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
    {
        date = date.Date;
        return true;
    }

    date = default;
    return false;
}

static void EnsureGroupedHeaders(IXLWorksheet worksheet)
{
    if (NormalizeHeader(worksheet.Cell(1, 2).GetString()) == "financial" &&
        NormalizeHeader(worksheet.Cell(2, 2).GetString()) == "opd")
    {
        return;
    }

    if (TryReadDateCell(worksheet.Cell(2, 1), out _))
        worksheet.Row(2).InsertRowsAbove(1);

    worksheet.Range(1, 1, 2, 24).Clear(XLClearOptions.Contents);
    worksheet.Range("B1:C1").Merge();
    worksheet.Range("D1:E1").Merge();
    worksheet.Range("F1:H1").Merge();
    worksheet.Range("I1:K1").Merge();
    worksheet.Range("L1:M1").Merge();
    worksheet.Range("N1:O1").Merge();
    worksheet.Range("P1:R1").Merge();
    worksheet.Range("S1:U1").Merge();
    worksheet.Range("V1:X1").Merge();

    worksheet.Cell("A1").Value = "Date";
    worksheet.Cell("B1").Value = "Financial";
    worksheet.Cell("D1").Value = "Doctor' Fee";
    worksheet.Cell("F1").Value = "LAB OPD";
    worksheet.Cell("I1").Value = "LAB IPD";
    worksheet.Cell("L1").Value = "X-Ray OPD";
    worksheet.Cell("N1").Value = "X-Ray IPD";
    worksheet.Cell("P1").Value = "OPD Patient";
    worksheet.Cell("S1").Value = "Old OPD";
    worksheet.Cell("V1").Value = "New OPD";

    worksheet.Cell("B2").Value = "OPD";
    worksheet.Cell("C2").Value = "IPD";
    worksheet.Cell("D2").Value = "OPD";
    worksheet.Cell("E2").Value = "IPD";
    worksheet.Cell("F2").Value = "Request";
    worksheet.Cell("G2").Value = "QTY";
    worksheet.Cell("H2").Value = "Amount";
    worksheet.Cell("I2").Value = "Request";
    worksheet.Cell("J2").Value = "QTY";
    worksheet.Cell("K2").Value = "Amount";
    worksheet.Cell("L2").Value = "Request";
    worksheet.Cell("M2").Value = "Amount";
    worksheet.Cell("N2").Value = "Request";
    worksheet.Cell("O2").Value = "Amount";
    worksheet.Cell("P2").Value = "Old";
    worksheet.Cell("Q2").Value = "New";
    worksheet.Cell("R2").Value = "Total";
    worksheet.Cell("S2").Value = "Male";
    worksheet.Cell("T2").Value = "Female";
    worksheet.Cell("U2").Value = "Child";
    worksheet.Cell("V2").Value = "Male";
    worksheet.Cell("W2").Value = "Female";
    worksheet.Cell("X2").Value = "Child";
}

static void FormatDailyReport(IXLWorksheet worksheet)
{
    const int lastColumn = 24;
    var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 2;
    var header = worksheet.Range(1, 1, 2, lastColumn);
    header.Style.Fill.BackgroundColor = XLColor.FromHtml("#163A5F");
    header.Style.Font.FontName = "Arial";
    header.Style.Font.FontSize = 10;
    header.Style.Font.Bold = true;
    header.Style.Font.FontColor = XLColor.White;
    header.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
    header.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
    header.Style.Alignment.WrapText = true;
    header.Style.Border.TopBorder = XLBorderStyleValues.None;
    header.Style.Border.LeftBorder = XLBorderStyleValues.None;
    header.Style.Border.RightBorder = XLBorderStyleValues.None;
    header.Style.Border.BottomBorder = XLBorderStyleValues.Medium;
    header.Style.Border.BottomBorderColor = XLColor.FromHtml("#2F75B5");
    worksheet.Range("B2:X2").Style.Fill.BackgroundColor = XLColor.FromHtml("#DCE6F1");
    worksheet.Range("B2:X2").Style.Font.FontColor = XLColor.FromHtml("#163A5F");
    worksheet.Row(1).Height = 32;
    worksheet.Row(2).Height = 24;
    worksheet.SheetView.FreezeRows(2);
    worksheet.SheetView.FreezeColumns(1);

    var currentGroup = string.Empty;
    for (var column = 1; column <= lastColumn; column++)
    {
        var groupHeader = worksheet.Cell(1, column).GetString().Trim();
        if (!string.IsNullOrEmpty(groupHeader))
            currentGroup = groupHeader;
        var subHeader = worksheet.Cell(2, column).GetString().Trim();
        var key = NormalizeHeader(currentGroup + subHeader);
        var columnRange = worksheet.Range(3, column, Math.Max(3, lastRow), column);

        if (NormalizeHeader(currentGroup) == "date")
        {
            worksheet.Column(column).Width = 14;
            columnRange.Style.NumberFormat.Format = "yyyy-mm-dd";
            columnRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        }
        else if (key.Contains("amount") || key.Contains("financial") || key.Contains("fee") || key.Contains("discount"))
        {
            worksheet.Column(column).Width = 18;
            columnRange.Style.NumberFormat.Format = "#,##0.00;[Red]-#,##0.00";
            columnRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        }
        else if (key.Contains("depname") || key.Contains("depengnam"))
        {
            worksheet.Column(column).Width = 30;
        }
        else if (key.Contains("cardname") || key == "name")
        {
            worksheet.Column(column).Width = 24;
        }
        else
        {
            worksheet.Column(column).Width = 15;
            if (key is "total" or "qty" || key.Contains("request") || key.Contains("patient") || key.Contains("male") ||
                key.Contains("female") || key.Contains("child") || key is "lr" or "icu" or "nicu" or "nursery" ||
                key.StartsWith("w4", StringComparison.Ordinal) || key.StartsWith("w5", StringComparison.Ordinal) ||
                key.StartsWith("w6", StringComparison.Ordinal) || key.Contains("semi") || key.Contains("stepdwon"))
            {
                columnRange.Style.NumberFormat.Format = "#,##0;[Red]-#,##0";
                columnRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
            }
        }
    }

    if (lastRow < 3)
        return;

    var dataRange = worksheet.Range(3, 1, lastRow, lastColumn);
    dataRange.Style.Border.TopBorder = XLBorderStyleValues.None;
    dataRange.Style.Border.BottomBorder = XLBorderStyleValues.None;
    dataRange.Style.Border.LeftBorder = XLBorderStyleValues.None;
    dataRange.Style.Border.RightBorder = XLBorderStyleValues.None;

    for (var row = 3; row <= lastRow; row++)
    {
        var rowRange = worksheet.Range(row, 1, row, lastColumn);
        rowRange.Style.Font.FontName = "Arial";
        rowRange.Style.Font.FontSize = 10;
        rowRange.Style.Font.FontColor = XLColor.FromHtml("#243447");
        rowRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        rowRange.Style.Border.BottomBorder = XLBorderStyleValues.Hair;
        rowRange.Style.Border.BottomBorderColor = XLColor.FromHtml("#D9E2EC");
        rowRange.Style.Fill.BackgroundColor = row % 2 == 0
            ? XLColor.FromHtml("#F2F6FA")
            : XLColor.White;
        worksheet.Row(row).Height = 20;
    }
}

static void UpsertDailyRows(string outputPath, string templatePath, IReadOnlyList<DailyReportRow> rows)
{
    var outputDirectory = Path.GetDirectoryName(outputPath)
        ?? throw new InvalidOperationException("ไม่สามารถระบุโฟลเดอร์ปลายทางของรายงานได้");
    Directory.CreateDirectory(outputDirectory);
    var outputExists = File.Exists(outputPath);
    var workbook = outputExists ? new XLWorkbook(outputPath) : new XLWorkbook(templatePath);
    using (workbook)
    {
        var worksheet = workbook.Worksheet("DailyReport");
        EnsureGroupedHeaders(worksheet);
        string[] metricKeys =
        [
            "financialopd", "financialipd", "doctorfeeopd", "doctorfeeipd",
            "labopdrequest", "labopdqty", "labopdamount",
            "labipdrequest", "labipdqty", "labipdamount",
            "xrayopdrequest", "xrayopdamount", "xrayipdrequest", "xrayipdamount",
            "opdpatientold", "opdpatientnew", "opdpatienttotal",
            "oldopdmale", "oldopdfemale", "oldopdchild",
            "newopdmale", "newopdfemale", "newopdchild"
        ];
        var columns = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var currentGroup = string.Empty;
        for (var column = 1; column <= 24; column++)
        {
            var groupHeader = worksheet.Cell(1, column).GetString().Trim();
            if (!string.IsNullOrEmpty(groupHeader))
                currentGroup = groupHeader;
            var subHeader = worksheet.Cell(2, column).GetString().Trim();
            var key = NormalizeHeader(currentGroup + subHeader);
            if (key == "date" || metricKeys.Contains(key, StringComparer.OrdinalIgnoreCase))
                columns[key] = column;
        }

        var missingColumns = metricKeys.Where(key => !columns.ContainsKey(key)).ToList();
        if (!columns.ContainsKey("date") || missingColumns.Count > 0)
        {
            throw new InvalidOperationException("ชีต DailyReport ขาดหัวคอลัมน์ที่จำเป็น: " + string.Join(", ", missingColumns));
        }

        var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 2;
        var rowByDate = new Dictionary<DateTime, int>();
        for (var row = 3; row <= lastRow; row++)
        {
            if (!TryReadDateCell(worksheet.Cell(row, columns["date"]), out var date))
                continue;
            if (!rowByDate.TryAdd(date, row))
                throw new InvalidOperationException($"พบวันที่ซ้ำในชีต DailyReport: {date:yyyy-MM-dd}");
        }

        foreach (var reportRow in rows.OrderBy(row => row.Date))
        {
            if (!rowByDate.TryGetValue(reportRow.Date, out var targetRow))
            {
                targetRow = ++lastRow;
                rowByDate.Add(reportRow.Date, targetRow);
            }

            worksheet.Cell(targetRow, columns["date"]).Value = reportRow.Date;
            worksheet.Cell(targetRow, columns["date"]).Style.DateFormat.Format = "yyyy-mm-dd";
            for (var metricIndex = 0; metricIndex < metricKeys.Length; metricIndex++)
                worksheet.Cell(targetRow, columns[metricKeys[metricIndex]]).Value = reportRow.Metrics[metricIndex];
        }

        FormatDailyReport(worksheet);

        if (outputExists)
            workbook.Save();
        else
            workbook.SaveAs(outputPath);
    }
}

static void UpsertRatioByTimeRows(string outputPath, IReadOnlyList<RatioByTimeRow> rows)
{
    using var workbook = new XLWorkbook(outputPath);
    var worksheet = workbook.Worksheet("Ratios by time");
    var columns = worksheet.Row(1).CellsUsed()
        .ToDictionary(cell => NormalizeHeader(cell.GetString()), cell => cell.Address.ColumnNumber, StringComparer.OrdinalIgnoreCase);
    string[] requiredHeaders = ["date", "depname", "time", "total"];
    var missingHeaders = requiredHeaders.Where(header => !columns.ContainsKey(header)).ToList();
    if (missingHeaders.Count > 0)
        throw new InvalidOperationException("ชีต Ratios by time ขาดหัวคอลัมน์: " + string.Join(", ", missingHeaders));

    var header = worksheet.Range(1, 1, 1, 4);
    header.Style.Fill.BackgroundColor = XLColor.FromHtml("#163A5F");
    header.Style.Font.FontName = "Arial";
    header.Style.Font.FontSize = 10;
    header.Style.Font.Bold = true;
    header.Style.Font.FontColor = XLColor.White;
    header.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
    header.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
    header.Style.Alignment.WrapText = true;
    header.Style.Border.BottomBorder = XLBorderStyleValues.Medium;
    header.Style.Border.BottomBorderColor = XLColor.FromHtml("#2F75B5");
    worksheet.Row(1).Height = 26;
    worksheet.SheetView.FreezeRows(1);

    var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
    var rowByKey = new Dictionary<(DateTime Date, string Department, string Time), int>();
    for (var row = 2; row <= lastRow; row++)
    {
        if (!TryReadDateCell(worksheet.Cell(row, columns["date"]), out var date))
            continue;
        var department = worksheet.Cell(row, columns["depname"]).GetString().Trim();
        var timeCode = worksheet.Cell(row, columns["time"]).GetString().Trim();
        var key = (date, department, timeCode);
        if (!rowByKey.TryAdd(key, row))
            throw new InvalidOperationException($"พบข้อมูลซ้ำในชีต Ratios by time: {date:yyyy-MM-dd}, {department}, {timeCode}");
    }

    foreach (var reportRow in rows.OrderBy(row => row.VisitDate)
                                  .ThenBy(row => row.DepName, StringComparer.CurrentCultureIgnoreCase)
                                  .ThenBy(row => row.TimeCode, StringComparer.Ordinal))
    {
        var key = (reportRow.VisitDate.Date, reportRow.DepName.Trim(), reportRow.TimeCode);
        if (!rowByKey.TryGetValue(key, out var targetRow))
        {
            targetRow = ++lastRow;
            rowByKey.Add(key, targetRow);
        }

        worksheet.Cell(targetRow, columns["date"]).Value = reportRow.VisitDate.Date;
        worksheet.Cell(targetRow, columns["date"]).Style.DateFormat.Format = "yyyy-mm-dd";
        worksheet.Cell(targetRow, columns["depname"]).Value = reportRow.DepName.Trim();
        worksheet.Cell(targetRow, columns["time"]).Value = reportRow.TimeCode;
        worksheet.Cell(targetRow, columns["total"]).Value = reportRow.Total;
    }

    worksheet.Column(columns["date"]).Width = 14;
    worksheet.Column(columns["depname"]).Width = 30;
    worksheet.Column(columns["time"]).Width = 10;
    worksheet.Column(columns["total"]).Width = 14;
    if (lastRow >= 2)
    {
        var dataRange = worksheet.Range(2, 1, lastRow, 4);
        dataRange.Style.Font.FontName = "Arial";
        dataRange.Style.Font.FontSize = 10;
        dataRange.Style.Font.FontColor = XLColor.FromHtml("#243447");
        dataRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        worksheet.Range(2, columns["date"], lastRow, columns["date"])
            .Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        worksheet.Range(2, columns["time"], lastRow, columns["time"])
            .Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        worksheet.Range(2, columns["total"], lastRow, columns["total"])
            .Style.NumberFormat.Format = "#,##0;[Red]-#,##0";
        worksheet.Range(2, columns["total"], lastRow, columns["total"])
            .Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        for (var row = 2; row <= lastRow; row++)
        {
            worksheet.Row(row).Height = 20;
            worksheet.Range(row, 1, row, 4).Style.Fill.BackgroundColor = row % 2 == 0
                ? XLColor.FromHtml("#F2F6FA")
                : XLColor.White;
        }
    }

    workbook.Save();
}

// ใช้การตั้งค่า connection string ตามเดิม
var builder = new ConfigurationBuilder()
    .SetBasePath(AppContext.BaseDirectory)
    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

IConfiguration config = builder.Build();
var connectionString = config.GetConnectionString("ReportDb");
if (string.IsNullOrWhiteSpace(connectionString))
{
    Console.Error.WriteLine("ไม่พบการตั้งค่า Connection string ที่ชื่อ 'ReportDb'");
    return 1;
}

var commandTimeout = config.GetValue<int>("Sql:CommandTimeoutSeconds", 60);
var startArgument = GetOption(args, "--start");
var endArgument = GetOption(args, "--end");
var outputArgument = GetOption(args, "--output");
if ((startArgument is null) != (endArgument is null))
{
    Console.Error.WriteLine("ถ้าระบุช่วงวันที่ ต้องใส่ทั้ง --start และ --end");
    return 1;
}

var desktopDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
if (string.IsNullOrWhiteSpace(desktopDirectory))
    desktopDirectory = @"C:\Users\atsawin.su\Desktop";
var defaultOutputPath = Path.Combine(desktopDirectory, "win", "daily-report", "Daily Report Dashboard - Grouped.xlsx");
var outputPath = Path.GetFullPath(outputArgument ?? defaultOutputPath);
var templatePath = Path.Combine(AppContext.BaseDirectory, "Templates", "Daily Report Dashboard - Template.xlsx");
if (string.Equals(outputPath, Path.GetFullPath(templatePath), StringComparison.OrdinalIgnoreCase))
{
    Console.Error.WriteLine("พาธปลายทางต้องไม่ใช่ไฟล์ Template ต้นฉบับ");
    return 1;
}
if (!File.Exists(templatePath))
{
    Console.Error.WriteLine($"ไม่พบไฟล์ Template: {templatePath}");
    return 1;
}

DateTime startDate;
DateTime endDate;
if (startArgument is not null && endArgument is not null)
{
    startDate = ParseCommandDate(startArgument, "--start");
    endDate = ParseCommandDate(endArgument, "--end");
}
else
{
    Console.WriteLine("กรุณากรอกวันที่ (รูปแบบ yyyy-MM-dd)");
    Console.Write("เริ่มต้น - ");
    var startInput = Console.ReadLine();
    Console.Write("สิ้นสุด - ");
    var endInput = Console.ReadLine();

    if (string.IsNullOrWhiteSpace(startInput) || string.IsNullOrWhiteSpace(endInput))
    {
        Console.Error.WriteLine("กรุณากรอกวันที่เริ่มต้นและวันที่สิ้นสุดให้ครบ");
        return 1;
    }

    try
    {
        startDate = ParseCommandDate(startInput, "เริ่มต้น");
        endDate = ParseCommandDate(endInput, "สิ้นสุด");
    }
    catch (ArgumentException ex)
    {
        Console.Error.WriteLine(ex.Message);
        return 1;
    }
}

if (startDate > endDate)
{
    Console.Error.WriteLine("วันเริ่มต้นต้องไม่อยู่หลังวันสิ้นสุด");
    return 1;
}

var startKey = startDate.ToString("yyyyMMdd", CultureInfo.InvariantCulture);
var endExclusiveKey = endDate.AddDays(1).ToString("yyyyMMdd", CultureInfo.InvariantCulture);
Console.WriteLine($"กำลังดึงข้อมูลรายงานวันที่ {startDate:yyyy-MM-dd} ถึง {endDate:yyyy-MM-dd}");

try
{
    var options = new DbContextOptionsBuilder<ReportDbContext>()
        .UseSqlServer(connectionString, sql => sql.CommandTimeout(commandTimeout))
        .Options;

    List<DailyReportRow> reportRows;
    List<RatioByTimeRow> ratioRows;
    await using (var database = new ReportDbContext(options))
    {
        // คำสั่งนี้จะถูกแปลงเป็น SQL และประมวลผลที่ Database Server โดยตรง
        var aggregatedRows = await database.OrpInfRows
            .AsNoTracking()
            .Where(row => row.OrpUpdDtm != null &&
                          row.OrpUpdDtm.Substring(0, 8).CompareTo(startKey) >= 0 &&
                          row.OrpUpdDtm.Substring(0, 8).CompareTo(endExclusiveKey) < 0 &&
                          row.OrpBilNum != null &&
                          (row.OrpBilNum.StartsWith("R") ||
                           row.OrpBilNum.StartsWith("I") ||
                           row.OrpBilNum.StartsWith("SSI")))
            .GroupBy(row => row.OrpUpdDtm!.Substring(0, 8))
            .Select(group => new
            {
                DateKey = group.Key,

                // คำนวณยอด OPD
                Opd = group.Sum(row => row.OrpPattyp == "O"
                    ? row.OrpBilNum!.StartsWith("R")
                        ? (row.OrpTotAmt ?? 0m) - (row.OrpInsAmt ?? 0m)
                        : (row.OrpInsAmt ?? 0m)
                    : 0m),

                // คำนวณยอด IPD
                Ipd = group.Sum(row => row.OrpPattyp == "I"
                    ? row.OrpBilNum!.StartsWith("R")
                        ? (row.OrpTotAmt ?? 0m) - (row.OrpInsAmt ?? 0m)
                        : (row.OrpInsAmt ?? 0m)
                    : 0m)
            })
            .OrderBy(row => row.DateKey)
            .ToListAsync();

        var financeRows = aggregatedRows
            .Select(row => new FinanceDailyRow(row.DateKey, row.Opd, row.Ipd))
            .ToDictionary(row => row.DateKey, StringComparer.Ordinal);

        var operationalRows = await ReportDataQueries.GetOperationalDailyRowsAsync(database, startKey, endDate.ToString("yyyyMMdd", CultureInfo.InvariantCulture));
        var patientRows = await ReportDataQueries.GetOpdPatientDailyRowsAsync(database, startKey, endDate.ToString("yyyyMMdd", CultureInfo.InvariantCulture));
        ratioRows = await ReportDataQueries.GetRatioByTimeRowsAsync(database, startDate, endDate);
        var patientRowsByDate = patientRows.ToDictionary(row => row.DateKey, StringComparer.Ordinal);
        reportRows = operationalRows
            .Select(row =>
            {
                financeRows.TryGetValue(row.DateKey, out var finance);
                patientRowsByDate.TryGetValue(row.DateKey, out var patients);
                return new DailyReportRow(row.DateKey,
                [
                    finance?.Opd ?? 0m,
                    finance?.Ipd ?? 0m,
                    row.DoctorFeeO,
                    row.DoctorFeeI,
                    row.LabORequest,
                    row.LabOQty,
                    row.LabOAmount,
                    row.LabIRequest,
                    row.LabIQty,
                    row.LabIAmount,
                    row.XrayORequest,
                    row.XrayOAmount,
                    row.XrayIRequest,
                    row.XrayIAmount,
                    patients?.OldTotal ?? 0,
                    patients?.NewTotal ?? 0,
                    patients?.Total ?? 0,
                    patients?.OldMale ?? 0,
                    patients?.OldFemale ?? 0,
                    patients?.OldChild ?? 0,
                    patients?.NewMale ?? 0,
                    patients?.NewFemale ?? 0,
                    patients?.NewChild ?? 0
                ]);
            })
            .ToList();
    }

    UpsertDailyRows(outputPath, templatePath, reportRows);
    UpsertRatioByTimeRows(outputPath, ratioRows);
    Console.WriteLine($"บันทึกรายงาน {reportRows.Count} วัน และ Ratios by time {ratioRows.Count} แถวแล้ว: {outputPath}");
}
catch (SqlException ex)
{
    var state = ex.Errors.Count > 0 ? ex.Errors[0].State : (byte)0;
    Console.Error.WriteLine($"การดึงข้อมูล SQL ไม่สำเร็จ (SQL error {ex.Number}, state {state})");
    return 1;
}
catch (Exception ex)
{
    Console.Error.WriteLine($"ทำงานไม่สำเร็จ ({ex.GetType().Name}): {ex.Message}");
    return 1;
}
return 0;
