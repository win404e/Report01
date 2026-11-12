using Microsoft.EntityFrameworkCore;

public static class ReportDataQueries
{
    private static readonly string[] ActiveDepartmentCodes =
    [
        "ACP", "ANS", "CHK", "CHKIS", "CHKP", "DENT", "DRESS", "ENT", "ER", "EYE", "HCT", "HEMO",
        "LAB", "MED", "OBGY", "OPREM", "ORTHO", "OUTSD", "PED", "PHA", "PHY", "PSY", "SKIN", "SURG",
        "WBC", "XRAY"
    ];

    public static async Task<List<OperationalDailyRow>> GetOperationalDailyRowsAsync(
        ReportDbContext database,
        string startKey,
        string endKey)
    {
        IQueryable<RawStsRow> RawSts(string patientType, string receiptType) =>
            from result in database.StsOdlEnd(startKey, endKey, patientType, receiptType)
            join subscriber in database.SbsInfRows on result.SbsNum equals subscriber.SbsNum
            select new RawStsRow
            {
                PatientType = patientType,
                AccDate = subscriber.SbsAccDte,
                IncNam = result.IncNam,
                IncCod = result.IncCod,
                OdlAmt = result.OdlAmt
            };

        var rawSts = RawSts("O", "R")
            .Concat(RawSts("O", "I"))
            .Concat(RawSts("I", "R"))
            .Concat(RawSts("I", "I"));

        var financialRows = await rawSts
            .GroupBy(row => row.AccDate)
            .Select(group => new FinancialAggregate
            {
                DateKey = group.Key,
                DoctorFeeO = group.Sum(row => row.PatientType == "O" &&
                    row.IncNam != null &&
                    (EF.Functions.Collate(row.IncNam, "Thai_CI_AI_KS_WS").Contains("ตรวจพิเศษ") ||
                     EF.Functions.Collate(row.IncNam, "Thai_CI_AI_KS_WS").Contains("ตรวจรักษา")) &&
                    row.IncCod != null && row.IncCod != 9996 && row.IncCod != 9998 && row.IncCod != 9999
                    ? row.OdlAmt ?? 0m : 0m),
                DoctorFeeI = group.Sum(row => row.PatientType == "I" &&
                    row.IncNam != null &&
                    (EF.Functions.Collate(row.IncNam, "Thai_CI_AI_KS_WS").Contains("ตรวจพิเศษ") ||
                     EF.Functions.Collate(row.IncNam, "Thai_CI_AI_KS_WS").Contains("ตรวจรักษา")) &&
                    row.IncCod != null && row.IncCod != 9996 && row.IncCod != 9998 && row.IncCod != 9999
                    ? row.OdlAmt ?? 0m : 0m),
                LabOAmount = group.Sum(row => row.PatientType == "O" &&
                    row.IncNam != null && EF.Functions.Collate(row.IncNam, "Thai_CI_AI_KS_WS").Contains("ปฏิบัติ") &&
                    row.IncCod != null && row.IncCod != 9996 && row.IncCod != 9998 && row.IncCod != 9999
                    ? row.OdlAmt ?? 0m : 0m),
                LabIAmount = group.Sum(row => row.PatientType == "I" &&
                    row.IncNam != null && EF.Functions.Collate(row.IncNam, "Thai_CI_AI_KS_WS").Contains("ปฏิบัติ") &&
                    row.IncCod != null && row.IncCod != 9996 && row.IncCod != 9998 && row.IncCod != 9999
                    ? row.OdlAmt ?? 0m : 0m),
                XrayOAmount = group.Sum(row => row.PatientType == "O" &&
                    row.IncNam != null && EF.Functions.Collate(row.IncNam, "Thai_CI_AI_KS_WS").Contains("เอกซเรย์") &&
                    row.IncCod != null && row.IncCod != 9996 && row.IncCod != 9998 && row.IncCod != 9999
                    ? row.OdlAmt ?? 0m : 0m),
                XrayIAmount = group.Sum(row => row.PatientType == "I" &&
                    row.IncNam != null && EF.Functions.Collate(row.IncNam, "Thai_CI_AI_KS_WS").Contains("เอกซเรย์") &&
                    row.IncCod != null && row.IncCod != 9996 && row.IncCod != 9998 && row.IncCod != 9999
                    ? row.OdlAmt ?? 0m : 0m)
            })
            .ToListAsync();

        var startBound = startKey + "0000";
        var endBound = endKey + "2359";

        var labOrderRows = await (
            from order in database.OdrInfRows
            join visit in database.OcmInfRows on order.OdrOcmNum equals visit.OcmNum
            where order.OdrDtm != null &&
                  string.Compare(order.OdrDtm, startBound) >= 0 &&
                  string.Compare(order.OdrDtm, endBound) <= 0 &&
                  order.OdrTrnDep == "LAB"
            group visit by order.OdrDtm!.Substring(0, 8) into groupedVisits
            select new CountAggregate
            {
                DateKey = groupedVisits.Key,
                Opd = groupedVisits.Count(visit => visit.OcmPatTyp == "O"),
                Ipd = groupedVisits.Count(visit => visit.OcmPatTyp == "I")
            })
            .ToListAsync();

        var labPatientRows = await (
            from reservation in database.LbqInfRows
            join visit in database.OcmInfRows on reservation.LbqOcmNum equals visit.OcmNum
            where reservation.LbqRsvDte != null &&
                  string.Compare(reservation.LbqRsvDte, startKey) >= 0 &&
                  string.Compare(reservation.LbqRsvDte, endKey) <= 0 &&
                  reservation.LbqTrnDep == "LAB" &&
                  reservation.LbqChtNum != null
            group reservation by new
            {
                DateKey = reservation.LbqRsvDte!.Substring(0, 8),
                visit.OcmPatTyp
            } into groupedReservations
            select new PatientCountAggregate
            {
                DateKey = groupedReservations.Key.DateKey,
                PatientType = groupedReservations.Key.OcmPatTyp,
                PatientCount = groupedReservations.Select(row => row.LbqChtNum).Distinct().Count()
            })
            .ToListAsync();

        var xrayOrderRows = await (
            from order in database.OdrInfRows
            join visit in database.OcmInfRows on order.OdrOcmNum equals visit.OcmNum
            where order.OdrDtm != null &&
                  string.Compare(order.OdrDtm, startBound) >= 0 &&
                  string.Compare(order.OdrDtm, endBound) <= 0 &&
                  order.OdrTrnDep == "XRAY" &&
                  order.OdrDelFlg != null && order.OdrDelFlg != "C"
            group visit by order.OdrDtm!.Substring(0, 8) into groupedVisits
            select new CountAggregate
            {
                DateKey = groupedVisits.Key,
                Opd = groupedVisits.Count(visit => visit.OcmPatTyp == "O"),
                Ipd = groupedVisits.Count(visit => visit.OcmPatTyp == "I")
            })
            .ToListAsync();

        var financialByDate = financialRows.ToDictionary(row => row.DateKey ?? "", row => row);
        var labOrdersByDate = labOrderRows.ToDictionary(row => row.DateKey, row => row);
        var labPatientsByDate = labPatientRows.ToDictionary(row => (row.DateKey, row.PatientType), row => row.PatientCount);
        var xrayOrdersByDate = xrayOrderRows.ToDictionary(row => row.DateKey, row => row);

        var startDate = DateTime.ParseExact(startKey, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
        var endDate = DateTime.ParseExact(endKey, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
        var rows = new List<OperationalDailyRow>();
        for (var date = startDate; date <= endDate; date = date.AddDays(1))
        {
            var dateKey = date.ToString("yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
            financialByDate.TryGetValue(dateKey, out var financial);
            labOrdersByDate.TryGetValue(dateKey, out var labOrders);
            xrayOrdersByDate.TryGetValue(dateKey, out var xrayOrders);
            labPatientsByDate.TryGetValue((dateKey, "O"), out var labOpdPatients);
            labPatientsByDate.TryGetValue((dateKey, "I"), out var labIpdPatients);

            rows.Add(new OperationalDailyRow(
                dateKey,
                financial?.DoctorFeeO ?? 0m,
                financial?.DoctorFeeI ?? 0m,
                labOrders?.Opd ?? 0,
                labOpdPatients,
                financial?.LabOAmount ?? 0m,
                labOrders?.Ipd ?? 0,
                labIpdPatients,
                financial?.LabIAmount ?? 0m,
                xrayOrders?.Opd ?? 0,
                financial?.XrayOAmount ?? 0m,
                xrayOrders?.Ipd ?? 0,
                financial?.XrayIAmount ?? 0m));
        }

        return rows;
    }

    public static async Task<List<OpdPatientDailyRow>> GetOpdPatientDailyRowsAsync(
        ReportDbContext database,
        string startKey,
        string endKey)
    {
        var startBound = startKey + "0000";
        var endBound = endKey + "2359";

        var activeCases = database.OcmInfRows
            .Where(row => row.OcmAcpDtm != null &&
                          string.Compare(row.OcmAcpDtm, startBound) >= 0 &&
                          string.Compare(row.OcmAcpDtm, endBound) <= 0 &&
                          row.OcmPatTyp == "O" &&
                          row.OcmVstNum != null && row.OcmVstNum.Trim() != "" &&
                          row.OcmComStt != null && !row.OcmComStt.StartsWith("C"))
            .Select(row => new OcmPatientCase
            {
                ChartNumber = row.OcmChtNum,
                AcceptDateTime = row.OcmAcpDtm,
                Age = row.OcmAgeOld,
                NewRepeatFlag = row.OcmDgsNfs
            });

        var historyCases = database.OcmHstRows
            .Where(row => row.OcmAcpDtm != null &&
                          string.Compare(row.OcmAcpDtm, startBound) >= 0 &&
                          string.Compare(row.OcmAcpDtm, endBound) <= 0 &&
                          row.OcmPatTyp == "O" &&
                          row.OcmVstNum != null && row.OcmVstNum.Trim() != "" &&
                          row.OcmComStt != null && !row.OcmComStt.StartsWith("C"))
            .Select(row => new OcmPatientCase
            {
                ChartNumber = row.OcmChtNum,
                AcceptDateTime = row.OcmAcpDtm,
                Age = row.OcmAgeOld,
                NewRepeatFlag = row.OcmDgsNfs
            });

        var combinedCases = activeCases.Concat(historyCases);
        var counts = await (
            from patientCase in combinedCases
            join patient in database.PbsInfRows on patientCase.ChartNumber equals patient.PbsChtNum
            group new { patientCase, patient } by patientCase.AcceptDateTime!.Substring(0, 8) into groupedCases
            select new OpdPatientDailyRow(
                groupedCases.Key,
                groupedCases.Count(row =>
                    Convert.ToInt32(row.patientCase.Age) >= 15 && row.patient.PbsSexTyp == "M" &&
                    (row.patientCase.NewRepeatFlag == null || row.patientCase.NewRepeatFlag == "0" || row.patientCase.NewRepeatFlag == "2")),
                groupedCases.Count(row =>
                    Convert.ToInt32(row.patientCase.Age) >= 15 && row.patient.PbsSexTyp == "F" &&
                    (row.patientCase.NewRepeatFlag == null || row.patientCase.NewRepeatFlag == "0" || row.patientCase.NewRepeatFlag == "2")),
                groupedCases.Count(row =>
                    Convert.ToInt32(row.patientCase.Age) < 15 &&
                    (row.patient.PbsSexTyp == "M" || row.patient.PbsSexTyp == "F") &&
                    (row.patientCase.NewRepeatFlag == null || row.patientCase.NewRepeatFlag == "0" || row.patientCase.NewRepeatFlag == "2")),
                groupedCases.Count(row =>
                    Convert.ToInt32(row.patientCase.Age) >= 15 && row.patient.PbsSexTyp == "M" &&
                    row.patientCase.NewRepeatFlag == "1"),
                groupedCases.Count(row =>
                    Convert.ToInt32(row.patientCase.Age) >= 15 && row.patient.PbsSexTyp == "F" &&
                    row.patientCase.NewRepeatFlag == "1"),
                groupedCases.Count(row =>
                    Convert.ToInt32(row.patientCase.Age) < 15 &&
                    (row.patient.PbsSexTyp == "M" || row.patient.PbsSexTyp == "F") &&
                    row.patientCase.NewRepeatFlag == "1")))
            .ToListAsync();

        return counts;
    }

    public static async Task<List<RatioByTimeRow>> GetRatioByTimeRowsAsync(
        ReportDbContext database,
        DateTime startDate,
        DateTime endDate)
    {
        var startKey = startDate.ToString("yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
        var endKey = endDate.ToString("yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
        var startBound = startKey + "0000";
        var endBound = endKey + "2359";

        var departmentNames = new List<(string Code, string Name)>();
        foreach (var code in ActiveDepartmentCodes)
        {
            var name = await database.OcmInfRows
                .AsNoTracking()
                .Select(_ => database.GetDepartmentName(code, endKey))
                .FirstOrDefaultAsync();
            departmentNames.Add((code, name?.Trim() ?? ""));
        }

        var visitAggregates = await database.OcmInfRows
            .AsNoTracking()
            .Where(row => row.OcmAcpDtm != null &&
                          string.Compare(row.OcmAcpDtm, startBound) >= 0 &&
                          string.Compare(row.OcmAcpDtm, endBound) <= 0 &&
                          row.OcmPatTyp == "O" &&
                          row.OcmVstNum != null && row.OcmVstNum != "          " &&
                          row.OcmComStt != null &&
                          row.OcmComStt != "CN" && row.OcmComStt != "CI" && row.OcmComStt != "CR" &&
                          row.OcmDepCod != null && ActiveDepartmentCodes.Contains(row.OcmDepCod))
            .GroupBy(row => new
            {
                DateKey = row.OcmAcpDtm!.Substring(0, 8),
                DepCode = row.OcmDepCod,
                TimeCode = string.Compare(row.OcmAcpDtm!.Substring(8, 4), "0800") >= 0 &&
                           string.Compare(row.OcmAcpDtm!.Substring(8, 4), "1600") <= 0
                    ? "T1"
                    : "T2"
            })
            .Select(group => new TimeVisitAggregate
            {
                DateKey = group.Key.DateKey,
                DepCode = group.Key.DepCode,
                TimeCode = group.Key.TimeCode,
                VisitTotal = group.Count()
            })
            .ToListAsync();

        var visitsByKey = visitAggregates.ToDictionary(
            row => (row.DateKey, row.DepCode?.Trim() ?? "", row.TimeCode),
            row => row.VisitTotal);
        var rows = new List<RatioByTimeRow>();
        for (var date = startDate.Date; date <= endDate.Date; date = date.AddDays(1))
        {
            var dateKey = date.ToString("yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
            foreach (var department in departmentNames.OrderBy(item => item.Name, StringComparer.CurrentCultureIgnoreCase))
            {
                foreach (var timeCode in new[] { "T1", "T2" })
                {
                    visitsByKey.TryGetValue((dateKey, department.Code, timeCode), out var total);
                    rows.Add(new RatioByTimeRow(date, department.Name, timeCode, total));
                }
            }
        }

        return rows;
    }

    private sealed class RawStsRow
    {
        public string? PatientType { get; set; }
        public string? AccDate { get; set; }
        public string? IncNam { get; set; }
        public short? IncCod { get; set; }
        public decimal? OdlAmt { get; set; }
    }

    private sealed class FinancialAggregate
    {
        public string? DateKey { get; set; }
        public decimal DoctorFeeO { get; set; }
        public decimal DoctorFeeI { get; set; }
        public decimal LabOAmount { get; set; }
        public decimal LabIAmount { get; set; }
        public decimal XrayOAmount { get; set; }
        public decimal XrayIAmount { get; set; }
    }

    private sealed class CountAggregate
    {
        public string DateKey { get; set; } = "";
        public int Opd { get; set; }
        public int Ipd { get; set; }
    }

    private sealed class PatientCountAggregate
    {
        public string DateKey { get; set; } = "";
        public string? PatientType { get; set; }
        public int PatientCount { get; set; }
    }

    private sealed class OcmPatientCase
    {
        public string? ChartNumber { get; set; }
        public string? AcceptDateTime { get; set; }
        public string? Age { get; set; }
        public string? NewRepeatFlag { get; set; }
    }

    private sealed class TimeVisitAggregate
    {
        public string DateKey { get; set; } = "";
        public string? DepCode { get; set; }
        public string TimeCode { get; set; } = "";
        public int VisitTotal { get; set; }
    }
}

public sealed record OperationalDailyRow(
    string DateKey,
    decimal DoctorFeeO,
    decimal DoctorFeeI,
    int LabORequest,
    int LabOQty,
    decimal LabOAmount,
    int LabIRequest,
    int LabIQty,
    decimal LabIAmount,
    int XrayORequest,
    decimal XrayOAmount,
    int XrayIRequest,
    decimal XrayIAmount);

public sealed record DailyReportRow(string DateKey, decimal[] Metrics)
{
    public DateTime Date => DateTime.ParseExact(DateKey, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
}

public sealed record OpdPatientDailyRow(
    string DateKey,
    int OldMale,
    int OldFemale,
    int OldChild,
    int NewMale,
    int NewFemale,
    int NewChild)
{
    public int OldTotal => OldMale + OldFemale + OldChild;
    public int NewTotal => NewMale + NewFemale + NewChild;
    public int Total => OldTotal + NewTotal;
}

public sealed record RatioByTimeRow(DateTime VisitDate, string DepName, string TimeCode, int Total);
