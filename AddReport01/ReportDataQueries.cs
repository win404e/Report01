using Microsoft.EntityFrameworkCore;

public static class ReportDataQueries
{
    private static readonly string[] ActiveDepartmentCodes =
    [
        "ACP", "ANS", "CHK", "CHKIS", "CHKP", "DENT", "DRESS", "ENT", "ER", "EYE", "HCT", "HEMO",
        "LAB", "MED", "OBGY", "OPREM", "ORTHO", "OUTSD", "PED", "PHA", "PHY", "PSY", "SKIN", "SURG",
        "WBC", "XRAY"
    ];

    private static async Task<List<TResult>> LoadInChunksAsync<TValue, TResult>(
        IEnumerable<TValue> values,
        Func<TValue[], Task<List<TResult>>> loadChunk)
    {
        var results = new List<TResult>();
        foreach (var chunk in values.Chunk(500))
            results.AddRange(await loadChunk(chunk));
        return results;
    }

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

    public static async Task<List<WardDailyCountRow>> GetWardDailyCountsAsync(
        ReportDbContext database,
        string startKey,
        string endKey)
    {
        return await (
            from endedVisit in database.OcmEndRows.AsNoTracking()
            join subscriber in database.SbsInfRows.AsNoTracking()
                on endedVisit.OcmSbsNum equals subscriber.SbsNum
            join patientRow in database.PbsInfRows.AsNoTracking()
                on endedVisit.OcmChtNum equals patientRow.PbsChtNum into patientJoin
            from patient in patientJoin.DefaultIfEmpty()
            where subscriber.SbsAccDte != null &&
                  string.Compare(subscriber.SbsAccDte, startKey) >= 0 &&
                  string.Compare(subscriber.SbsAccDte, endKey) <= 0 &&
                  subscriber.SbsCurStt == "E" &&
                  endedVisit.OcmPatTyp == "I" &&
                  (endedVisit.OcmComTyp == "W" || endedVisit.OcmComTyp == "E" || endedVisit.OcmComTyp == "D")
            let validDates = database.IsSqlDate(subscriber.SbsAccDte) == 1 &&
                             patient.PbsBirDte != null &&
                             database.IsSqlDate(patient.PbsBirDte.Substring(0, 8)) == 1
            let age = validDates
                ? Convert.ToInt32(subscriber.SbsAccDte!.Substring(0, 4)) -
                  Convert.ToInt32(patient!.PbsBirDte!.Substring(0, 4)) -
                  (string.Compare(subscriber.SbsAccDte!.Substring(4, 4), patient!.PbsBirDte!.Substring(4, 4)) < 0 ? 1 : 0)
                : 15
            group new { endedVisit, PatientGroup = age < 15 ? "PED" : "NORMAL" } by new
            {
                DateKey = subscriber.SbsAccDte,
                WardCode = endedVisit.OcmWrdCod,
                PatientGroup = age < 15 ? "PED" : "NORMAL"
            }
            into grouped
            orderby grouped.Key.DateKey, grouped.Key.WardCode, grouped.Key.PatientGroup
            select new WardDailyCountRow(
                grouped.Key.DateKey!,
                grouped.Key.WardCode,
                grouped.Key.PatientGroup,
                grouped.Count(row => row.endedVisit.OcmComTyp == "W"),
                grouped.Count(row => row.endedVisit.OcmComTyp == "E"),
                grouped.Count(row => row.endedVisit.OcmComTyp == "D")))
            .ToListAsync();
    }

    public static async Task<List<DepartmentOpdPatientDailyRow>> GetDepartmentOpdPatientDailyRowsAsync(
        ReportDbContext database,
        string startKey,
        string endKey)
    {
        var startBound = startKey + "0000";
        var endBound = endKey + "2359";

        var rows = await (
            from visit in database.OcmInfRows.AsNoTracking()
            where visit.OcmAcpDtm != null &&
                  string.Compare(visit.OcmAcpDtm, startBound) >= 0 &&
                  string.Compare(visit.OcmAcpDtm, endBound) <= 0 &&
                  visit.OcmPatTyp == "O" &&
                  visit.OcmVstNum != null && visit.OcmVstNum.Trim() != "" &&
                  visit.OcmComStt != null && !visit.OcmComStt.StartsWith("C")
            join patient in database.PbsInfRows.AsNoTracking()
                on visit.OcmChtNum equals patient.PbsChtNum
            group new { visit, patient } by new
            {
                DateKey = visit.OcmAcpDtm!.Substring(0, 8),
                visit.OcmDepCod
            }
            into grouped
            orderby grouped.Key.DateKey, grouped.Key.OcmDepCod
            select new DepartmentOpdPatientDailyRow(
                grouped.Key.DateKey,
                database.GetDepartmentName(grouped.Key.OcmDepCod!, grouped.Key.DateKey) ?? "",
                grouped.Count(row => row.patient.PbsSexTyp == "M" &&
                    (row.visit.OcmDgsNfs == null || row.visit.OcmDgsNfs.Trim() == "" ||
                     row.visit.OcmDgsNfs.Trim() == "0" || row.visit.OcmDgsNfs.Trim() == "2")),
                grouped.Count(row => row.patient.PbsSexTyp == "F" &&
                    (row.visit.OcmDgsNfs == null || row.visit.OcmDgsNfs.Trim() == "" ||
                     row.visit.OcmDgsNfs.Trim() == "0" || row.visit.OcmDgsNfs.Trim() == "2")),
                grouped.Count(row => row.patient.PbsSexTyp == "M" && row.visit.OcmDgsNfs == "1"),
                grouped.Count(row => row.patient.PbsSexTyp == "F" && row.visit.OcmDgsNfs == "1"),
                grouped.Count()))
            .ToListAsync();

        return rows;
    }

    public static async Task<RppayQueryRows> GetRppayReportRowsAsync(
        ReportDbContext database,
        string startDateTime,
        string endDateTime)
    {
        var rawOrp = await database.OrpInfRows.AsNoTracking()
            .Where(row => row.OrpUpdDtm != null &&
                          string.Compare(row.OrpUpdDtm, startDateTime) >= 0 &&
                          string.Compare(row.OrpUpdDtm, endDateTime) <= 0 &&
                          row.OrpBilNum != null && row.OrpBilNum != "SKIP")
            .Select(row => new RppaySourceRow(
                row.OrpOcmNum,
                row.OrpCtrCod == null ? "" : row.OrpCtrCod.Trim(),
                row.OrpInsCod,
                row.OrpUidCod == null ? null : row.OrpUidCod.Trim(),
                row.OrpUpdDtm!,
                row.OrpUpdDtm!.Substring(0, 8),
                row.OrpTotAmt ?? 0m,
                row.OrpDisAmt ?? 0m))
            .ToListAsync();

        var firstRawByVisit = rawOrp
            .GroupBy(row => row.OcmNumber)
            .Select(group => group.OrderBy(row => row.ContractCode == "TOTAL" ? 0 : 1)
                .ThenBy(row => row.UpdateDateTime)
                .First())
            .ToList();

        var insuranceCodes = rawOrp.Select(row => row.InsuranceCode).Where(code => code != null).Distinct().ToList();
        var insuranceRows = insuranceCodes.Count == 0
            ? []
            : await database.InsMstRows.AsNoTracking()
                .Where(row => row.InsCod != null && insuranceCodes.Contains(row.InsCod))
                .Select(row => new { row.InsCod, row.InsCodNam })
                .ToListAsync();
        var insuranceTypes = insuranceRows
            .Where(row => row.InsCod != null)
            .GroupBy(row => row.InsCod!.Value)
            .ToDictionary(group => group.Key, group => group.First().InsCodNam?.Trim());

        var filteredRows = rawOrp
            .Select(row => new
            {
                Row = row,
                InsuranceType = insuranceTypes.TryGetValue(row.InsuranceCode ?? 0, out var type) ? type : null
            })
            .Where(item => item.InsuranceType != "Contract" ||
                           firstRawByVisit.Any(first => ReferenceEquals(first, item.Row)))
            .Select(item => new RppayFilteredRow(item.Row, item.InsuranceType))
            .ToList();

        var visitNumbers = filteredRows.Select(row => row.Source.OcmNumber).Where(number => number != null)
            .Select(number => number!).Distinct().ToList();
        var ocmRows = await LoadInChunksAsync(visitNumbers, batch => database.OcmInfRows.AsNoTracking()
            .Where(row => row.OcmNum != null && batch.Contains(row.OcmNum))
            .Select(row => new RppayPatientStatusRow(row.OcmNum, row.OcmPatTyp))
            .ToListAsync());
        var patientTypes = ocmRows.Where(row => row.OcmNumber != null).GroupBy(row => row.OcmNumber!.Trim())
            .ToDictionary(group => group.Key, group => group.First().PatientType, StringComparer.OrdinalIgnoreCase);
        var discountRows = await LoadInChunksAsync(visitNumbers, batch => database.OdlInfRows.AsNoTracking()
            .Where(row => row.OdlOcmNum != null && batch.Contains(row.OdlOcmNum))
            .Select(row => new RppayDiscountRow(row.OdlOcmNum, row.OdlTotAmt, row.OdlChkAmt, row.OdlCtrTyp))
            .ToListAsync());
        var contractDiscounts = discountRows
            .Where(row => row.ContractType != null && row.ContractType.Trim() != "TOTAL")
            .GroupBy(row => row.OcmNumber)
            .Where(group => group.Key != null)
            .ToDictionary(
                group => group.Key!.Trim(),
                group => group.Sum(row => (row.TotalAmount ?? 0m) * (row.CheckAmount ?? 0m)),
                StringComparer.Ordinal);

        foreach (var group in filteredRows.GroupBy(row => row.Source.OcmNumber))
        {
            var first = group.OrderBy(row => row.Source.ContractCode == "TOTAL" ? 0 : 1)
                .ThenBy(row => row.Source.UpdateDateTime)
                .First();
            first.IsFirstVisitRow = true;
        }

        var cardRows = await LoadInChunksAsync(visitNumbers, batch => database.CrdInfRows.AsNoTracking()
            .Where(row => row.CrdOcmNum != null && batch.Contains(row.CrdOcmNum))
            .GroupBy(row => row.CrdOcmNum)
            .Select(group => new RppayCardInfoRow(group.Key, group.Max(row => row.CrdCod)))
            .ToListAsync());
        var cardCodes = cardRows
            .Where(row => row.OcmNumber != null)
            .GroupBy(row => row.OcmNumber!)
            .ToDictionary(
                group => group.Key.Trim(),
                group => group.First().Code?.Trim(),
                StringComparer.OrdinalIgnoreCase);

        var userCodes = filteredRows.Select(row => row.Source.UserCode).Where(code => code != null)
            .Select(code => code!).Distinct().ToList();
        var users = await LoadInChunksAsync(userCodes, batch => database.UidMstRows.AsNoTracking()
            .Where(row => row.UidCod != null && batch.Contains(row.UidCod))
            .Select(row => new RppayUserRow(row.UidCod, row.UidNam))
            .ToListAsync());
        var userNames = users.Where(row => row.UserCode != null).GroupBy(row => row.UserCode!)
            .ToDictionary(group => group.Key.Trim(), group => group.First().UserName, StringComparer.OrdinalIgnoreCase);

        var requestedCardCodes = cardCodes.Values.Where(code => code != null).Select(code => code!).Distinct().ToList();
        var details = await LoadInChunksAsync(requestedCardCodes, batch => database.DtlMstRows.AsNoTracking()
            .Where(row => row.DtlTblCod == "CRDTBL" && row.DtlCod != null && batch.Contains(row.DtlCod))
            .Select(row => new RppayCardNameRow(row.DtlCod, row.DtlCodNam))
            .ToListAsync());
        var cardNames = details.Where(row => row.Code != null).GroupBy(row => row.Code!.Trim())
            .ToDictionary(group => group.Key, group => group.First().Name, StringComparer.OrdinalIgnoreCase);

        var preparedRows = filteredRows
            .Select(row =>
            {
                var source = row.Source;
                cardCodes.TryGetValue(source.OcmNumber?.Trim() ?? "", out var cardCode);
                userNames.TryGetValue(source.UserCode ?? "", out var userName);
                cardNames.TryGetValue(cardCode ?? "", out var cardName);
                contractDiscounts.TryGetValue(source.OcmNumber?.Trim() ?? "", out var contractDiscount);
                patientTypes.TryGetValue(source.OcmNumber?.Trim() ?? "", out var patientType);
                var contractAmount = row.IsFirstVisitRow ? contractDiscount : 0m;
                var status = patientType == "I" ? "IPD" : patientType == "O" ? "OPD" : patientType ?? "";
                return new
                {
                    source.DateKey,
                    UserName = userName ?? "",
                    Status = status,
                    Type = row.InsuranceType ?? "",
                    CardCode = cardCode ?? "",
                    CardName = cardName ?? "",
                    NetAmount = source.Amount - source.DCCash - contractAmount,
                    DiscountAmount = source.DCCash + contractAmount
                };
            })
            .ToList();

        var byUser = preparedRows
            .GroupBy(row => new { row.DateKey, row.UserName, row.CardCode, row.CardName })
            .Select(group => new RppayDailyRow(
                group.Key.DateKey,
                group.Key.UserName,
                group.Key.CardCode,
                group.Key.CardName,
                group.Sum(row => row.NetAmount)))
            .OrderBy(row => row.DateKey)
            .ThenBy(row => row.UserName, StringComparer.CurrentCultureIgnoreCase)
            .ThenBy(row => row.Code, StringComparer.Ordinal)
            .ToList();

        var byType = preparedRows
            .GroupBy(row => new { row.DateKey, row.Status, row.Type, row.CardCode, row.CardName })
            .Select(group => new RppayByTypeRow(
                group.Key.DateKey,
                group.Key.Status,
                group.Key.Type,
                group.Key.CardCode,
                group.Key.CardName,
                group.Sum(row => row.NetAmount),
                group.Sum(row => row.DiscountAmount)))
            .OrderBy(row => row.DateKey)
            .ThenBy(row => row.Status, StringComparer.Ordinal)
            .ThenBy(row => row.Type, StringComparer.CurrentCultureIgnoreCase)
            .ThenBy(row => row.Code, StringComparer.Ordinal)
            .ToList();

        return new RppayQueryRows(byUser, byType);
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

    private sealed record RppaySourceRow(
        string? OcmNumber,
        string ContractCode,
        short? InsuranceCode,
        string? UserCode,
        string UpdateDateTime,
        string DateKey,
        decimal Amount,
        decimal DCCash);

    private sealed class RppayFilteredRow(RppaySourceRow source, string? insuranceType)
    {
        public RppaySourceRow Source { get; } = source;
        public string? InsuranceType { get; } = insuranceType;
        public bool IsFirstVisitRow { get; set; }
    }

    private sealed record RppayDiscountRow(
        string? OcmNumber,
        decimal? TotalAmount,
        decimal? CheckAmount,
        string? ContractType);

    private sealed record RppayPatientStatusRow(string? OcmNumber, string? PatientType);
    private sealed record RppayCardInfoRow(string? OcmNumber, string? Code);
    private sealed record RppayUserRow(string? UserCode, string? UserName);
    private sealed record RppayCardNameRow(string? Code, string? Name);
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

public sealed record WardDailyCountRow(
    string DateKey,
    string? WardCode,
    string PatientGroup,
    int WardInfo,
    int NewAdmit,
    int Discharge);

public sealed record DepartmentOpdPatientDailyRow(
    string DateKey,
    string DepName,
    int OldMale,
    int OldFemale,
    int NewMale,
    int NewFemale,
    int Total)
{
    public DateTime VisitDate => DateTime.ParseExact(DateKey, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
}

public sealed record RppayDailyRow(
    string DateKey,
    string UserName,
    string Code,
    string CardName,
    decimal NetAmount)
{
    public DateTime Date => DateTime.ParseExact(DateKey, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
}

public sealed record RppayByTypeRow(
    string DateKey,
    string Status,
    string Type,
    string Code,
    string CardName,
    decimal SumTotalAmount,
    decimal SumDiscount)
{
    public DateTime Date => DateTime.ParseExact(DateKey, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
}

public sealed record RppayQueryRows(List<RppayDailyRow> ByUser, List<RppayByTypeRow> ByType);

public sealed record RatioByTimeRow(DateTime VisitDate, string DepName, string TimeCode, int Total);
