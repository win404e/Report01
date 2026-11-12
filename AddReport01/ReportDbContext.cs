using Microsoft.EntityFrameworkCore;

public sealed class OrpInfRow
{
    public string? OrpUpdDtm { get; set; }
    public string? OrpPattyp { get; set; }
    public string? OrpBilNum { get; set; }
    public decimal? OrpTotAmt { get; set; }
    public decimal? OrpInsAmt { get; set; }
}

public sealed class ReportDbContext(DbContextOptions<ReportDbContext> options) : DbContext(options)
{
    public DbSet<OrpInfRow> OrpInfRows => Set<OrpInfRow>();
    public DbSet<SbsInfRow> SbsInfRows => Set<SbsInfRow>();
    public DbSet<OdrInfRow> OdrInfRows => Set<OdrInfRow>();
    public DbSet<OcmInfRow> OcmInfRows => Set<OcmInfRow>();
    public DbSet<OcmHstRow> OcmHstRows => Set<OcmHstRow>();
    public DbSet<PbsInfRow> PbsInfRows => Set<PbsInfRow>();
    public DbSet<LbqInfRow> LbqInfRows => Set<LbqInfRow>();

    public IQueryable<StsOdlEndRow> StsOdlEnd(string startDate, string endDate, string patientType, string receiptType) =>
        FromExpression(() => StsOdlEnd(startDate, endDate, patientType, receiptType));

    public string? GetDepartmentName(string departmentCode, string dateKey) =>
        throw new InvalidOperationException("This function can only be used inside an EF Core query.");

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<OrpInfRow>(entity =>
        {
            entity.HasNoKey();
            entity.ToTable("orpinf");
            entity.Property(row => row.OrpUpdDtm).HasColumnName("orpupddtm");
            entity.Property(row => row.OrpPattyp).HasColumnName("OrpPattyp");
            entity.Property(row => row.OrpBilNum).HasColumnName("OrpBilNum");
            entity.Property(row => row.OrpTotAmt).HasColumnName("OrpTotAmt");
            entity.Property(row => row.OrpInsAmt).HasColumnName("OrpInsAmt");
        });

        modelBuilder.Entity<SbsInfRow>(entity =>
        {
            entity.HasNoKey();
            entity.ToTable("SbsInf");
            entity.Property(row => row.SbsNum).HasColumnName("SbsNum");
            entity.Property(row => row.SbsAccDte).HasColumnName("SbsAccDte");
        });

        modelBuilder.Entity<OdrInfRow>(entity =>
        {
            entity.HasNoKey();
            entity.ToTable("OdrInf");
            entity.Property(row => row.OdrDtm).HasColumnName("OdrDtm");
            entity.Property(row => row.OdrOcmNum).HasColumnName("OdrOcmNum");
            entity.Property(row => row.OdrTrnDep).HasColumnName("OdrTrnDep");
            entity.Property(row => row.OdrDelFlg).HasColumnName("OdrDelFlg");
        });

        modelBuilder.Entity<OcmInfRow>(entity =>
        {
            entity.HasNoKey();
            entity.ToTable("OcmInf");
            entity.Property(row => row.OcmNum).HasColumnName("OcmNum");
            entity.Property(row => row.OcmPatTyp).HasColumnName("OcmPatTyp");
            entity.Property(row => row.OcmChtNum).HasColumnName("OcmChtNum");
            entity.Property(row => row.OcmAcpDtm).HasColumnName("OcmAcpDtm");
            entity.Property(row => row.OcmAgeOld).HasColumnName("OcmAgeOld");
            entity.Property(row => row.OcmDgsNfs).HasColumnName("OcmDgsNfs");
            entity.Property(row => row.OcmVstNum).HasColumnName("OcmVstNum");
            entity.Property(row => row.OcmComStt).HasColumnName("OcmComStt");
            entity.Property(row => row.OcmDepCod).HasColumnName("OcmDepCod");
        });

        modelBuilder.Entity<OcmHstRow>(entity =>
        {
            entity.HasNoKey();
            entity.ToTable("OcmHst");
            entity.Property(row => row.OcmChtNum).HasColumnName("OcmChtNum");
            entity.Property(row => row.OcmAcpDtm).HasColumnName("OcmAcpDtm");
            entity.Property(row => row.OcmAgeOld).HasColumnName("OcmAgeOld");
            entity.Property(row => row.OcmDgsNfs).HasColumnName("OcmDgsNfs");
            entity.Property(row => row.OcmVstNum).HasColumnName("OcmVstNum");
            entity.Property(row => row.OcmComStt).HasColumnName("OcmComStt");
            entity.Property(row => row.OcmPatTyp).HasColumnName("OcmPatTyp");
        });

        modelBuilder.Entity<PbsInfRow>(entity =>
        {
            entity.HasNoKey();
            entity.ToTable("PbsInf");
            entity.Property(row => row.PbsChtNum).HasColumnName("PbsChtNum");
            entity.Property(row => row.PbsSexTyp).HasColumnName("PbsSexTyp");
        });

        modelBuilder.Entity<LbqInfRow>(entity =>
        {
            entity.HasNoKey();
            entity.ToTable("LbqInf");
            entity.Property(row => row.LbqOcmNum).HasColumnName("LbqOcmNum");
            entity.Property(row => row.LbqRsvDte).HasColumnName("LbqRsvDte");
            entity.Property(row => row.LbqTrnDep).HasColumnName("LbqTrnDep");
            entity.Property(row => row.LbqChtNum).HasColumnName("LbqChtNum");
        });

        var stsOdlEnd = modelBuilder.HasDbFunction(typeof(ReportDbContext).GetMethod(
            nameof(StsOdlEnd), [typeof(string), typeof(string), typeof(string), typeof(string)])!);
        stsOdlEnd.HasName("STS_OdlEnd").HasSchema("dbo");

        var getDepartmentName = modelBuilder.HasDbFunction(typeof(ReportDbContext).GetMethod(
            nameof(GetDepartmentName), [typeof(string), typeof(string)])!);
        getDepartmentName.HasName("GetDepartmentName").HasSchema("dbo");

        var stsOdlEndResult = modelBuilder.Entity<StsOdlEndRow>();
        stsOdlEndResult.HasNoKey();
        stsOdlEndResult.Property(row => row.IncNam).HasColumnName("IncNam");
        stsOdlEndResult.Property(row => row.IncCod).HasColumnName("IncCod");
        stsOdlEndResult.Property(row => row.SbsNum).HasColumnName("SbsNum");
        stsOdlEndResult.Property(row => row.OdlAmt).HasColumnName("OdlAmt");
    }
}

public sealed class SbsInfRow
{
    public long? SbsNum { get; set; }
    public string? SbsAccDte { get; set; }
}

public sealed class OdrInfRow
{
    public string? OdrDtm { get; set; }
    public string? OdrOcmNum { get; set; }
    public string? OdrTrnDep { get; set; }
    public string? OdrDelFlg { get; set; }
}

public sealed class OcmInfRow
{
    public string? OcmNum { get; set; }
    public string? OcmPatTyp { get; set; }
    public string? OcmChtNum { get; set; }
    public string? OcmAcpDtm { get; set; }
    public string? OcmAgeOld { get; set; }
    public string? OcmDgsNfs { get; set; }
    public string? OcmVstNum { get; set; }
    public string? OcmComStt { get; set; }
    public string? OcmDepCod { get; set; }
}

public sealed class OcmHstRow
{
    public string? OcmChtNum { get; set; }
    public string? OcmAcpDtm { get; set; }
    public string? OcmAgeOld { get; set; }
    public string? OcmDgsNfs { get; set; }
    public string? OcmVstNum { get; set; }
    public string? OcmComStt { get; set; }
    public string? OcmPatTyp { get; set; }
}

public sealed class PbsInfRow
{
    public string? PbsChtNum { get; set; }
    public string? PbsSexTyp { get; set; }
}

public sealed class LbqInfRow
{
    public string? LbqOcmNum { get; set; }
    public string? LbqRsvDte { get; set; }
    public string? LbqTrnDep { get; set; }
    public string? LbqChtNum { get; set; }
}

public sealed class StsOdlEndRow
{
    public string? IncNam { get; set; }
    public short? IncCod { get; set; }
    public long? SbsNum { get; set; }
    public decimal? OdlAmt { get; set; }
}

public sealed record FinanceDailyRow(string DateKey, decimal Opd, decimal Ipd)
{
    public DateTime Date => DateTime.ParseExact(DateKey, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
}
