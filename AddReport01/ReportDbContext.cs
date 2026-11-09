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
    }
}

public sealed record FinanceDailyRow(string DateKey, decimal Opd, decimal Ipd)
{
    public DateTime Date => DateTime.ParseExact(DateKey, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
}
