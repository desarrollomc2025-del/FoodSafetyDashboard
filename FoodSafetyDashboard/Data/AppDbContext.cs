using FoodSafetyDashboard.Data.Models;
using Microsoft.EntityFrameworkCore;

namespace FoodSafetyDashboard.Data;

public class AppDbContext(DbContextOptions<AppDbContext> options) : DbContext(options)
{
    public DbSet<Audit> Audits => Set<Audit>();
    public DbSet<AuditSection> AuditSections => Set<AuditSection>();
    public DbSet<AuditFinding> AuditFindings => Set<AuditFinding>();
    public DbSet<Store> Stores => Set<Store>();

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<AuditSection>()
            .HasOne(s => s.Audit)
            .WithMany(a => a.Sections)
            .HasForeignKey(s => s.AuditId);

        modelBuilder.Entity<AuditFinding>()
            .HasOne(f => f.Audit)
            .WithMany(a => a.Findings)
            .HasForeignKey(f => f.AuditId);
    }
}
