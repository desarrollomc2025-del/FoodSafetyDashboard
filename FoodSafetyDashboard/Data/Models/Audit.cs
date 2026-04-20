using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace FoodSafetyDashboard.Data.Models;

[Table("audits")]
public class Audit
{
    [Key]
    [Column("audit_id")]
    public long AuditId { get; set; }

    [Column("store_id")]
    public int? StoreId { get; set; }

    [Column("location")]
    public string? Location { get; set; }

    [Column("audit_start")]
    public DateTime? AuditStart { get; set; }

    [Column("audit_end")]
    public DateTime? AuditEnd { get; set; }

    [Column("auditor")]
    public string? Auditor { get; set; }

    [Column("franchisee")]
    public string? Franchisee { get; set; }

    [Column("manager_in_charge")]
    public string? ManagerInCharge { get; set; }

    [Column("points_earned")]
    public int? PointsEarned { get; set; }

    [Column("points_possible")]
    public int? PointsPossible { get; set; }

    [Column("percentage_score")]
    public decimal? PercentageScore { get; set; }

    [Column("critical_violations")]
    public int? CriticalViolations { get; set; }

    [Column("total_violations")]
    public int? TotalViolations { get; set; }

    [Column("source_file")]
    public string? SourceFile { get; set; }

    [Column("departamento")]
    public string? Departamento { get; set; }

    [Column("municipio")]
    public string? Municipio { get; set; }

    public ICollection<AuditSection> Sections { get; set; } = new List<AuditSection>();
    public ICollection<AuditFinding> Findings { get; set; } = new List<AuditFinding>();
}
