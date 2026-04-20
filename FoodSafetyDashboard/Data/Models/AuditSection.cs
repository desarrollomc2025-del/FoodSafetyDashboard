using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace FoodSafetyDashboard.Data.Models;

[Table("audit_sections")]
public class AuditSection
{
    [Key]
    [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
    public long Id { get; set; }

    [Column("audit_id")]
    public long AuditId { get; set; }

    [Column("section_name")]
    public string? SectionName { get; set; }

    [Column("points_earned")]
    public int? PointsEarned { get; set; }

    [Column("points_possible")]
    public int? PointsPossible { get; set; }

    [Column("section_score")]
    public decimal? SectionScore { get; set; }

    [Column("total_violations")]
    public int? TotalViolations { get; set; }

    [ForeignKey(nameof(AuditId))]
    public Audit? Audit { get; set; }
}
