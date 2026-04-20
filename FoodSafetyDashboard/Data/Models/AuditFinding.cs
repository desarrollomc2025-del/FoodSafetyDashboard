using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace FoodSafetyDashboard.Data.Models;

[Table("audit_findings")]
public class AuditFinding
{
    [Key]
    [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
    public long Id { get; set; }

    [Column("audit_id")]
    public long AuditId { get; set; }

    [Column("section_name")]
    public string? SectionName { get; set; }

    [Column("question_text")]
    public string? QuestionText { get; set; }

    [Column("answer_value")]
    public string? AnswerValue { get; set; }

    [Column("points_earned")]
    public int? PointsEarned { get; set; }

    [Column("points_possible")]
    public int? PointsPossible { get; set; }

    [Column("finding_type")]
    public string? FindingType { get; set; }

    [Column("comment_text")]
    public string? CommentText { get; set; }

    [Column("evidence_page")]
    public int? EvidencePage { get; set; }

    [ForeignKey(nameof(AuditId))]
    public Audit? Audit { get; set; }
}
