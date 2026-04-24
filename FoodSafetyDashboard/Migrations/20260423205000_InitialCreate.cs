using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace FoodSafetyDashboard.Migrations
{
    /// <inheritdoc />
    public partial class InitialCreate : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "audits",
                columns: table => new
                {
                    audit_id = table.Column<long>(type: "bigint", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    store_id = table.Column<int>(type: "int", nullable: true),
                    location = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    audit_start = table.Column<DateTime>(type: "datetime2", nullable: true),
                    audit_end = table.Column<DateTime>(type: "datetime2", nullable: true),
                    auditor = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    franchisee = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    manager_in_charge = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    points_earned = table.Column<int>(type: "int", nullable: true),
                    points_possible = table.Column<int>(type: "int", nullable: true),
                    percentage_score = table.Column<decimal>(type: "decimal(18,2)", nullable: true),
                    critical_violations = table.Column<int>(type: "int", nullable: true),
                    total_violations = table.Column<int>(type: "int", nullable: true),
                    source_file = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    departamento = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    municipio = table.Column<string>(type: "nvarchar(max)", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_audits", x => x.audit_id);
                });

            migrationBuilder.CreateTable(
                name: "audit_findings",
                columns: table => new
                {
                    Id = table.Column<long>(type: "bigint", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    audit_id = table.Column<long>(type: "bigint", nullable: false),
                    section_name = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    question_text = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    answer_value = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    points_earned = table.Column<int>(type: "int", nullable: true),
                    points_possible = table.Column<int>(type: "int", nullable: true),
                    finding_type = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    comment_text = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    evidence_page = table.Column<int>(type: "int", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_audit_findings", x => x.Id);
                    table.ForeignKey(
                        name: "FK_audit_findings_audits_audit_id",
                        column: x => x.audit_id,
                        principalTable: "audits",
                        principalColumn: "audit_id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateTable(
                name: "audit_sections",
                columns: table => new
                {
                    Id = table.Column<long>(type: "bigint", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    audit_id = table.Column<long>(type: "bigint", nullable: false),
                    section_name = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    points_earned = table.Column<int>(type: "int", nullable: true),
                    points_possible = table.Column<int>(type: "int", nullable: true),
                    section_score = table.Column<decimal>(type: "decimal(18,2)", nullable: true),
                    total_violations = table.Column<int>(type: "int", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_audit_sections", x => x.Id);
                    table.ForeignKey(
                        name: "FK_audit_sections_audits_audit_id",
                        column: x => x.audit_id,
                        principalTable: "audits",
                        principalColumn: "audit_id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateIndex(
                name: "IX_audit_findings_audit_id",
                table: "audit_findings",
                column: "audit_id");

            migrationBuilder.CreateIndex(
                name: "IX_audit_sections_audit_id",
                table: "audit_sections",
                column: "audit_id");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "audit_findings");

            migrationBuilder.DropTable(
                name: "audit_sections");

            migrationBuilder.DropTable(
                name: "audits");
        }
    }
}
