using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace FoodSafetyDashboard.Data.Models;

[Table("store")]
public class Store
{
    [Key]
    [Column("store_id")]
    public int StoreId { get; set; }

    [Column("store_name")]
    public string? StoreName { get; set; }

    [Column("store_type")]
    public string? StoreType { get; set; }
}
