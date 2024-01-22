using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace XLLib.Test;

public class Model
{
    [DisplayName("Name")] [Required] public string? Name { get; set; }
    public string? Description { get; set; }
    public int Age { get; set; }
    [DisplayName("Created at")] public DateTime? Date { get; set; }
    public decimal? Amount { get; set; }
    [DisplayName("Is true")] public bool? IsTrue { get; set; }
}