using System.ComponentModel.DataAnnotations;

namespace ImportExcelApp.Models
{
    public class Questions
    {
        public int ID { get; set; }

        [Required]
        public string? Question { get; set; }

        [Required]
        public string? RequirementID { get; set; }

        [Required]
        public bool? QuestionType { get; set; }

        [Required]
        public double? Score { get; set; }

        [Required]
        public bool? Required { get; set; }

        [Required]
        public bool? Explanation { get; set; }

        [Required]
        public bool? Attachment { get; set; }

        [Required]
        public string? RefNumber { get; set; }
    }
}
