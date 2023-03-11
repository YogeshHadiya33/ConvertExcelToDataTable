using System.ComponentModel.DataAnnotations;

namespace ExcelToDataTable.Models
{
    public class FileUploadModel
    {
        [Required(ErrorMessage = "Please select file")]
        public IFormFile File { get; set; }
    }
}
