using L3S.FromExcelToListClass.CustomValidators;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics.CodeAnalysis;

namespace L3S.FromExcelToListClass.Models
{
    public class PublicacionExcelDTO
    {
        [Required(AllowEmptyStrings = true)]
        [MaxLength(50)]
        [AllowNull]
        public string Nombre { get; set; }

        [Required]
        public string Codigo { get; set; }

        [Required]
        public string Talle { get; set; }

        [Required]
        public string Moneda { get; set; }

        [Required]
        public double Precio { get; set; }

        [Required]
        public string CantMinAplicacion { get; set; }

        [Required]
        public DateTime FechaDesde { get; set; }

        [Required]
        public string Proveedor { get; set; }

        [Required]
        public string Provincia { get; set; }

        [RequiredButNullable]
        public string Zona { get; set; }

        [Required]
        [CustomValidBoolEntry("Habilitada", "Deshabilitada")]
        public bool Estado { get; set; }

    }

}
