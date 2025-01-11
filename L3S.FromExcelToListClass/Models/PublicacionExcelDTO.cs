using L3S.FromExcelToListClass.CustomValidators;
using L3S.FromExcelToListClass.Enums;
using L3S.FromExcelToListClass.Interfaces;
using System.ComponentModel.DataAnnotations;

namespace L3S.FromExcelToListClass.Models
{
    public class PublicacionExcelDTO : IExcelEntity
    {
        public int Row { get; set; }
        public Errors Error { get; set; }
        public string TotalErrors { get; set; }

        [RequiredButNullable]
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

        [Required]
        [RequiredButNullable]
        public string Zona { get; set; }

        [Required]
        [CustomValidBoolEntry("Habilitada", "Deshabilitada")]
        public bool Estado { get; set; }



    }

}
