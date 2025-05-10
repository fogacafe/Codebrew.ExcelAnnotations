using Codebrew.ExcelAnnotations.Attributes;
using Codebrew.ExcelAnnotations.Attributes.Converters;

namespace Codebrew.ExcelAnnotations.Example.Worksheets
{
    [WorksheetName("Books")]
    public class BookWorksheet
    {
        [Header("Titulo")]
        public string? Title { get; set; }

        [Header("Preco")]
        public decimal? Price { get; set; }

        [Header("Data de Lancamento")]
        public DateTime Date { get; set; }

        [Header("Nota")]
        [PercentageConverter]
        public decimal Percentage { get; set; }

        [Header("Ativo")]
        [BoolConverter(["Sim", "S", "Ativo"], ["Não", "Nao", "N", "Desativado"])]
        public bool? IsActive { get; set; }
    }
}
