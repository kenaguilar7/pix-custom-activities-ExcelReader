using System;
using BR.Core;
using BR.Core.Attributes;
using Activities.Office.Consultor;
using Activities.Office.QueryConsultor.Properties;

namespace Namespace_Office
{
    [LocalizableScreenName(nameof(Resources.ConsultaEstructuradaExcel_ScreenName), typeof(Resources))]
    [BR.Core.Attributes.Path("Office")]
    public class ConsultaEstructuradaExcel : BR.Core.Activity
    {
        [LocalizableScreenName(nameof(Resources.DocumentoExcelUrl_ScreenName), typeof(Resources))]
        [LocalizableDescription(nameof(Resources.DocumentoExcelUrl_Description), typeof(Resources))]
        [IsFilePathChooser]
        public System.String DocumentoExcelUrl {get; set;}

        [LocalizableScreenName(nameof(Resources.Query_ScreenName), typeof(Resources))]
        [LocalizableDescription(nameof(Resources.Query_Description), typeof(Resources))]
        public System.String Query { get; set; }

        [LocalizableScreenName(nameof(Resources.Resultado_ScreenName), typeof(Resources))]
        [LocalizableDescription(nameof(Resources.Resultado_Description), typeof(Resources))]
        [IsOut]
        public List<Dictionary<string, string>> Resultado {get; set;} 
        
        public override void Execute(int? optionID)
        {
            IExcelQueryConverter converter = new ExcelQueryConverter(DocumentoExcelUrl);
            converter = new ValidateFileExist(converter);
            converter = new ValidationDecorator(converter);
            Resultado = converter.ExecuteQuery(Query);
        }
    }
}
