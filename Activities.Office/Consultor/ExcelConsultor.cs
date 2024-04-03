using BR.Core;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Activities.Office.Consultor
{
    internal class ExcelConsultor
    {
    }

    public interface IExcelQueryConverter
    {
        List<Dictionary<string, string>> ExecuteQuery(string query);
        string FilePath { get; }
    }

    public class ExcelQueryConverter : IExcelQueryConverter
    {
        private readonly string filePath;
        public string FilePath => filePath;

        public ExcelQueryConverter(string filePath)
        {
            this.filePath = filePath;
        }


        public List<Dictionary<string, string>> ExecuteQuery(string query)
        {
            // Extrae las partes clave de la consulta
            var selectPart = query.Substring("SELECT ".Length, query.IndexOf("FROM") - "SELECT ".Length).Trim();
            var wherePart = query.Substring(query.IndexOf("WHERE") + "WHERE".Length).Trim();
            var fromPart = new SQLQueryParser().GetFromValue(query);

            var columnsName = ExtractElements(selectPart);
            var whereCondition = wherePart.Split(new[] { " = " }, StringSplitOptions.None);
            
            
            var whereColumn = whereCondition[0].Trim();
            var whereValue = whereCondition[1].Trim('\'');

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(fromPart);

                var matchedColumn = MatchColumns(worksheet.RangeUsed());
                var columnElemenet = matchedColumn.First(x => x.Value == whereColumn);
                IEnumerable<KeyValuePair<string, string>> matchedColumnInSelect = matchedColumn.Where(x => columnsName.Contains(x.Value));

                IEnumerable<Dictionary<string, string>> queryResult = worksheet.RangeUsed().RowsUsed()
                    .Where(row => row.Cell(columnElemenet.Key).GetString() == whereValue)
                    .Select(row =>
                                matchedColumnInSelect.ToDictionary(
                                    columnPair => columnPair.Value,
                                    columnPair => row.Cell(columnPair.Key).Value.ToString()
                                ))
                    .ToList();


                return queryResult.ToList();
            }
        }

        public Dictionary<string, string> MatchColumns(IXLRange range)
        {
            var matchedColumns = new Dictionary<string, string>();
            // Asumiendo que estamos buscando coincidencias en la primera fila del rango
            var firstRow = range.FirstRow();
            foreach (var cell in firstRow.Cells())
                matchedColumns[cell.WorksheetColumn().ColumnLetter()] = cell.Value.ToString();
            return matchedColumns;
        }

        public string[] ExtractElements(string input)
        {
            if (string.IsNullOrEmpty(input))
            {
                return new string[0];
            }

            string[] elements = input.Split(',');
            for (int i = 0; i < elements.Length; i++)
            {
                elements[i] = elements[i].Trim();
            }

            return elements;
        }

    }

    public abstract class ExcelQueryConverterDecorator : IExcelQueryConverter
    {
        protected IExcelQueryConverter wrappedConverter;
        public string FilePath => wrappedConverter.FilePath;

        public ExcelQueryConverterDecorator(IExcelQueryConverter converter)
        {
            this.wrappedConverter = converter;
        }


        public virtual List<Dictionary<string, string>> ExecuteQuery(string query)
        {
            return wrappedConverter.ExecuteQuery(query);
        }
    }

    public class ValidationDecorator : ExcelQueryConverterDecorator
    {
        public ValidationDecorator(IExcelQueryConverter converter) : base(converter) { }

        public override List<Dictionary<string, string>> ExecuteQuery(string query)
        {
            ValidateQuery(query);  // Método para validar la consulta
            return base.ExecuteQuery(query);
        }

        private void ValidateQuery(string query)
        {
            // Lógica de validación
            if (string.IsNullOrWhiteSpace(query))
                throw new ArgumentException("La consulta no puede estar vacía.");
            // Otras validaciones...
        }
    }

    public class ValidateFileExist : ExcelQueryConverterDecorator
    {
        public ValidateFileExist(IExcelQueryConverter converter) : base(converter)
        {
        }
        public override List<Dictionary<string, string>> ExecuteQuery(string query)
        {
            ValidateFilePath(FilePath);  // Método para validar la consulta
            return base.ExecuteQuery(query);
        }

        private void ValidateFilePath(string filePath)
        {
            if (!File.Exists(filePath)) { throw new FileNotFoundException(filePath); }
        }
    }

    public class SQLQueryParser
    {
        public string GetFromValue(string query)
        {
            int fromIndex = query.IndexOf("FROM");
            if (fromIndex == -1)
            {
                return "No se encontró la cláusula FROM";
            }

            int fromValueStart = fromIndex + "FROM".Length;
            string afterFrom = query.Substring(fromValueStart).Trim();

            // Dividir el texto restante para obtener el primer token, que debería ser el nombre de la tabla
            string[] tokens = afterFrom.Split("WHERE", StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Length > 0)
            {
                return tokens[0].Trim();
            }
            throw new NotFromClausedFound();
        }
    }

    public class NotFromClausedFound : Exception
    {
        public NotFromClausedFound() : base()
        {

        }
    }

}
