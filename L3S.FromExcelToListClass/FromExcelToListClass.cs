using L3S.FromExcelToListClass.CustomValidators;
using L3S.FromExcelToListClass.Enums;
using L3S.FromExcelToListClass.Interface;
using L3S.FromExcelToListClass.Models.DTO;
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using System.Reflection.Metadata;
using System.Text;


namespace CambioEstadoMasivoDePublicaciones.FromExcelToListClass
{
    public class FromExcelToListClass<T> where T : class, IExcelEntity
    {
        #region Properties
        //This is horrible, change it accesing the IExcelEntity somehow
        public static readonly PropertyInfo[] Properties = typeof(T).GetProperties().Where(x => x.Name != "Row" && x.Name != "Error" && x.Name != "TotalErrors").ToArray();
        public static readonly PropertyInfo Row = typeof(T).GetProperty("Row");
        public static readonly PropertyInfo Error = typeof(T).GetProperty("Error");
        public static readonly PropertyInfo TotalErrors = typeof(T).GetProperty("TotalErrors");

        #endregion Properties

        #region Queue Dict
        public static readonly Queue<Dictionary<string, string>> QueueDictionary = new();
        private static Dictionary<string, string> GetDictionaryFromPool()
        {
            return QueueDictionary.Count > 0 ? QueueDictionary.Dequeue() : new Dictionary<string, string>();
        }
        private static void ReturnDictionaryToPool(Dictionary<string, string> dict)
        {
            dict.Clear();
            QueueDictionary.Enqueue(dict);
        }
        #endregion Queue Dict

        public List<T> ParseExcelToClass(FileInfo file)
        {
            var ListResultT = new List<T>();
            var FirstValidations = new PropResultDTO<T>(Activator.CreateInstance<T>());

            #region Validaciones

            #region Tipo de Archivo

            var sheetResult = GetWorkbookType(file, file.Extension);

            if (sheetResult.Error)
            {
                Error.SetValue(FirstValidations.Objeto, Errors.FileError);
                ListResultT.Add(FirstValidations.Objeto);
                return ListResultT;
            }

            #endregion Tipo de Archivo

            #region Headers

            ISheet sheet = sheetResult.Sheet;
            IRow headerRow = sheet.GetRow(0);

            var requiredHeaders = GetRequiredHeaders();
            var checkHeaders = CheckHeaders(headerRow, requiredHeaders);

            if (checkHeaders.Error)
            {
                TotalErrors.SetValue(FirstValidations.Objeto, checkHeaders.ErrorMessage);
                Error.SetValue(FirstValidations.Objeto, Errors.HeaderError);
                ListResultT.Add(FirstValidations.Objeto);
                return ListResultT;
            }

            #endregion Headers

            #endregion Validaciones

            for (int i = 1; i <= 1000; i++)
            {
                var row = sheet.GetRow(i);

                if (row == null) continue;

                var cellResult = ParseRowToClass(row, headerRow, requiredHeaders.Length, headerRow.LastCellNum);

                Row.SetValue(cellResult.Objeto, i);

                if (cellResult.Error)
                {
                    Error.SetValue(cellResult.Objeto, cellResult.Objeto.Error);
                    TotalErrors.SetValue(cellResult.Objeto, cellResult.Objeto.TotalErrors);
                    ListResultT.Add(cellResult.Objeto);
                    continue;
                }

                ListResultT.Add(cellResult.Objeto);
            }

            return ListResultT;
        }

        public static ResultDTO CheckHeaders(IRow headerRow, string[] requiredHeaders)
        {
            var result = new ResultDTO() { Error = false, ErrorMessage = string.Empty };
            var msgErrorHeaders = new StringBuilder();

            var checkRequiredHeaders = CheckHeadersName(headerRow, requiredHeaders);

            if (checkRequiredHeaders.Error)
            {
                result.Error = true;
                msgErrorHeaders.Append(checkRequiredHeaders.ErrorMessage);
            }

            var notRequiredHeaders = GetNotRequiredHeaders(requiredHeaders.Length, headerRow);
            var checkNotRequiredHeaders = CheckNotRequiredHeaders(notRequiredHeaders);

            if (checkNotRequiredHeaders.Error)
            {
                result.Error = true;
                msgErrorHeaders.Append(checkNotRequiredHeaders.ErrorMessage);
            }

            result.ErrorMessage = msgErrorHeaders.ToString();
            return result;
        }

        private static List<string> GetNotRequiredHeaders(int cantReqHeader, IRow headerRow)
        {
            var auxList = new List<string>();
            for (int i = cantReqHeader; i <= headerRow.LastCellNum; i++)
            {
                ICell cell = headerRow.GetCell(i);
                ICell nextCell = headerRow.GetCell(i + 1);

                if ((cell == null || cell?.ToString() == string.Empty) && (nextCell == null || nextCell?.ToString() == string.Empty))
                {
                    return auxList;
                }
                else if ((cell == null || cell?.ToString() == string.Empty) && (nextCell != null || nextCell?.ToString() != string.Empty))
                {
                    return null;
                }

                auxList.Add(cell?.ToString());
            }

            return auxList;
        }

        private static string[] GetRequiredHeaders()
        {
            return Properties
                .Where(prop => Attribute.IsDefined(prop, typeof(RequiredAttribute)) || Attribute.IsDefined(prop, typeof(RequiredButNullableAttribute)))
                .Select(prop => prop.Name)
                .ToArray();
        }

        private static SheetResultDTO GetWorkbookType(FileInfo file, string extension)
        {
            var result = new SheetResultDTO() { Error = false, ErrorMessage = string.Empty };

            using (var stream = new FileStream(file.FullName, FileMode.Open, FileAccess.Read))
            {
                if (extension == ".xls")
                {
                    result.Sheet = new HSSFWorkbook(stream).GetSheetAt(0);
                    return result;
                }
                else if (extension == ".xlsx")
                {
                    result.Sheet = new XSSFWorkbook(stream).GetSheetAt(0);
                    return result;
                }

                result.Error = true;

                return result;
            }
        }

        private static ResultDTO CheckHeadersName(IRow row, string[] requiredHeaders)
        {
            var result = new ResultDTO() { Error = false, ErrorMessage = string.Empty };
            var msgError = new StringBuilder();

            for (int i = 0; i < requiredHeaders.Length; i++)
            {
                var cellValue = row.GetCell(i)?.StringCellValue ?? string.Empty;
                if (cellValue != requiredHeaders[i])
                {
                    result.Error = true;
                    msgError.AppendFormat("Nombre inválido en columna {0}: se esperaba '{1}', pero se encontró '{2}'.\n", i + 1, requiredHeaders[i], cellValue);
                }
            }
            result.ErrorMessage = msgError.ToString();
            return result;
        }

        public static ResultDTO CheckNotRequiredHeaders(List<string> notRequiredHeadersName)
        {
            var result = new ResultDTO() { Error = false, ErrorMessage = string.Empty };

            if (notRequiredHeadersName == null)
            {
                result.Error = true;
                result.ErrorMessage = "Corrobore que no haya encabezados o columnas en blanco intermedias.";
                return result;
            }

            var duplicates = notRequiredHeadersName.GroupBy(x => x)
                        .Where(g => g.Count() > 1)
                        .Select(g => g.Key)
                        .ToList();

            if (duplicates.Any())
            {
                result.Error = true;
                result.ErrorMessage = "Hay columnas NO OBLIGATORIAS con el mismo nombre. Porfavor revise y vuelva a intentar.";
            }

            return result;
        }

        private static PropResultDTO<T> ParseRowToClass(IRow row, IRow headerRow, int cantidadColumnasObligatorias, int cantidadColumnasRecibidas)
        {
            var result = new PropResultDTO<T>() { Error = false, ErrorMessage = string.Empty };
            var GenericObj = Activator.CreateInstance<T>();
            var msgError = new StringBuilder();

            #region Columnas Obligatorias

            for (int j = 0; j < cantidadColumnasObligatorias; j++)
            {
                PropertyInfo propiedad = Properties[j];
                ICell celda = row.GetCell(j);

                var parsedCell = ParseCellToType(celda, propiedad);

                Row.SetValue(GenericObj, row.RowNum + 1);

                if (parsedCell.Error)
                {
                    result.Error = parsedCell.Error;
                    Error.SetValue(GenericObj, parsedCell.Property);
                    msgError.AppendFormat(parsedCell.ErrorMessage, celda?.ColumnIndex + 1 ?? j + 1);
                    continue;
                }

                propiedad.SetValue(GenericObj, parsedCell.Property);
            }

            #endregion Columnas Obligatoias

            #region Columnas Opcionales

            if (cantidadColumnasObligatorias < cantidadColumnasRecibidas)
            {
                var dict = GetDictionaryFromPool();
                var propiedadDict = Properties.First(prop =>
                    prop.PropertyType == typeof(Dictionary<string, string>) ||
                    prop.PropertyType == typeof(IDictionary<string, string>));

                for (int i = cantidadColumnasObligatorias; i < cantidadColumnasRecibidas; i++)
                {
                    ICell header = headerRow.GetCell(i);
                    ICell celda = row.GetCell(i);

                    dict.Add(header.StringCellValue, celda?.ToString() ?? string.Empty);
                }

                propiedadDict.SetValue(GenericObj, new Dictionary<string, string>(dict));

                ReturnDictionaryToPool(dict);
            }

            #endregion Columnas Opcionales

            TotalErrors.SetValue(GenericObj, msgError.ToString());
            result.Objeto = GenericObj;
            return result;
        }

        private static PropertyDTO ParseCellToType(ICell celda, PropertyInfo property)
        {
            var result = new PropertyDTO() { Error = false, ErrorMessage = string.Empty };
            object parsedValue = null;
            string cellStringValue = celda?.ToString();

            var validationErrors = CheckValueWithValidationAttributes(property, cellStringValue);

            if (validationErrors.Error)
            {
                result.Property = Errors.ValidationError;
                result.ErrorMessage += "No se pudo VALIDAR la celda OBLIGATORIA COL: {0}:\n" + validationErrors.ErrorMessage;
                result.Error = true;
                return result;
            }

            switch (Type.GetTypeCode(property.PropertyType))
            {
                case TypeCode.DateTime:
                    if (DateTime.TryParse(cellStringValue, out DateTime date))
                    {
                        parsedValue = date;
                    }
                    break;

                case TypeCode.Int32:
                    if (int.TryParse(cellStringValue, out int num))
                    {
                        parsedValue = num;
                    }
                    break;

                case TypeCode.Int64:
                    if (int.TryParse(cellStringValue, out int num64))
                    {
                        parsedValue = num64;
                    }
                    break;

                case TypeCode.Boolean:
                    if (Attribute.IsDefined(property, typeof(CustomValidBoolEntryAttribute)))
                    {
                        if (property.GetCustomAttribute<CustomValidBoolEntryAttribute>().BoolCustomTryParse(cellStringValue, out bool parsedBool))
                        {
                            parsedValue = parsedBool;
                        }
                        break;
                    }

                    if (bool.TryParse(cellStringValue, out bool Bool))
                    {
                        parsedValue = Bool;
                    }
                    break;

                case TypeCode.Decimal:
                    if (decimal.TryParse(cellStringValue, out decimal Decimal))
                    {
                        parsedValue = Decimal;
                    }
                    break;

                case TypeCode.Double:
                    if (double.TryParse(cellStringValue, out double Double))
                    {
                        parsedValue = Double;
                    }
                    break;

                default:
                    if (property.PropertyType == typeof(DateOnly))
                    {
                        if (DateTime.TryParse(cellStringValue, out DateTime dateOnlyValue))
                        {
                            parsedValue = DateOnly.FromDateTime(dateOnlyValue);
                        }
                    }

                    else
                    {
                        if (!string.IsNullOrEmpty(cellStringValue))
                        {
                            parsedValue = Convert.ChangeType(cellStringValue, property.PropertyType);
                        }
                        else
                        {
                            parsedValue = Convert.ChangeType(null, property.PropertyType);
                        }
                    }
                    break;
            }

            if ((parsedValue == null || string.IsNullOrEmpty(cellStringValue)) && !Attribute.IsDefined(property, typeof(RequiredButNullableAttribute)))
            {
                result.Property = Errors.ParseError;
                result.ErrorMessage += "No se pudo PROCESAR la celda OBLIGATORIA COL: {0}\n";
                result.Error = true;
                return result;
            }

            result.Property = parsedValue;
            return result;
        }

        private static ResultDTO CheckValueWithValidationAttributes(PropertyInfo property, object value)
        {
            var result = new ResultDTO() { Error = false, ErrorMessage = "" };

            var validadores = property.GetCustomAttributes()
            .OfType<ValidationAttribute>()
            .ToList();

            if (!validadores.Any())
            {
                return result;
            }

            foreach (var validador in validadores)
            {
                if (!validador.IsValid(value))
                {
                    result.Error = true;
                    result.ErrorMessage += validador.ErrorMessage + "\n";
                }
            }

            return result;
        }
    }
}
