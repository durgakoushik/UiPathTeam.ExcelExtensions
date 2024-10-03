using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using UiPathTeam.Excel.Extensions.Activities.Properties;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;
using System.ComponentModel;
using _Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace UiPathTeam.Excel.Extensions.Activities
{
    [LocalizedDisplayName(nameof(Resources.ConvertColumnToValue_DisplayName))]
    [LocalizedDescription(nameof(Resources.ConvertColumnToValue_Description))]
    public class ConvertColumnToValue : ContinuableAsyncCodeActivity
    {
        #region Properties

        /// <summary>
        /// If set, continue executing the remaining activities even if the current activity has failed.
        /// </summary>
        [LocalizedCategory(nameof(Resources.Common_Category))]
        [LocalizedDisplayName(nameof(Resources.ContinueOnError_DisplayName))]
        [LocalizedDescription(nameof(Resources.ContinueOnError_Description))]
        public override InArgument<bool> ContinueOnError { get; set; }

        #endregion

        #region InputArguments
        [Description("Please enter the column name")]
        [Category("Input")]
        [RequiredArgument]

        public InArgument<string> ColumnName { get; set; }
        #endregion

        #region Constructors

        public ConvertColumnToValue()
        {
        }

        #endregion


        #region Protected Methods

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {

            base.CacheMetadata(metadata);
        }

        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, CancellationToken cancellationToken)
        {
            #region INIT
            var property = context.DataContext.GetProperties()[ExcelExtensionScope.ExcelTag];
            var excelProperty = property.GetValue(context.DataContext) as ExcelSession;
            Microsoft.Office.Interop.Excel.Range usedRange = (Microsoft.Office.Interop.Excel.Range)excelProperty.worksheet.UsedRange;

            if (usedRange == null)
            {
                throw new Exception("The sheet contains no data.");
            }


            string columnHeader = ColumnName.Get(context);

            if (String.IsNullOrWhiteSpace(columnHeader))
            {
                throw new ArgumentException("Invalid header name.");
            }

            #endregion

            // Find the column index based on the header name
            int columnIndex = GetColumnIndexFromHeader(excelProperty.worksheet, columnHeader);
            if (columnIndex <= 0)
            {
                throw new Exception($"Column '{columnHeader}' not found.");
            }

            // Get the range for the entire column
            _Excel.Range columnRange = excelProperty.worksheet.Columns[columnIndex] as _Excel.Range;

            if (columnRange == null)
            {
                throw new Exception($"Column '{columnHeader}' not found.");
            }

            // Select the column
            columnRange.Select();

            // Copy the entire column (including formulas)
            columnRange.Copy();

            // Paste the values over the copied formulas
            columnRange.PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

            // Clear the clipboard after the paste operation
            excelProperty.application.CutCopyMode = XlCutCopyMode.xlCopy;

            Console.WriteLine("Completed removing the formula from the column.");

            return (ctx) => {
            };
        }

        #endregion

        static int GetColumnIndexFromHeader(Worksheet worksheet, string columnHeader)
        {
            int rowCount = worksheet.UsedRange.Rows.Count;
            int columnCount = worksheet.UsedRange.Columns.Count;

            for (int col = 1; col <= columnCount; col++)
            {
                for (int row = 1; row <= rowCount; row++)
                {
                    _Excel.Range cell = worksheet.Cells[row, col] as _Excel.Range;

                    if (cell != null && cell.Value != null && cell.Value.ToString() == columnHeader)
                    {
                        return col; // Return the column index of the matching header
                    }
                }
            }
            return -1; // Return -1 if the column is not found
        }

    }
}

