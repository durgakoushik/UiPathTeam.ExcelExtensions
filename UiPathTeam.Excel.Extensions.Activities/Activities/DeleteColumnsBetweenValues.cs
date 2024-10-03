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
    [LocalizedDisplayName(nameof(Resources.DeleteColumnsBetweenValues_DisplayName))]
    [LocalizedDescription(nameof(Resources.DeleteColumnsBetweenValues_Description))]
    public class DeleteColumnsBetweenValues : ContinuableAsyncCodeActivity
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
        [Description("Please enter the start value.")]
        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> StartValue { get; set; }

        [Description("Please enter the end value.")]
        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> EndValue { get; set; }
        #endregion

        #region Constructors

        public DeleteColumnsBetweenValues()
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
            Microsoft.Office.Interop.Excel.Range range = (Microsoft.Office.Interop.Excel.Range)excelProperty.worksheet.UsedRange;

            string startValue = StartValue.Get(context);
            string endValue = EndValue.Get(context);

            if (String.IsNullOrWhiteSpace(startValue) || String.IsNullOrWhiteSpace(endValue))
            {
                throw new ArgumentException("Invalid startvalue or endvalue specified.");
            }

            int startColumn = -1;
            int endColumn = -1;
            _Excel.Range usedRange = excelProperty.worksheet.UsedRange;

            #endregion

            #region functionality

            // Find the first occurrence of startValue ("Old")
            foreach (_Excel.Range row in usedRange.Rows)
            {
                foreach (_Excel.Range cell in row.Columns)
                {
                    if (cell.Value2 != null && cell.Value2.ToString() == startValue)
                    {
                        startColumn = cell.Column;
                        break;
                    }
                }
                if (startColumn != -1) break; // Found "Old", stop further checking
            }

            // Find the last occurrence of endValue ("New")
            for (int i = usedRange.Rows.Count; i >= 1; i--) // Iterate through rows from bottom to top
            {
                _Excel.Range row =(_Excel.Range) usedRange.Rows[i];
                foreach (_Excel.Range cell in row.Columns)
                {
                    if (cell.Value2 != null && cell.Value2.ToString() == endValue)
                    {
                        endColumn = cell.Column;
                        break;
                    }
                }
                if (endColumn != -1) break; // Found "New", stop further checking
            }

            // Ensure valid columns are found
            if (startColumn == -1)
            {
                throw new Exception($"Start value '{startValue}' not found in the sheet.");
            }

            if (endColumn == -1)
            {
                throw new Exception($"End value '{endValue}' not found in the sheet.");
            }

            if (startColumn >= endColumn)
            {
                throw new Exception("The start value occurs after or on the same column as the end value.");
            }

            // Deleting columns between startColumn and endColumn
            _Excel.Range columnsToDelete = excelProperty.worksheet.Range[excelProperty.worksheet.Cells[1, startColumn + 1], excelProperty.worksheet.Cells[1, endColumn - 1]].EntireColumn;
            columnsToDelete.Delete();


            #endregion

            Console.WriteLine("Columns between values deleted successfully.");

            if (excelProperty.save)
            {
                excelProperty.workbook.Save();
            }

            return (ctx) => {
            };
        }

        #endregion
    }
}

