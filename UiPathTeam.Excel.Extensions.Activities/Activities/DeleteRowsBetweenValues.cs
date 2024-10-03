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
    [LocalizedDisplayName(nameof(Resources.DeleteRowsBetweenValues_DisplayName))]
    [LocalizedDescription(nameof(Resources.DeleteRowsBetweenValues_Description))]
    public class DeleteRowsBetweenValues : ContinuableAsyncCodeActivity
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

        public DeleteRowsBetweenValues()
        {
        }

        #endregion

        #region MainFunctionality

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

            int startRow = -1;
            int endRow = -1;
            #endregion

            #region Functionality
            
            _Excel.Range usedRange = excelProperty.worksheet.UsedRange;

            // Find the first occurrence of startValue
            foreach (_Excel.Range row in usedRange.Rows)
            {
                foreach (_Excel.Range cell in row.Cells)
                {
                    if (cell.Value2 != null && cell.Value2.ToString() == startValue)
                    {
                        startRow = cell.Row;
                        break;
                    }
                }
                if (startRow != -1) break; // Found "Old", no need to check further
            }

            // Find the last occurrence of endValue ("New")
            for (int i = usedRange.Rows.Count; i >= 1; i--) // Iterate from bottom to top
            {
                _Excel.Range row =(_Excel.Range)usedRange.Rows[i];
                foreach (_Excel.Range cell in row.Columns) // Iterate over columns within the row
                {
                    if (cell.Value2 != null && cell.Value2.ToString() == endValue)
                    {
                        endRow = cell.Row;
                        break;
                    }
                }
                if (endRow != -1) break; // Found "New", no need to check further
            }

            // Ensure valid rows found
            if (startRow == -1)
            {
                throw new Exception($"Start value '{startValue}' not found in the sheet.");
            }

            if (endRow == -1)
            {
                throw new Exception($"End value '{endValue}' not found in the sheet.");
            }

            if (startRow >= endRow)
            {
                throw new Exception("The start value occurs after or on the same row as the end value.");
            }

            // Deleting rows between startRow and endRow
            _Excel.Range rowsToDelete = excelProperty.worksheet.Range[$"A{startRow + 1}:A{endRow - 1}"].EntireRow;
            rowsToDelete.Delete();


            #endregion



            Console.WriteLine("Rows between values deleted successfully.");

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

