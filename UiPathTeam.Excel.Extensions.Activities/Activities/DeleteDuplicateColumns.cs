using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using UiPathTeam.Excel.Extensions.Activities.Properties;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;
using System.ComponentModel;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace UiPathTeam.Excel.Extensions.Activities
{
    [LocalizedDisplayName(nameof(Resources.DeleteDuplicateColumns_DisplayName))]
    [LocalizedDescription(nameof(Resources.DeleteDuplicateColumns_Description))]
    public class DeleteDuplicateColumns : ContinuableAsyncCodeActivity
    {
        #region Properties

        /// <summary>
        /// If set, continue executing the remaining activities even if the current activity has failed.
        /// </summary>
        [LocalizedCategory(nameof(Resources.Common_Category))]
        [LocalizedDisplayName(nameof(Resources.ContinueOnError_DisplayName))]
        [LocalizedDescription(nameof(Resources.ContinueOnError_Description))]
        public override InArgument<bool> ContinueOnError { get; set; }

        [Category("Input")]
        [Description("Index of the Header row")]
        [DisplayName("Header Row Index")]
        public InArgument<string> HeaderRowIndex { get; set; }

        #endregion


        #region Constructors

        public DeleteDuplicateColumns()
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
            Microsoft.Office.Interop.Excel.Range rangeVal = (Microsoft.Office.Interop.Excel.Range)excelProperty.application.Selection;

            int headerRowIndex = Convert.ToInt32(HeaderRowIndex.Get(context));

            #endregion

            #region HeaderRowIndex Validation
            // Check if the header row value is invalid or empty 

            if (headerRowIndex == 0)
            {
                headerRowIndex = 1;
            }
            else
            {
                bool successfulconversion = Int32.TryParse(headerRowIndex.ToString(), out headerRowIndex);
                if (Convert.ToInt32(headerRowIndex) <= 0)
                {
                    throw new ArgumentException("The header row index value is leass than or equal to zero.");
                }
            }
            #endregion

            #region HeaderRowIndex Validation with Range
            if (headerRowIndex > rangeVal.Rows.Count)
            {
                throw new ArgumentException("Invalid Header Row Index Value. The Header Row Index value is greater than the Total Rows in the Given/Used Range");
            }
            #endregion

            #region MainLogic

            _Excel.Range headerRow = (Microsoft.Office.Interop.Excel.Range)rangeVal.Rows[headerRowIndex];
            
            // Array to track which columns are duplicates
            bool[] isThisADuplicateColumn = new bool[rangeVal.Columns.Count + 1];

            // Find and mark duplicate columns based on all header names
            for (int i = 1; i <= rangeVal.Columns.Count - 1; i++)
            {
                if (isThisADuplicateColumn[i])
                {

                    continue; // Skip columns already marked as duplicates
                }

                _Excel.Range currentColumn = (Microsoft.Office.Interop.Excel.Range)rangeVal.Columns[i];

                for (int j = i + 1; j <= rangeVal.Columns.Count; j++)
                {
                    _Excel.Range comparisonColumn = (Microsoft.Office.Interop.Excel.Range)rangeVal.Columns[j];

                    if (AreColumnsEqual(headerRow, currentColumn, comparisonColumn))
                    {

                        isThisADuplicateColumn[j] = true;
                    }
                }
            }

            // Delete duplicate columns
            for (int i = rangeVal.Columns.Count; i >= 1; i--)
            {
                if (isThisADuplicateColumn[i])
                {
                    Console.WriteLine(String.Format("Deleted Column with Name : {0} at Column Index : {1}", (excelProperty.worksheet.Cells[headerRowIndex, i] as _Excel.Range).Value.ToString(), i.ToString()));
                    ((Microsoft.Office.Interop.Excel.Range)(rangeVal.Columns[i])).Delete();
                }
            }

            Console.WriteLine("Activity for deleting duplicate columns executed successfully.");

            if (excelProperty.save)
            {
                excelProperty.workbook.Save();
            }


            #endregion

            return (ctx) => {
            };
        }

        #endregion

        #region AreColumnsEqualMethod
        static bool AreColumnsEqual(Microsoft.Office.Interop.Excel.Range headerRow, Microsoft.Office.Interop.Excel.Range column1, Microsoft.Office.Interop.Excel.Range column2)
        {
            // Ensure the header row and both columns are not null
            if (headerRow == null || column1 == null || column2 == null)
            {
                return false;
            }

            // Iterate through cells in the header row
            for (int i = 1; i <= headerRow.Columns.Count; i++)
            {
                var cellValue1 = ((Microsoft.Office.Interop.Excel.Range)(column1.Cells[1, 1])).Value;
                var cellValue2 = ((Microsoft.Office.Interop.Excel.Range)(column2.Cells[1, 1])).Value;

                // Compare cell values (case-insensitive comparison)
                if (!string.Equals(cellValue1.ToString(), cellValue2.ToString(), StringComparison.OrdinalIgnoreCase))
                {

                    return false; // Header names are not equal
                }
                else
                {

                    return true;
                }
            }

            return true; // Header names are equal
        }
        #endregion


    }
}

