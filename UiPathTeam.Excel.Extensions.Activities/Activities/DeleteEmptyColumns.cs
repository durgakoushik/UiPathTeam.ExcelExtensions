using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using UiPathTeam.Excel.Extensions.Activities.Properties;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;
using Microsoft.Office.Interop.Excel;
using _Excel= Microsoft.Office.Interop.Excel;

namespace UiPathTeam.Excel.Extensions.Activities
{
    [LocalizedDisplayName(nameof(Resources.DeleteEmptyColumns_DisplayName))]
    [LocalizedDescription(nameof(Resources.DeleteEmptyColumns_Description))]
    public class DeleteEmptyColumns : ContinuableAsyncCodeActivity
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


        #region Constructors

        public DeleteEmptyColumns()
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
            Microsoft.Office.Interop.Excel.Range range = (Microsoft.Office.Interop.Excel.Range)excelProperty.application.Selection;

            #endregion

            // Loop through columns in reverse order to delete empty columns
            for (int column = range.Columns.Count; column >= 1; column--)
            {
                _Excel.Range columnRange = (Microsoft.Office.Interop.Excel.Range)range.Columns[column];
                bool isEmpty = true;

                foreach (_Excel.Range cell in columnRange.Cells)
                {
                    if (!string.IsNullOrEmpty(cell.Value?.ToString()))
                    {
                        isEmpty = false;
                        break;
                    }
                }

                if (isEmpty)
                {
                    columnRange.EntireColumn.Delete();
                }
            }

            Console.WriteLine("Empty columns deleted successfully.");

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

