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
    [LocalizedDisplayName(nameof(Resources.DeleteEmptyRows_DisplayName))]
    [LocalizedDescription(nameof(Resources.DeleteEmptyRows_Description))]
    public class DeleteEmptyRows : CodeActivity
    {
        #region Properties

        /// <summary>
        /// If set, continue executing the remaining activities even if the current activity has failed.
        /// </summary>
      //  [LocalizedCategory(nameof(Resources.Common_Category))]
       // [LocalizedDisplayName(nameof(Resources.ContinueOnError_DisplayName))]
        //[LocalizedDescription(nameof(Resources.ContinueOnError_Description))]
        //public override InArgument<bool> ContinueOnError { get; set; }

        #endregion


        #region Constructors

        public DeleteEmptyRows()
        {
        }

        #endregion


        #region Protected Methods

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {

            base.CacheMetadata(metadata);
        }

        protected override void Execute(CodeActivityContext context)
        {
            #region INIT
            var property = context.DataContext.GetProperties()[ExcelExtensionScope.ExcelTag];
            var excelProperty = property.GetValue(context.DataContext) as ExcelSession;
            Microsoft.Office.Interop.Excel.Range rangeVal = (Microsoft.Office.Interop.Excel.Range)excelProperty.application.Selection;

            #endregion

            try
            {
                RemoveEmptyRows(excelProperty.application, excelProperty.worksheet, rangeVal);
                Console.WriteLine("Empty rows deleted successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception occured while deleting empty rows. Message : "+ex.Message.ToString());
                throw ex;
            }

            if (excelProperty.save)
            {
                excelProperty.workbook.Save();
            }

            //return (ctx) => {
            //};
        }

        #endregion

        static void RemoveEmptyRows(_Excel.Application excelApp, _Excel.Worksheet worksheet, _Excel.Range rangeVal)
        {
            var LastRow = rangeVal.Rows.Count;
            LastRow = LastRow + rangeVal.Row - 1;
            for (int i = LastRow; i >= 1; i--)
            {
                if (excelApp.WorksheetFunction.CountA(rangeVal.Rows[i]) == 0)
                    (worksheet.Rows[i] as Microsoft.Office.Interop.Excel.Range).Delete();
            }
        }
    }
}

