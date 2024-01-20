using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using UiPathTeam.Excel.Extensions.Activities.Properties;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;
using System.ComponentModel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace UiPathTeam.Excel.Extensions.Activities
{
    [LocalizedDisplayName(nameof(Resources.ConvertRangeToHTMLCode_DisplayName))]
    [LocalizedDescription(nameof(Resources.ConvertRangeToHTMLCode_Description))]
    public class ConvertRangeToHTMLCode : ContinuableAsyncCodeActivity
    {
        #region Properties

        /// <summary>
        /// If set, continue executing the remaining activities even if the current activity has failed.
        /// </summary>
        [LocalizedCategory(nameof(Resources.Common_Category))]
        [LocalizedDisplayName(nameof(Resources.ContinueOnError_DisplayName))]
        [LocalizedDescription(nameof(Resources.ContinueOnError_Description))]
        public override InArgument<bool> ContinueOnError { get; set; }

        [Category("Output")]
        [Description("HTML Code")]
        [DisplayName("HTMLCode")]
        public OutArgument<string> HTMLCode { get; set; }



        #endregion


        #region Constructors

        public ConvertRangeToHTMLCode()
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
            // Inputs

            ///////////////////////////
            // Add execution logic HERE
            ///////////////////////////

            #region INIT
            var property = context.DataContext.GetProperties()[ExcelExtensionScope.ExcelTag];
            var excelProperty = property.GetValue(context.DataContext) as ExcelSession;
            Microsoft.Office.Interop.Excel.Range rangeVal = (Microsoft.Office.Interop.Excel.Range)excelProperty.application.Selection;
            string htmlCode = "";
            #endregion

            htmlCode = ConvertToHTML(excelProperty.worksheet.UsedRange);

            

            if (excelProperty.save)
            {
                excelProperty.workbook.Save();
            }
            Console.WriteLine("Completed creating HTML Code for the active range");

            // Outputs
            return (ctx) => {
                HTMLCode.Set(ctx, htmlCode);
            };
        }

        #endregion

        #region Function
        static string ConvertToHTML(_Excel.Range range)
        {
            string html = "<table style=\"border-collapse: collapse; width: 100%;\" border=\"1\">";

            foreach (_Excel.Range row in range.Rows)
            {
                html += "<tr>";
                foreach (_Excel.Range cell in row.Cells)
                {
                    html += $"<td>{cell.Value}</td>";
                }
                html += "</tr>";
            }

            html += "</table>";

            return html;
        }
        #endregion

    }
}

