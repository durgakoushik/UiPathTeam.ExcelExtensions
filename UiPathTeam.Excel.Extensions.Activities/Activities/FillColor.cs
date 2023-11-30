using Microsoft.Office.Interop.Excel;
using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Range = Microsoft.Office.Interop.Excel.Range;
namespace UiPathTeam.Excel.Extensions.Activities
{

    public class FillColor : CodeActivity
    {

        [Category("Input")]
        [Description("Select the color that needs to be apply. It will apply on the selected range or active cell.")]
        public System.Drawing.Color Color { get; set; }

        public FillColor()
        {
            Constraints.Add(CheckParentConstraint.GetCheckParentConstraint<FillColor>(typeof(ExcelExtensionScope).Name));

        }
        protected override void Execute(CodeActivityContext context)
        {
            var property = context.DataContext.GetProperties()[ExcelExtensionScope.ExcelTag];
            var excelProperty = property.GetValue(context.DataContext) as ExcelSession;
            Range rng = (Microsoft.Office.Interop.Excel.Range)excelProperty.worksheet.Application.Selection;           

            rng.Interior.Color = System.Drawing.ColorTranslator.ToOle(Color);
            if (excelProperty.save)
            {
                excelProperty.workbook.Save();
            }
        }
    }
}
