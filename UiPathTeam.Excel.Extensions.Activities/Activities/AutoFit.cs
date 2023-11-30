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

    public class AutoFit : CodeActivity
    {
        [Description("Column range that needs to be AutoFit. Ex: H:G. If left empty, it will autofit columns for entire excel.")]
        [Category("Optional")]
        public InArgument<string> Range { get; set; }
        public AutoFit()
        {
            Constraints.Add(CheckParentConstraint.GetCheckParentConstraint<AutoFit>(typeof(ExcelExtensionScope).Name));
        }

        protected override void Execute(CodeActivityContext context)
        {
            var property = context.DataContext.GetProperties()[ExcelExtensionScope.ExcelTag];
            var excelProperty = property.GetValue(context.DataContext) as ExcelSession;

            string range = Range.Get(context);
            if (string.IsNullOrEmpty(range))
            {
                excelProperty.worksheet.Cells.Select();
                excelProperty.worksheet.Cells.EntireColumn.AutoFit();
                excelProperty.worksheet.Rows.EntireRow.AutoFit();
            }
            else
            {
                Range rng = (Microsoft.Office.Interop.Excel.Range)excelProperty.worksheet.Columns[range];
                //if ((range.Split(':').Count()) == 1)
                //{
                //    rng = (Microsoft.Office.Interop.Excel.Range)excelProperty.worksheet.Columns[range.Split(':')[0], range.Split(':')[0]];
                //}else if ((range.Split(':').Count()) == 2)
                //{
                //    rng = (Microsoft.Office.Interop.Excel.Range)excelProperty.worksheet.Columns[range.Split(':')[0], range.Split(':')[1]];
                //}
                if(rng != null)
                {
                    rng.Select();
                    rng.EntireColumn.AutoFit();
                    rng.EntireRow.AutoFit();
                }

            }
            if (excelProperty.save)
            {
                excelProperty.workbook.Save();
            }
        }
    }
}
