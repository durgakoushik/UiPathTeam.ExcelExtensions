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

namespace UiPathTeam.Excel.Extensions.Activities
{

    public class AutoFill : CodeActivity
    {
        [Category("Input")]
        [Description("The range which specifies the starting point of the fill and contains an initial value. It can be single cell 'A1' or it can be range 'A1:A2'")]
        [RequiredArgument]
        public InArgument<string> Range { get; set; }
        [Category("Input")]
        [Description("CellName which specifies the ending point of the fill.")]
        [RequiredArgument]
        public InArgument<string> CellName { get; set; }

        public AutoFill()
        {
            Constraints.Add(CheckParentConstraint.GetCheckParentConstraint<AutoFill>(typeof(ExcelExtensionScope).Name));
        }
        protected override void Execute(CodeActivityContext context)
        {
            var property = context.DataContext.GetProperties()[ExcelExtensionScope.ExcelTag];
            var excelProperty = property.GetValue(context.DataContext) as ExcelSession;

            string range = Range.Get(context);
            string cellName = CellName.Get(context);

            var Column = new String(range.Split(':')[range.Split(':').Length - 1].Where(Char.IsLetter).ToArray());

            Microsoft.Office.Interop.Excel.Range r1 = excelProperty.worksheet.Range[range.Split(':')[0], range.Split(':')[range.Split(':').Length-1]];
            Microsoft.Office.Interop.Excel.Range r2 = excelProperty.worksheet.Range[range.Split(':')[0], cellName];
            r1.AutoFill(r2);
            if (excelProperty.save)
            {
                excelProperty.workbook.Save();
            }
        }
    }
}
