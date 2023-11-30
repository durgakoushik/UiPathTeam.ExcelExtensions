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

    public class WriteCell : CodeActivity
    {
        [Description("If left blank then it will write the value in the active cell")]
        [Category("Input")]
        public InArgument<string> CellName { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> Value { get; set; }

        public WriteCell()
        {
            Constraints.Add(CheckParentConstraint.GetCheckParentConstraint<WriteCell>(typeof(ExcelExtensionScope).Name));

        }

        protected override void Execute(CodeActivityContext context)
        {
            var property = context.DataContext.GetProperties()[ExcelExtensionScope.ExcelTag];
            var excelProperty = property.GetValue(context.DataContext) as ExcelSession;
            string cellName = CellName.Get(context);
            if (string.IsNullOrEmpty(cellName))
            {
                WrtCell(excelProperty.worksheet, Value.Get(context));
            }
            else
                WrtCell(excelProperty.worksheet, Value.Get(context), cellName);
            if (excelProperty.save)
            {
                excelProperty.workbook.Save();
            }
        }
        static void WrtCell(Worksheet ws, string value, string cellName = null)
        {
            Microsoft.Office.Interop.Excel.Range rng;
            if (string.IsNullOrEmpty(cellName))
                rng = (Range)ws.Application.ActiveCell;
            else
                rng = ws.get_Range(cellName, cellName);


            rng.Activate();
            rng.Value = value;

        }
    }
}
