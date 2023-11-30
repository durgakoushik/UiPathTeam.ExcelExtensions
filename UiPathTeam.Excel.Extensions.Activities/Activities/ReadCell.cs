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

    public class ReadCell : CodeActivity
    {
        [Description("If left blank then it returns active cell value")]
        [Category("Input")]
        public InArgument<string> CellName { get; set; } 

        [Category("Output")]
        [RequiredArgument]
        public OutArgument<string> Value { get; set; }

        public ReadCell()
        {
            Constraints.Add(CheckParentConstraint.GetCheckParentConstraint<ReadCell>(typeof(ExcelExtensionScope).Name));

        }
        protected override void Execute(CodeActivityContext context)
        {
            var property = context.DataContext.GetProperties()[ExcelExtensionScope.ExcelTag];
            var excelProperty = property.GetValue(context.DataContext) as ExcelSession;
            string cellName = CellName.Get(context);
            if (string.IsNullOrEmpty(cellName))
            {
                Value.Set(context, ReadCellValue(excelProperty.worksheet));
            }
            else
                Value.Set(context, ReadCellValue(excelProperty.worksheet, cellName));
            if (excelProperty.save)
            {
                excelProperty.workbook.Save();
            }

        }
         string ReadCellValue(Worksheet ws,string cellName = null)
        {
            Microsoft.Office.Interop.Excel.Range rng;
            if (string.IsNullOrEmpty(cellName))
                rng = (Range)ws.Application.ActiveCell;
            else
                rng = ws.get_Range(cellName, cellName);

            return rng.Value.ToString();
        }
    }
}
