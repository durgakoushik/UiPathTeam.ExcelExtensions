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
    public class GetLastColumn : CodeActivity
    {
        [Description("Last column name of the used range")]
        [Category("Output")]
        [RequiredArgument]
        public OutArgument<String> ColumnName { get; set; }

        public GetLastColumn()
        {
            Constraints.Add(CheckParentConstraint.GetCheckParentConstraint<GetLastColumn>(typeof(ExcelExtensionScope).Name));

        }
        protected override void Execute(CodeActivityContext context)
        {
            var property = context.DataContext.GetProperties()[ExcelExtensionScope.ExcelTag];
            var excelProperty = property.GetValue(context.DataContext) as ExcelSession;

            //Worksheet ws = (Worksheet)excelProperty.workbook.Sheets[SheetName.Get(context)];
            Worksheet ws = excelProperty.worksheet;
            Range last = ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);


            int columnNumber = last.Column;

            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }
            if (excelProperty.save)
            {
                excelProperty.workbook.Save();
            }
            ColumnName.Set(context, columnName);
        }
    }
}
