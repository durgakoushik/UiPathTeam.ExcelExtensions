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
    public class GetLastRow : CodeActivity
    {
        [Description("Last row number of the used range")]
        [Category("Output")]
        [RequiredArgument]
        public OutArgument<int> RowNumber { get; set; }


        //private ExcelSession _excelProperty;

        public GetLastRow()
        {
            Constraints.Add(CheckParentConstraint.GetCheckParentConstraint<GetLastRow>(typeof(ExcelExtensionScope).Name));

        }
        //public GetLastRow(ExcelSession excelProperty) : this()
        //{
            
        //    this._excelProperty = excelProperty;
        //}
        protected override void Execute(CodeActivityContext context)
        {
            //if(this._excelProperty == null) {
            //    var property = context.DataContext.GetProperties()[ExcelExtensionScope.ExcelTag];
            //    //var excelProperty = property.GetValue(context.DataContext) as ExcelSession;
            //    this._excelProperty = property.GetValue(context.DataContext) as ExcelSession;
            //}
            //else
            //{

            //}

            var property = context.DataContext.GetProperties()[ExcelExtensionScope.ExcelTag];
            var excelProperty = property.GetValue(context.DataContext) as ExcelSession;

            //Worksheet ws = (Worksheet)excelProperty.workbook.Sheets[SheetName.Get(context)];
            Worksheet ws = excelProperty.worksheet;
            Range last = ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            //  Range range = ws.get_Range("A1", last);

            int lastUsedRow = last.Row;

            RowNumber.Set(context, lastUsedRow);
        }
    }
}
