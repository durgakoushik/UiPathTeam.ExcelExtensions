using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace UiPathTeam.Excel.Extensions.Activities
{
    public class SetWidthAndHeight : CodeActivity
    {
        [Category("Input")]
        public InArgument<int> ColumnWidth { get; set; }
        [Category("Input")]
        public InArgument<int> RowHeight { get; set; }
        public SetWidthAndHeight()
        {
            Constraints.Add(CheckParentConstraint.GetCheckParentConstraint<SetWidthAndHeight>(typeof(ExcelExtensionScope).Name));
        }

        protected override void Execute(CodeActivityContext context)
        {

            var property = context.DataContext.GetProperties()[ExcelExtensionScope.ExcelTag];
            var excelProperty = property.GetValue(context.DataContext) as ExcelSession;
            excelProperty.worksheet.Cells.Select();
            Microsoft.Office.Interop.Excel.Range range = (Microsoft.Office.Interop.Excel.Range)excelProperty.application.Selection;
            if (!string.IsNullOrEmpty(ColumnWidth.Get(context).ToString()))
                range.ColumnWidth = ColumnWidth.Get(context);
            if (!string.IsNullOrEmpty(RowHeight.Get(context).ToString()))
                range.RowHeight = RowHeight.Get(context);
        }
    }
}
