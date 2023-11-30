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
    public class GetActiveCell : CodeActivity
    {
        [Description("Cell name of the active cell")]
        [Category("Output")]
        [RequiredArgument]
        public OutArgument<string> CellName { get; set; }

        public GetActiveCell()
        {
            Constraints.Add(CheckParentConstraint.GetCheckParentConstraint<GetActiveCell>(typeof(ExcelExtensionScope).Name));

        }
        protected override void Execute(CodeActivityContext context)
        {

            var property = context.DataContext.GetProperties()[ExcelExtensionScope.ExcelTag];
            var excelProperty = property.GetValue(context.DataContext) as ExcelSession;

            CellName.Set(context, excelProperty.application.ActiveCell.Address);
            if (excelProperty.save)
            {
                excelProperty.workbook.Save();
            }
        }
    }
}
