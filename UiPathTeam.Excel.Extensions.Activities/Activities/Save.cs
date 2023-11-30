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
    public class SaveWorkbook : CodeActivity
    {

        public SaveWorkbook()
        {
            Constraints.Add(CheckParentConstraint.GetCheckParentConstraint<SaveWorkbook>(typeof(ExcelExtensionScope).Name));

        }
        protected override void Execute(CodeActivityContext context)
        {
            var property = context.DataContext.GetProperties()[ExcelExtensionScope.ExcelTag];
            var excelProperty = property.GetValue(context.DataContext) as ExcelSession;
            excelProperty.workbook.Save();
        }


    }
}
