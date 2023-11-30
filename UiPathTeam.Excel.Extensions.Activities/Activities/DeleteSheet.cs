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
using UiPathTeam.Excel.Extensions;

namespace UiPathTeam.Excel.Extensions.Activities
{
    public class DeleteSheet : CodeActivity
    {
        [Description("Name of the sheet")]
        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> SheetName { get; set; }
        public DeleteSheet()
        {
            Constraints.Add(CheckParentConstraint.GetCheckParentConstraint<DeleteSheet>(typeof(ExcelExtensionScope).Name));
        }
        protected override void Execute(CodeActivityContext context)
        {
            var property = context.DataContext.GetProperties()[ExcelExtensionScope.ExcelTag];
            var excelProperty = property.GetValue(context.DataContext) as ExcelSession;

            List<string> sheets = new List<string>();
            for (int sheetNum = 1; sheetNum < excelProperty.workbook.Sheets.Count + 1; sheetNum++)
            {
                Worksheet sheet = (Worksheet)excelProperty.workbook.Sheets[sheetNum];
                sheets.Add(sheet.Name);
            }
            if (sheets.Contains(SheetName.Get(context)))
            {
                excelProperty.application.DisplayAlerts = false;
                Worksheet sheet = (Worksheet)excelProperty.workbook.Sheets[SheetName.Get(context)];
                sheet.Delete();
                excelProperty.application.DisplayAlerts = true;
                if (excelProperty.save)
                {
                    excelProperty.workbook.Save();
                }
            }
            else
                throw new Exception("Sheet Name " + SheetName.Get(context) + " was not found");
        }

    }
}
