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

    public class GetSheets : CodeActivity
    {
        [Description("All the sheets that are present in the file")]
        [Category("Output")]
        [RequiredArgument]
        public OutArgument<List<string>> Sheets { get; set; }
        public GetSheets()
        {
            Constraints.Add(CheckParentConstraint.GetCheckParentConstraint<GetSheets>(typeof(ExcelExtensionScope).Name));
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
            if (excelProperty.save)
            {
                excelProperty.workbook.Save();
            }
            Sheets.Set(context, sheets);
        }
    }
}
