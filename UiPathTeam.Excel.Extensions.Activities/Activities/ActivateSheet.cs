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


    public class ActivateSheet : CodeActivity
    {
        [Description("Please enter sheet name that you want to activate.")]
        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> SheetName { get; set; }

        [DefaultValue(true)]
        [Description("If Sheet doesn't exist, it will create one")]
        [Category("Options")]
        public bool CreateNewSheet { get; set; }

        public ActivateSheet()
        {
            CreateNewSheet = true;
            Constraints.Add(CheckParentConstraint.GetCheckParentConstraint<ActivateSheet>(typeof(ExcelExtensionScope).Name));
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
                excelProperty.worksheet = (Worksheet)excelProperty.workbook.Sheets[SheetName.Get(context)];
                excelProperty.worksheet.Activate();
                if (excelProperty.save)
                {
                    excelProperty.workbook.Save();
                }
            }
            else
            {
                if (CreateNewSheet)
                {
                    Worksheet sheet = (Worksheet)excelProperty.workbook.Worksheets.Add();
                    sheet.Name = SheetName.Get(context);
                    sheets.Add(SheetName.Get(context));
                    excelProperty.worksheet = sheet;
                    sheet.Activate();
                }
                else
                    throw new Exception("Sheet Name " + SheetName.Get(context) + " was not found");
              
            }
            if (excelProperty.save)
            {
                excelProperty.workbook.Save();
            }

        }
    }
}
