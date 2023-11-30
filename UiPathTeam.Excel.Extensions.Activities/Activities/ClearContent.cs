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
    public class ClearContent : CodeActivity
    {
        [Category("Options")]
        [Description("Select all, if unchecked, then it will perfrom action on the selected range")]
        public bool All { get; set; }
        
        [Category("Options")]
        [Description("It will clear the content along with the format. If unchecked, then it will clear the content by keeping the format")]
        public bool Clear { get; set; }


        public ClearContent()
        {
            Constraints.Add(CheckParentConstraint.GetCheckParentConstraint<ClearContent>(typeof(ExcelExtensionScope).Name));
        }
        protected override void Execute(CodeActivityContext context)
        {
            var property = context.DataContext.GetProperties()[ExcelExtensionScope.ExcelTag];
            var excelProperty = property.GetValue(context.DataContext) as ExcelSession;

            Range rng = excelProperty.worksheet.Application.Selection as Range;
            if (All)
            rng= excelProperty.worksheet.get_Range("a1").EntireRow.EntireColumn;
            
            if (Clear)            
                rng.Clear();            
            else
                rng.ClearContents();
        }



    }
}
