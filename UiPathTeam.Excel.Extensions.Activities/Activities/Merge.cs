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
     public class Merge : CodeActivity
    {
        [Category("Input")]
        [Description("Please enter the range that you want to merge or active selection will be used.")]
        public InArgument<string> Range { get; set; }

        [Category("Option")]
        [Description("If you want to unmerge the range")]
        public bool  UnMerge { get; set; }

        public Merge()
        {
            Constraints.Add(CheckParentConstraint.GetCheckParentConstraint<Merge>(typeof(ExcelExtensionScope).Name));
        }
        protected override void Execute(CodeActivityContext context)
        {
            var property = context.DataContext.GetProperties()[ExcelExtensionScope.ExcelTag];
            var excelProperty = property.GetValue(context.DataContext) as ExcelSession;

            MergeRange(excelProperty, Range.Get(context));
        }
        void MergeRange(ExcelSession excelProperty, string range = null)
        {
            Worksheet ws = excelProperty.worksheet;
            Microsoft.Office.Interop.Excel.Range rng = ws.Application.Selection as Range;
            if (!string.IsNullOrEmpty(range))
            {
                rng = ws.Range[range.Split(':')[0], range.Split(':')[range.Split(':').Count() - 1]];
            }
            rng.Select();
            rng.Activate();
            if (!UnMerge) {
                excelProperty.application.DisplayAlerts = false;
                rng.Merge();
                excelProperty.application.DisplayAlerts = true;
            }
            else
                rng.UnMerge();

            if (excelProperty.save)
            {
                excelProperty.workbook.Save();
            }


        }
    }
}
