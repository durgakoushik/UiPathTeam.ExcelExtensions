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
    public class SelectRange : CodeActivity
    {
        [Category("Input")]
        [Description("Please enter the range that you want to select or select different option from properties.")]
        public InArgument<string>  Range { get; set; }
        [Category("Options")]
        [Description("Select all used range")]
        public bool All { get; set; }
        [Category("Options")]
        [Description("Select entire used row of the active cell")]
        public bool EntireRow { get; set; }
        [Category("Options")]
        [Description("Select entire used column of the active cell")]
        public bool EntireColumn { get; set; }

        public SelectRange()
        {
            Constraints.Add(CheckParentConstraint.GetCheckParentConstraint<SelectRange>(typeof(ExcelExtensionScope).Name));
        }

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            if (this.All && this.EntireRow && this.EntireColumn)
                metadata.AddValidationError("Only one of the All, EntireRow and EntireColumn options can be set");
            else
            {
                if (this.All && this.EntireRow)
                    metadata.AddValidationError("Only one of the All and EntireRow options can be set");
                if (this.All && this.EntireColumn)
                    metadata.AddValidationError("Only one of the All and EntireColumn options can be set");
                if (this.EntireColumn && this.EntireRow)
                    metadata.AddValidationError("Only one of the EntireRow and EntireColumn options can be set");
            }
            base.CacheMetadata(metadata);
        
        }

        protected override void Execute(CodeActivityContext context)
        {

            var property = context.DataContext.GetProperties()[ExcelExtensionScope.ExcelTag];
            var excelProperty = property.GetValue(context.DataContext) as ExcelSession;

            selectRange(excelProperty.worksheet, All, Range.Get(context), EntireRow, EntireColumn);
        }
        private void selectRange(Worksheet ws, bool All, string range, bool EntireRow, bool EntireColumn)
        {
            Range selectRange = ws.Application.ActiveCell;
            if (!string.IsNullOrEmpty(range))
            {
                //selectRange = ws.Range[range.Split(':')[0], range.Split(':')[range.Split(':').Length - 1]];
                selectRange = ws.Range[range];
            }
            else if (All)
            {
                selectRange = ws.UsedRange;
            }
            else if (EntireRow)
            {
                selectRange = (Range)ws.UsedRange.Rows[ws.Application.ActiveCell.Row - ws.UsedRange.Row + 1];
            }
            else if (EntireColumn)
            {
                selectRange = (Range)ws.UsedRange.Columns[ws.Application.ActiveCell.Column - ws.UsedRange.Column + 1];
            }
            selectRange.Select();
        }
    }
}
