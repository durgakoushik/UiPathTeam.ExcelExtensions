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
  
    
    public class ChangeFormat : CodeActivity
    {
        public enum format
        {
            Number,
            Comma,
            Currency,
            Percentage,
            ShortDate,
            LongDate,
            Time,
            Text,
            Fraction,
            Scientific
        }

        [Description("If checked, it will add two decimal values")]
        [Category("Format Cell")]
        public bool WithDecimal  { get; set; }
        
        [Category("Format Cell")]
        public format Format { get; set; }

        public ChangeFormat()
        {
            Constraints.Add(CheckParentConstraint.GetCheckParentConstraint<ChangeFormat>(typeof(ExcelExtensionScope).Name));
        }

        protected override void Execute(CodeActivityContext context)
        {
            var property = context.DataContext.GetProperties()[ExcelExtensionScope.ExcelTag];
            var excelProperty = property.GetValue(context.DataContext) as ExcelSession;
            Range rng = (Microsoft.Office.Interop.Excel.Range)excelProperty.worksheet.Application.Selection;

            if (Format.Equals(format.Number))
                rng.NumberFormat = WithDecimal ? "0.00" : "0";
            if (Format.Equals(format.Currency))
            rng.NumberFormat = WithDecimal ? "$ #,##0.00" : "$ #,##0";
            if (Format.Equals(format.Percentage))
                rng.NumberFormat = WithDecimal ? "0.00%" : "0%";
            if (Format.Equals(format.ShortDate))
                rng.NumberFormat = "m/d/yyyy";
            if (Format.Equals(format.LongDate))
                rng.NumberFormat ="[$-x-sysdate]dddd, mmmm dd, yyyy";
            if (Format.Equals(format.Time))
                rng.NumberFormat = "[$-x-systime]h:mm:ss AM/PM";
            if (Format.Equals(format.Fraction))
                rng.NumberFormat = "# ?/?";
            if (Format.Equals(format.Scientific))
                rng.NumberFormat = WithDecimal ? "0.00E+00": "0E+00";
            if (Format.Equals(format.Text))
                rng.NumberFormat = "@";
            if (Format.Equals(format.Comma))
                rng.Style = "Comma";

            if (excelProperty.save)
            {
                excelProperty.workbook.Save();
            }

        }
       
    }
}
