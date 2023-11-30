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
    public enum verticalAllign
    {
        Bottom,
        Center,
        Top,
        Justify,
        Distributed
    }
    public enum horizontalAllign
    {
        Center,
        Left,
        Right,
        Justify,
        Distributed
    }

    public enum border
    {
        Thin,
        Thick,
        Medium,
        Hairline
    }
    [Description("Applies style to the range that is selected")]
    public class FontStyle : CodeActivity
    {   
       
        [Category("Style")]
        public bool Bold { get; set; }
        [Category("Style")]
        public bool Italic { get; set; }
        [Category("Style")]
        public bool Underline { get; set; }
        [Category("Style")]
        public InArgument<string> FontName { get; set; }
        [RequiredArgument]
        [Category("Style")]
        public InArgument<int> Size { get; set; }

        [Category("Style")]
        public horizontalAllign HorizontalAlignment { get; set; }

        [Category("Style")]
        public verticalAllign VerticalAlignment { get; set; }


        [Category("Border")]
        public bool AllBorders { get; set; }

        [Category("Border")]
        public border BorderWidth { get; set; }

        [Category("Style")]
        public System.Drawing.Color FontColor { get; set; }

        public FontStyle()
        {
            Constraints.Add(CheckParentConstraint.GetCheckParentConstraint<FontStyle>(typeof(ExcelExtensionScope).Name));
            Size = new InArgument<int>(10);
        }
        protected override void Execute(CodeActivityContext context)
        {

            var property = context.DataContext.GetProperties()[ExcelExtensionScope.ExcelTag];
            var excelProperty = property.GetValue(context.DataContext) as ExcelSession;
            Range rng = (Microsoft.Office.Interop.Excel.Range)excelProperty.worksheet.Application.Selection;

            if (Bold)
                rng.Font.Bold = true;
            if (Italic)
                rng.Font.Italic = true;
            if (Underline)
                rng.Font.Underline = true;

            #region Horizontal Allignment
            if (HorizontalAlignment.Equals(horizontalAllign.Center))
                rng.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            if (HorizontalAlignment.Equals(horizontalAllign.Distributed))
                rng.HorizontalAlignment = XlHAlign.xlHAlignDistributed;
            if (HorizontalAlignment.Equals(horizontalAllign.Justify))
                rng.HorizontalAlignment = XlHAlign.xlHAlignJustify;
            if (HorizontalAlignment.Equals(horizontalAllign.Left))
                rng.HorizontalAlignment = XlHAlign.xlHAlignLeft;
            if (HorizontalAlignment.Equals(horizontalAllign.Right))
                rng.HorizontalAlignment = XlHAlign.xlHAlignRight;
            #endregion



            #region Vertical Allignment
            if (VerticalAlignment.Equals(verticalAllign.Center))
                rng.VerticalAlignment = XlVAlign.xlVAlignCenter;
            if (VerticalAlignment.Equals(verticalAllign.Distributed))
                rng.VerticalAlignment = XlVAlign.xlVAlignDistributed;
            if (VerticalAlignment.Equals(verticalAllign.Justify))
                rng.VerticalAlignment = XlVAlign.xlVAlignJustify;
            if (VerticalAlignment.Equals(verticalAllign.Bottom))
                rng.VerticalAlignment = XlVAlign.xlVAlignBottom;
            if (VerticalAlignment.Equals(verticalAllign.Top))
                rng.VerticalAlignment = XlVAlign.xlVAlignTop;
            #endregion

          

            rng.Font.Size = Size.Get(context);
            rng.Font.Name = FontName.Get(context);
            rng.Font.Color = FontColor;

            #region Border
            if (AllBorders)
            {
                rng.Borders.LineStyle = XlLineStyle.xlContinuous;
                if (BorderWidth.Equals(border.Thin))
                    rng.Borders.Weight = XlBorderWeight.xlThin;
                if (BorderWidth.Equals(border.Thick))
                    rng.Borders.Weight = XlBorderWeight.xlThick;
                if (BorderWidth.Equals(border.Medium))
                    rng.Borders.Weight = XlBorderWeight.xlMedium;
                if (BorderWidth.Equals(border.Hairline))
                    rng.Borders.Weight = XlBorderWeight.xlHairline;
            } 
            #endregion

            if (excelProperty.save)
            {
                excelProperty.workbook.Save();
            }


        }
    }

}
