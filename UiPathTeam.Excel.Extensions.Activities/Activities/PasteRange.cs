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

    class PasteRange : CodeActivity
    {

        public PasteRange()
        {
            Constraints.Add(CheckParentConstraint.GetCheckParentConstraint<PasteRange>(typeof(ExcelExtensionScope).Name));

        }
        [Category("Input")]
        [Description("Please enter the range that you want to select or select different option from properties.")]
        public InArgument<string> Range { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var property = context.DataContext.GetProperties()[ExcelExtensionScope.ExcelTag];
            var excelProperty = property.GetValue(context.DataContext) as ExcelSession;
            string range = Range.Get(context);
        
        }
      
    }
}
