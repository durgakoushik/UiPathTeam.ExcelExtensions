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

    public class SetZoom : CodeActivity
    {
        [RequiredArgument]
        public InArgument<int> ZoomLevel { get; set; }
        public SetZoom()
        {
            Constraints.Add(CheckParentConstraint.GetCheckParentConstraint<SetZoom>(typeof(ExcelExtensionScope).Name));
        }

        protected override void Execute(CodeActivityContext context)
        {
            var property = context.DataContext.GetProperties()[ExcelExtensionScope.ExcelTag];
            var excelProperty = property.GetValue(context.DataContext) as ExcelSession;

            excelProperty.application.ActiveWindow.Zoom = ZoomLevel.Get(context);


            if (excelProperty.save)
            {
                excelProperty.workbook.Save();
            }
        }
    }
}
