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

    public class InsertIDelete : CodeActivity
    {
        public enum iorD
        {
            Insert,
            Delete
        }
        public enum rorC
        {
            Row,
            Column
        }
        [Description("Insert adds a column or row prior to the selected cell.")]
        [Category("Input")]
        public iorD InsertOrDelete { get; set; }
        [Category("Input")]
        public rorC RowOrColumn { get; set; }

        public InsertIDelete()
        {
            Constraints.Add(CheckParentConstraint.GetCheckParentConstraint<InsertIDelete>(typeof(ExcelExtensionScope).Name));

        }

        protected override void Execute(CodeActivityContext context)
        {
            var property = context.DataContext.GetProperties()[ExcelExtensionScope.ExcelTag];
            var excelProperty = property.GetValue(context.DataContext) as ExcelSession;
            Microsoft.Office.Interop.Excel.Range rng = excelProperty.application.ActiveCell;
            if (InsertOrDelete.Equals(iorD.Delete))
            {
                if (RowOrColumn.Equals(rorC.Row))
                    rng.EntireRow.Delete();
                else
                    rng.EntireColumn.Delete();
            }
            else
            {
                if (RowOrColumn.Equals(rorC.Row))
                    rng.EntireRow.Insert();
                else
                    rng.EntireColumn.Insert();
            }
            if (excelProperty.save)
            {
                excelProperty.workbook.Save();
            }
        }
    }
}
