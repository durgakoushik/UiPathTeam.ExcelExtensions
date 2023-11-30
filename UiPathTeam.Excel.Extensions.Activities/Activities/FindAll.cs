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
    public class FindAll : CodeActivity
    {
        [Description("Value that you want to search.")]
        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> Value { get; set; }

        [Description("All that cell names that contains the value")]
        [Category("Output")]
        [RequiredArgument]
        public OutArgument<List<string>> Cells { get; set; }
        public FindAll()
        {
            Constraints.Add(CheckParentConstraint.GetCheckParentConstraint<FindAll>(typeof(ExcelExtensionScope).Name));

        }
        protected override void Execute(CodeActivityContext context)
        {
            var property = context.DataContext.GetProperties()[ExcelExtensionScope.ExcelTag];
            var excelProperty = property.GetValue(context.DataContext) as ExcelSession;

            Workbook wb = excelProperty.workbook;
            Application app = excelProperty.application;

            Worksheet ws = excelProperty.worksheet;
            Microsoft.Office.Interop.Excel.Range currentFind = null;
            object missing = Type.Missing;
            Range r = (Range)ws.UsedRange;
            currentFind = r.Find(Value.Get(context), missing,
            Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues, Microsoft.Office.Interop.Excel.XlLookAt.xlPart,
            Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext, false,
            missing, missing);

            List<string> address = new List<string>();
            try
            {
                do
                {
                    string first = currentFind.get_Address(Type.Missing, Type.Missing, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

                    if (address.Contains(first))
                    {
                        break;
                    }
                    address.Add(first);
                    currentFind = r.FindNext(currentFind);
                } while (true);

                if (excelProperty.save)
                {
                    excelProperty.workbook.Save();
                }
            }
            catch (Exception)
            {

            }

            Cells.Set(context, address);

        }
    }
}
