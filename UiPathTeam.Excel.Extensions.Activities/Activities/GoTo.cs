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


    public class GoTo : CodeActivity
    {
        public enum direction
        {
            None,
            Right,
            Left,
            Up,
            Down,
            FirstRow,
            FirstColumn,
            LastRow,
            LastColumn
        }
        [Description("Activate the cell.")]
        [Category("Input")]
        public InArgument<string> CellName { get; set; }
        [Description("If both CellName and Direction is passed then it will first activate the cellName and then move to the direction which is passed. If only direction is passed then it will perform action on to the active cell")]
        [Category("Input")]
        public direction Direction { get; set; }
        public GoTo()
        {
            Constraints.Add(CheckParentConstraint.GetCheckParentConstraint<GoTo>(typeof(ExcelExtensionScope).Name));

        }
        protected override void Execute(CodeActivityContext context)
        {
            var property = context.DataContext.GetProperties()[ExcelExtensionScope.ExcelTag];
            var excelProperty = property.GetValue(context.DataContext) as ExcelSession;
            Range rng;

            Range visibleCells = excelProperty.worksheet.UsedRange.SpecialCells(
                                    XlCellType.xlCellTypeVisible,
                               Type.Missing);
            if (!string.IsNullOrEmpty(CellName.Get(context)))
            {
                rng = excelProperty.worksheet.get_Range(CellName.Get(context), CellName.Get(context));
                rng.Activate();
            }
            if (this.Direction == direction.Right)
            {
                rng = excelProperty.application.ActiveCell.Offset[0, 1];

            }
            else if (this.Direction == direction.Left)
            {
                rng = excelProperty.application.ActiveCell.Offset[0, -1];

            }
            else if (this.Direction == direction.Up)
            {
                rng = excelProperty.application.ActiveCell.Offset[-1, 0];
            }
            else if (this.Direction == direction.Down)
            {
                rng = excelProperty.application.ActiveCell.Offset[1, 0];

            }
            else if (this.Direction == direction.FirstRow)
            {
                rng = excelProperty.worksheet.UsedRange;
                Microsoft.Office.Interop.Excel.Range activeRange = excelProperty.application.ActiveCell;
                string address = activeRange.Address;
                int firstRow = rng.Row;

                string FirstRow = address.Split('$')[1] + firstRow;
                rng = excelProperty.worksheet.get_Range(FirstRow, FirstRow);
            }
            else if (this.Direction == direction.FirstColumn)
            {
                rng = excelProperty.worksheet.UsedRange;
                var addres = rng.Address;
                int firstCol = rng.Column;


                int dividend = firstCol;
                string columnName = String.Empty;
                int modulo;
                while (dividend > 0)
                {
                    modulo = (dividend - 1) % 26;
                    columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                    dividend = (int)((dividend - modulo) / 26);
                }

                Microsoft.Office.Interop.Excel.Range activeRange = excelProperty.application.ActiveCell;
                string address = activeRange.Address;

                string firstColumn = columnName + address.Split('$')[2];
                rng = excelProperty.worksheet.get_Range(firstColumn, firstColumn);
            }
            else if (this.Direction == direction.LastRow)
            {
                Microsoft.Office.Interop.Excel.Range last = excelProperty.worksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
                int rowNumber = last.Row;
                Microsoft.Office.Interop.Excel.Range activeRange = excelProperty.application.ActiveCell;
                string address = activeRange.Address;

                string LastRow = address.Split('$')[1] + rowNumber;
                rng = excelProperty.worksheet.get_Range(LastRow, LastRow);
            }
            else if (this.Direction == direction.LastColumn)
            {
                Microsoft.Office.Interop.Excel.Range last = excelProperty.worksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
                int columnNumber = last.Column;

                int dividend = columnNumber;
                string columnName = String.Empty;
                int modulo;
                while (dividend > 0)
                {
                    modulo = (dividend - 1) % 26;
                    columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                    dividend = (int)((dividend - modulo) / 26);
                }

                Microsoft.Office.Interop.Excel.Range activeRange = excelProperty.application.ActiveCell;
                string address = activeRange.Address;

                string LastColumn = columnName + address.Split('$')[2];
                rng = excelProperty.worksheet.get_Range(LastColumn, LastColumn);

            }
            else
            {
                rng = excelProperty.application.ActiveCell;
            }
            rng.Select();
            rng.Activate();


            if (excelProperty.save)
            {
                excelProperty.workbook.Save();
            }
        }
    }
}
