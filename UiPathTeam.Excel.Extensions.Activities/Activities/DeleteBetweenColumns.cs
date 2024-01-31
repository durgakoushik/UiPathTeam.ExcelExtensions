using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using UiPathTeam.Excel.Extensions.Activities.Properties;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;
using System.ComponentModel;
using System.Data.Common;
using System.IO.Packaging;
using System.Windows.Automation;
using Microsoft.Office.Interop.Excel;
using System.Windows.Input;

namespace UiPathTeam.Excel.Extensions.Activities
{
    //[LocalizedDisplayName(nameof(Resources.DeleteBetweenColumns_DisplayName))]
    //[LocalizedDescription(nameof(Resources.DeleteBetweenColumns_Description))]
    //public class DeleteBetweenColumns : CodeActivity
    //{
    //    #region Properties

    //    /// <summary>
    //    /// If set, continue executing the remaining activities even if the current activity has failed.
    //    /// </summary>
        
    //    //[LocalizedCategory(nameof(Resources.Common_Category))]
    //    //[LocalizedDisplayName(nameof(Resources.ContinueOnError_DisplayName))]
    //    //[LocalizedDescription(nameof(Resources.ContinueOnError_Description))]
    //    //public override InArgument<bool> ContinueOnError { get; set; }

    //    [Description("Column Name")]
    //    [Category("Input")]
    //    [DisplayName("Start Column Name")]
    //    public InArgument<string> StartColumnName { get; set; }

    //    [Description("Column Name")]
    //    [Category("Input")]
    //    [DisplayName("End Column Name")]
    //    public InArgument<string> EndColumnName { get; set; }

    //    #endregion


    //    #region Constructors

    //    public DeleteBetweenColumns()
    //    {
    //        Constraints.Add(CheckParentConstraint.GetCheckParentConstraint<DeleteBetweenColumns>(typeof(ExcelExtensionScope).Name));
    //    }

    //    #endregion


    //    #region Protected Methods

    //    protected override void CacheMetadata(CodeActivityMetadata metadata)
    //    {

    //        base.CacheMetadata(metadata);
    //    }

    //    protected override void Execute(CodeActivityContext context)
    //    {
    //        // Inputs

    //        ///////////////////////////
    //        // Add execution logic HERE
    //        ///////////////////////////

    //        #region INIT

    //        string startColumnName = StartColumnName.Get(context);
    //        string endColumnName = EndColumnName.Get(context);
    //        Console.WriteLine("Start name from context");

    //        var property = context.DataContext.GetProperties()[ExcelExtensionScope.ExcelTag];
    //        var excelProperty = property.GetValue(context.DataContext) as ExcelSession;
    //        Console.WriteLine("excel props");

    //        // Get the column indexes for the start and end columns
    //        int startIndex = GetColumnIndex(excelProperty.worksheet, startColumnName,excelProperty);
    //        int endIndex = GetColumnIndex(excelProperty.worksheet, endColumnName,excelProperty);
    //        Console.WriteLine("start index and end index");
            
    //        Console.WriteLine("Startindex : "+startIndex.ToString());
    //        Console.WriteLine("endindex : " + endIndex.ToString());
            
    //        #endregion

    //        #region Verify column name for empty value

    //        if (string.IsNullOrEmpty(startColumnName) || String.IsNullOrWhiteSpace(startColumnName))
    //            {
    //                throw new ArgumentNullException("Value of the start column name is empty");
    //            }
    //        if (string.IsNullOrEmpty(endColumnName) || String.IsNullOrWhiteSpace(endColumnName))
    //        {
    //            throw new ArgumentNullException("Value of the end column name is empty");
    //        }
    //        Console.WriteLine("column names empty verification");

    //        #endregion

    //        #region Check if the column is available in range 

    //        if (startIndex == -1 || endIndex == -1)
    //        {
    //            throw new ArgumentException("One or both of the specified columns do not exist in the worksheet.");
    //        }
    //        Console.WriteLine("completed index verification");

    //        #endregion

    //        #region Delete the columns between the specified columns         

    //       // for (int i = endIndex - 1; i > startIndex; i--)
    //        //{
    //            // Use the DirectColumns property to directly access columns by index
    //          // ((Microsoft.Office.Interop.Excel.Range)excelProperty.worksheet.Columns[i, Type.Missing]).Delete(Type.Missing);
    //        //}

    //       // int columnsToDelete = endIndex - startIndex - 1;
    //       // (Microsoft.Office.Interop.Excel.Range) (excelProperty.worksheet.Columns[startIndex + 1, Type.Missing]).Resize[Type.Missing, columnsToDelete].Delete();

    //        for (int i = endIndex; i > startIndex; i--)
    //        {
    //            ((Microsoft.Office.Interop.Excel.Range)excelProperty.worksheet.Columns[i, Type.Missing]).Delete(Type.Missing);
    //        }

    //        Console.WriteLine("Deleted columns between " + startColumnName + " and " + endColumnName);

    //        #endregion

    //        // Outputs
    //        // return (ctx) => {};
    //    }

    //    #endregion

    //    #region function
    //    static int GetColumnIndex(Worksheet worksheet, string columnName,ExcelSession excelProperty)
    //    {
    //        Console.WriteLine("inside func");
    //        int columnCount = excelProperty.worksheet.UsedRange.Columns.Count;
    //        for (int i = 1; i <= columnCount; i++)
    //        {
    //            Microsoft.Office.Interop.Excel.Range cell = excelProperty.worksheet.Cells[1, i] as Microsoft.Office.Interop.Excel.Range;
    //            if (cell != null && cell.Value != null && cell.Value.ToString().Trim() == columnName)
    //            {
    //               // ReleaseComObject(cell);
    //                return i;
    //            }

    //            //ReleaseComObject(cell);
    //        }
    //        return -1;
    //    }
    //    #endregion

    //}

    
    
}

