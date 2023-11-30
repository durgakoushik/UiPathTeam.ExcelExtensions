using Microsoft.Office.Interop.Excel;
using Moq;
using System.Activities;
using System.Activities.Statements;
using System.ComponentModel;
using System.Dynamic;
using System.Reflection;
using UiPathTeam.Excel.Extensions.Activities;

namespace UiPathTeam.Excel.Extensions.UnitTest
{
    public class Tests
    {

        private ExcelSession excelSession;
        private Mock<CodeActivityContext> mockContext;
        private PropertyDescriptorCollection properties;
        [SetUp]
        public void Setup()
        {
            excelSession = CreateExcelSession(@"E:\OneDrive - UiPath\Desktop\MasterFile.xlsx");

            // Create mock PropertyDescriptor
            var mockPropertyDescriptor = new Mock<PropertyDescriptor>(new object[] { "ExcelScope", new Attribute[0] });
            mockPropertyDescriptor.Setup(pd => pd.GetValue(It.IsAny<object>())).Returns(excelSession);

            // Set up the PropertyDescriptorCollection
            properties = new PropertyDescriptorCollection(new[] { mockPropertyDescriptor.Object });



            // Set up a mock CodeActivityContext
            mockContext = new Mock<CodeActivityContext>();
            mockContext.Setup(ctx => ctx.DataContext.GetProperties()).Returns(properties);


        }

        [Test]
        public void ActivateSheet_WithSheetName_SheetActivated()
        {
            ActivateSheet activateSheet = new ActivateSheet() { SheetName = "Jobs" };
            var scopeActivity = CreateScopeActivityWithChildActivities(activateSheet);

            WorkflowInvoker.Invoke(scopeActivity);
            TestContext.Out.WriteLine("Executed");
            Assert.Pass();
        }

        [Test]
        public void GetLastRow_NoInput_GetLastRowNumber()
        {
            

            GetLastRow getLastRow = new GetLastRow();

            var methodInfo = typeof(GetLastRow).GetMethod("Execute", BindingFlags.NonPublic | BindingFlags.Instance);
            if (methodInfo == null)
            {
                Assert.Fail("Method 'Execute' not found");
                return;
            }
            var output = methodInfo.Invoke(getLastRow, new object[] { mockContext.Object });




            Assert.Pass();
        }
        [Test]
        public void GetLastRow_NoInput_GetLastRowNumber2()
        {
            //var scopeActivity = new ExcelExtensionScope
            //{
            //    FilePath = @"E:\OneDrive - UiPath\Desktop\MasterFile.xlsx"
            //};


            //WorkflowInvoker.Invoke(scopeActivity);


            //GetLastRow getLastRow = new GetLastRow();
            //var sec = (Sequence)scopeActivity.Body.Handler;
            //sec.Activities.Add(getLastRow);
            //var outputs = WorkflowInvoker.Invoke(scopeActivity2);

            //TestContext.Out.WriteLine("Executed");
            //Assert.Pass();
        }
        //[Test]
        //public void GetActiveCell_NoInput_CellValue()
        //{
        //    GetActiveCell getActiveCell = new GetActiveCell();
        //    var scopeActivity = CreateScopeActivityWithChildActivities(getActiveCell);

        //    var inputs = new Dictionary<string, object>
        //    {
        //        {"CellName" }
        //    }
        //    var output = WorkflowInvoker.Invoke(scopeActivity);
        //    TestContext.Out.WriteLine("Executed");
        //    Assert.Pass();
        //}
        #region ScopeActivity
        public static ExcelExtensionScope CreateScopeActivityWithChildActivities(params Activity[] childActivites)
        {
            
            return CreateScopeActivityWithChildActivities(@"E:\OneDrive - UiPath\Desktop\MasterFile.xlsx", childActivites);
            //return CreateScopeActivityWithChildActivities(Env.Instance.TestExcelFilePath, childActivites);
        }
        public static ExcelExtensionScope CreateScopeActivityWithChildActivities(string filePath, params Activity[] childActivites)
        {
            var scopeActivity = new ExcelExtensionScope
            {
                FilePath = filePath
            };

            Sequence handlerSequence = (Sequence)scopeActivity.Body.Handler;

            foreach (var activity in childActivites)
            {
                handlerSequence.Activities.Add(activity);
            }

            return scopeActivity;
        }
        public static ExcelSession CreateExcelSession(string filePath)
        {
            Application app = null;
            Workbook wb;
            Worksheet ws;
            bool isOpned = false;
            object missing = Type.Missing;

            bool fileExist = File.Exists(Path.Combine(Environment.CurrentDirectory, filePath));

            if (fileExist)
                filePath = Path.Combine(Environment.CurrentDirectory, filePath);

            wb = null; ;

            try
            {

                app = MarshalForCore.GetActiveObject("Excel.Application") as Application;
                foreach (Workbook workbook in app.Workbooks)
                {
                    if (filePath.ToLower().Equals(workbook.FullName.ToLower()))
                    {
                        wb = workbook;
                        isOpned = true;
                        break;
                    }
                }
            }
            catch (Exception)
            {
            }


            if (wb == null)
            {
                app = new Application();
                app.Visible = true;

                bool excelExist = File.Exists(filePath);
                if (excelExist)
                {
                    // open the workbook. 
                    wb = app.Workbooks.Open(
                    filePath, true
                  , false, missing, "", missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
                }
                else
                {
                    throw new FileNotFoundException("Excel File " + filePath + " was not found");
                }
            }

            app.ActiveWindow.Activate();
            ws = (Worksheet)wb.ActiveSheet;
            ws.Activate();


            return new ExcelSession()
            {
                save = true,
                worksheet = ws,
                workbook = wb,
                application = app
            };
        }
        #endregion

    }


}