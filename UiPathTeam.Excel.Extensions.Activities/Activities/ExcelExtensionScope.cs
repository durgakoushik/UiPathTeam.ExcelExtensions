using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using System.Activities.Statements;
using System.ComponentModel;
using UiPathTeam.Excel.Extensions.Activities.Properties;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Runtime.InteropServices;
using UiPath.Shared.Activities.Localization;

namespace UiPathTeam.Excel.Extensions.Activities
{
    [LocalizedDisplayName(nameof(Resources.ExcelExtensionsScope_DisplayName))]
    [LocalizedDescription(nameof(Resources.ExcelExtensionsScope_Description))]
    public class ExcelExtensionScope : NativeActivity
    {
        #region Properties

        [Browsable(false)]
        public ActivityAction<ExcelSession> Body { get; set; }

        /// <summary>
        /// If set, continue executing the remaining activities even if the current activity has failed.
        /// </summary>
        //[LocalizedCategory(nameof(Resources.Common_Category))]
        //[LocalizedDisplayName(nameof(Resources.ContinueOnError_DisplayName))]
        //[LocalizedDescription(nameof(Resources.ContinueOnError_Description))]
        //public override InArgument<bool> ContinueOnError { get; set; }

        //[LocalizedDisplayName(nameof(Resources.ExcelExtensionsScope_FilePath_DisplayName))]
        //[LocalizedDescription(nameof(Resources.ExcelExtensionsScope_FilePath_Description))]
        //[LocalizedCategory(nameof(Resources.Input_Category))]
        //public InArgument<string> FilePath { get; set; }

        
        //-----------------
        [Description("Excel File Path")]
        [Category("File")]
        [RequiredArgument]
        public InArgument<string> FilePath { get; set; }

        [Description("Password of the workbook, if necessary")]
        [Category("File")]
        public InArgument<string> Password { get; set; }


        [Description("Save excel after each child excel activity exectues. If not checked, it will execute all the activities and at the end it will save the file.")]
        [Category("Options")]
        [DefaultValue(true)]
        public bool Save { get; set; }

        [DefaultValue(true)]
        [Category("Options")]
        public bool Visible { get; set; }

        [DefaultValue(true)]
        [Category("Options")]
        public bool CreateNewFile { get; set; }

        [DefaultValue(true)]
        [Category("Options")]
        public bool Activate { get; set; }


        [Category("Output")]
        public OutArgument<ExcelSession> Session { get; set; }


        [Category("Use Existing Session")]
        public InArgument<ExcelSession> ExistingSession { get; set; }

        [Category("Use Existing Session")]
        [DefaultValue(true)]
        public bool CloseSession { get; set; }

        private ExcelSession excelProperty;
        private Application app;
        private Workbook wb;
        private Worksheet ws;
        private bool isOpned = false;

        // A tag used to identify the scope in the activity context
        internal static string ExcelTag => "ExcelScope";
        //-------------------
        // Object Container: Add strongly-typed objects here and they will be available in the scope's child activities.
        private readonly IObjectContainer _objectContainer;

        #endregion


        #region Constructors

        public ExcelExtensionScope(IObjectContainer objectContainer)
        {
            _objectContainer = objectContainer;
            CloseSession = true;
            Save = true;
            Visible = true;
            CreateNewFile = true;
            Activate = true;
            Body = new ActivityAction<ExcelSession>
            {
                Argument = new DelegateInArgument<ExcelSession> (ExcelTag),
                Handler = new Sequence { DisplayName = Resources.Do }
            };
        }

        public ExcelExtensionScope() : this(new ObjectContainer())
        {

        }

        #endregion


        #region Protected Methods

        protected override void CacheMetadata(NativeActivityMetadata metadata)
        {

            base.CacheMetadata(metadata);
        }

        protected override void Execute(NativeActivityContext  context)
        {
            if (ExistingSession.Get(context) != null)
            {
                excelProperty = ExistingSession.Get(context);
            }
            else
            {
                object missing = Type.Missing;
                string filePath = FilePath.Get(context);

                bool fileExist = File.Exists(Path.Combine(Environment.CurrentDirectory, filePath));

                if (fileExist)
                    filePath = Path.Combine(Environment.CurrentDirectory, filePath);
                string password = Password.Get(context);

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
                    if (Visible)
                        app.Visible = true;

                    bool excelExist = File.Exists(filePath);
                    if (excelExist)
                    {
                        // open the workbook. 
                        wb = app.Workbooks.Open(
                        filePath, true
                      , false, missing, password, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
                    }
                    else
                    {
                        if (CreateNewFile)
                        {
                            string dir = Path.GetDirectoryName(filePath);
                            bool dirExist = Directory.Exists(dir);
                            if (!dirExist)
                            {
                                dirExist = Directory.Exists(Path.Combine(Environment.CurrentDirectory, dir));
                                if (dirExist)
                                    filePath = Path.Combine(Environment.CurrentDirectory, filePath);
                            }
                            wb = app.Workbooks.Add();
                            wb.Password = password;
                            wb.SaveAs(filePath);
                        }
                        else
                            throw new FileNotFoundException("Excel File " + filePath + " was not found");
                    }
                }
                if (Activate)
                    app.ActiveWindow.Activate();
                ws = (Worksheet)wb.ActiveSheet;
                ws.Activate();


                excelProperty = new ExcelSession()
                {
                    save = this.Save,
                    worksheet = ws,
                    workbook = wb,
                    application = app
                };
            }
            if (Body != null)
            {
                Session.Set(context, excelProperty);
                context.ScheduleAction<ExcelSession>(Body, excelProperty, OnCompleted, OnFaulted);
            }
        }

        #endregion


        #region Events

        private void OnFaulted(NativeActivityFaultContext faultContext, Exception propagatedException, ActivityInstance propagatedFrom)
        {
            Session.Set(faultContext, excelProperty);
            CloseExcelSession(faultContext);
            faultContext.CancelChildren();
            Cleanup();
        }

        private void OnCompleted(NativeActivityContext context, ActivityInstance completedInstance)
        {
            CloseExcelSession(context);
            Cleanup();
        }

        #endregion


        #region Helpers
        
        private void Cleanup()
        {
            var disposableObjects = _objectContainer.Where(o => o is IDisposable);
            foreach (var obj in disposableObjects)
            {
                if (obj is IDisposable dispObject)
                    dispObject.Dispose();
            }
            _objectContainer.Clear();
        }
        private void CloseExcelSession(NativeActivityContext context)
        {
            if (excelProperty == null)
            {
                return;
            }
            if (ExistingSession.Get(context) == null && Session.Expression == null && CloseSession)
            {

                if (!isOpned)
                {
                    if (wb != null)
                        wb.Close();

                    if (app != null)
                        app.Quit();

                }
                if (ws != null)
                {
                    Marshal.ReleaseComObject(ws);
                    ws = null;
                }
                if (wb != null)
                {
                    Marshal.ReleaseComObject(wb);
                    wb = null;
                }
                if (app != null)
                {
                    Marshal.ReleaseComObject(app);
                    app = null;
                }
            }
        }
        #endregion
    }
    public class ExcelSession
    {
        public Workbook workbook { get; set; }
        public Application application { get; set; }
        public Worksheet worksheet { get; set; }
        public bool save { get; set; }
    }
}

