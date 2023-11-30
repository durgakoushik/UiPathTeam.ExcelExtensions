using Microsoft.Office.Interop.Excel;
using System;
using System.Activities;
using System.Activities.Statements;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace UiPathTeam.Excel.Extensions.Activities
{
    public class CloseSession : CodeActivity
    {
        public InArgument<ExcelSession> Session { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            Workbook wb;
            Application app;
            Worksheet ws;
            ExcelSession session = Session.Get(context);
            wb = session.workbook;
            app = session.application;
            ws = session.worksheet;

            if (ws != null)
            {
                Marshal.ReleaseComObject(ws);
                ws = null;
            }
            if (wb != null)
            {
                wb.Close();
                Marshal.ReleaseComObject(wb);
                wb = null;
            }
            if (app != null)
            {
                app.Quit();
                Marshal.ReleaseComObject(app);
                app = null;
            }
        }
    }
}
