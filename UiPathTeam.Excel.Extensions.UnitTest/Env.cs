using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UiPathTeam.Excel.Extensions.UnitTest
{
    public class Env
    {
        private static Env instance;
        public static Env Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new Env();
                }
                return instance;
            }
        }

        private Env() => DotNetEnv.Env.Load(AppDomain.CurrentDomain.BaseDirectory + "\\..\\..\\.env");

        public string TestExcelFilePath => GetValue("ExcelFilePath");


        public string GetValue(string key) => DotNetEnv.Env.GetString(key);
    }
}
