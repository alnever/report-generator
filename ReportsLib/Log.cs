using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportsLib
{
    public class Log
    {
        private static object sync = new object();
        public static void Write(int idRes, string Oper, Exception ex = null)
        {
            try
            {
                string message = (ex == null ? "" : ex.Message);
                string ex_type = (ex == null ? "" : ex.TargetSite.DeclaringType.ToString());
                string ex_target = (ex == null ? "" : ex.TargetSite.Name);
                string ex_string = (ex == null ? "" : ex.ToString());
                // Путь .\\Log
                // AppDomain.CurrentDomain.BaseDirectory - current applictation path
                string pathToLog = Path.Combine(Path.GetTempPath(), "ReportDLL_Log");
                if (!Directory.Exists(pathToLog))
                    Directory.CreateDirectory(pathToLog); // Создаем директорию, если нужно
                string filename = Path.Combine(pathToLog, string.Format("{0}_{1:dd.MM.yyy}.log",
                AppDomain.CurrentDomain.FriendlyName, DateTime.Now));
                string fullText = string.Format("[{0:dd.MM.yyy HH:mm:ss.fff}] [{1}.{2}()] {3}\r\n{4}\r\n",
                DateTime.Now, ex_type, ex_target, message, ex_string);
                lock (sync)
                {
                    File.AppendAllText(filename, "----------------------------------------", Encoding.GetEncoding("Windows-1251"));
                    File.AppendAllText(filename, idRes.ToString() + " " + Oper, Encoding.GetEncoding("Windows-1251"));
                    File.AppendAllText(filename, fullText, Encoding.GetEncoding("Windows-1251"));
                }
            }
            catch
            {
                // Перехватываем все и ничего не делаем
            }
        }
    }
}
