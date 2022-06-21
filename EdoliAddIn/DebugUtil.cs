using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EdoliAddIn
{
    class DebugUtil
    {
        public static System.IO.StreamWriter logFile = new System.IO.StreamWriter("C:/Users/edoli/workspace/EdoliAddin/log.txt");
        //public static System.IO.StreamWriter logFile = new System.IO.StreamWriter("log.txt");


        public static void WriteLine(object obj)
        {
            logFile.WriteLine(obj);
            logFile.Flush();
        }
    }
}
