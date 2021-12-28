using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CNC_Run_Times_and_Material_Counts
{
    class RunTime
    {
        public RunTime(string fileName, int seconds)
        {
            this.FileName = fileName;
            this.Seconds = seconds;
        }

        public string FileName { get; set; }

        public int Seconds { get; set; }
    }
}
