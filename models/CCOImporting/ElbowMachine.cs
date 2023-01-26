using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportGen.Models.CCOImporting
{
    public class ElbowMachine
    {
        public ElbowMachine(int machineId)
        {
            this.IntMachineID = machineId;
            this.MachineID = $"EL-{machineId.ToString()}";
        }

        public int IntMachineID { get; set; }
        public string MachineID { get; set; }
        public readonly int TableStartRow = 11;
        public readonly int TableHeight = 34;
        public readonly int TableColOffset = 7;
    }
}
