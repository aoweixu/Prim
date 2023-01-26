using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportGen.Models.CCOImporting
{
    public class ReducerMachine
    {
        public ReducerMachine(int machineId)
        {
            this.IntMachineID = machineId;
            this.MachineID = $"R-{machineId.ToString()}";
        }

        public int IntMachineID { get; set; }
        public string MachineID { get; set; }
        public readonly int TableStartRow = 12;
        public readonly int TableHeight = 86;
        public readonly int TableColOffset = 7;
    }
}
