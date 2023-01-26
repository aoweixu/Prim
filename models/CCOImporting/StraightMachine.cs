using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportGen.Models.CCOImporting
{
    public class StraightMachine
    {
        /// <summary>
        /// Provide Machine ID (number)
        /// </summary>
        /// <param name="machineId"></param>
        public StraightMachine(int machineId)
        {
            this.IntMachineID = machineId;
            this.MachineID = $"ST-{machineId.ToString()}";
        }

        public string MachineID { get; set; }
        public int IntMachineID { get; set; }
        public readonly int TableStartRow = 12;
        public readonly int TableHeight = 67;
        public readonly int TableColOffset = 7;

    }
}
