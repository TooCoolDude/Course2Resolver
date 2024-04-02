using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using unvell.ReoGrid.IO;
using unvell.ReoGrid;

namespace Course2
{
    public static class WiresASReader
    {
        public static List<WireAS> GetWiresAS()
        {
            var path = @"sources\WiresAS.xlsx";

            using var workbook = ReoGridControl.CreateMemoryWorkbook();
            workbook.Load(path, FileFormat.Excel2007);
            var sheet1 = workbook.Worksheets[0];

            List<WireAS> wires = new List<WireAS>();

            for (int i = 11; i < 74; i += 7)
            {

                var wire = new WireAS
                    (
                        WireSize: sheet1.GetCellData<double>("A" + i),
                        TypeA_Mass: sheet1.GetCellData<double>("A" + (i + 1)),
                        TypeA_Current: sheet1.GetCellData<double>("A" + (i + 2)),
                        TypeA_Resistance: sheet1.GetCellData<double>("A" + (i + 3)),
                        TypeAS_Mass: sheet1.GetCellData<double>("A" + (i + 4)),
                        TypeAS_Current: sheet1.GetCellData<double>("A" + (i + 5)),
                        TypeAS_Resistance: sheet1.GetCellData<double>("A" + (i + 6)),

                        Inductance1: sheet1.GetCellData<double>("A" + (i + 72)),
                        Inductance6to10: sheet1.GetCellData<double>("A" + (i + 73)),
                        Inductance35: sheet1.GetCellData<double>("A" + (i + 74)),
                        Inductance110: sheet1.GetCellData<double>("A" + (i + 75))
                    );

                wires.Add(wire);
            }

            return wires;
        }
    }
}
