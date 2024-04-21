using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using unvell.ReoGrid.IO;
using unvell.ReoGrid;
using System.Globalization;

namespace Course2
{
    public static class WiresSIPReader
    {
        public static List<WireSIP> GetWiresSIP()
        {
            var path = @"sources\WiresSIP.xlsx";

            using var workbook = ReoGridControl.CreateMemoryWorkbook();
            workbook.Load(path, FileFormat.Excel2007);
            var sheet1 = workbook.Worksheets[0];

            List<WireSIP> wires = new List<WireSIP>();

            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("en-US");
            for (int i = 15; i < 112; i+=8)
            {

                var wire = new WireSIP
                    (
                        WiresNumAndSize: sheet1.GetCellData<string>("A" + i),
                        ResistancePhase25: sheet1.GetCellData<double>("A" + (i + 1)),
                        ResistancePhase50: sheet1.GetCellData<double>("A" + (i + 2)),
                        ResistanceEarth25: sheet1.GetCellData<double>("A" + (i + 3)),
                        ResistanceEarth50: sheet1.GetCellData<double>("A" + (i + 4)),
                        InductancePhase: sheet1.GetCellData<double>("A" + (i + 5)),
                        InductanceEarth: sheet1.GetCellData<double>("A" + (i + 6)),
                        Current25: sheet1.GetCellData<double>("A" + (i + 7))
                    );

                wires.Add(wire);
            }

            return wires;
        }
    }
}
