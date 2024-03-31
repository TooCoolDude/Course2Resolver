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
    public static class ConsumerObjectsReader
    {
        public static List<ConsumerObject> GetConsumerObjects()
        {
            var path = @"sources\ElectricalConsumers.xlsx";

            using var workbook = ReoGridControl.CreateMemoryWorkbook();
            workbook.Load(path, FileFormat.Excel2007);
            var sheet1 = workbook.Worksheets[0];

            List<ConsumerObject> consumers = new List<ConsumerObject>();

            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("en-us");

            for (int i = 3; i < 15; i++)
            {

                var consumer = new ConsumerObject
                    (
                        ObjName: sheet1.GetCellData<string>("A" + i),
                        DayCosFi: sheet1.GetCellData<double>("B" + i),
                        DayTgFi: sheet1.GetCellData<double>("C" + i),
                        EveningCosFi: sheet1.GetCellData<double>("D" + i),
                        EveningTgFi: sheet1.GetCellData<double>("E" + i)
                    );

                consumers.Add(consumer);
            }

            return consumers;
        }
    }
}
