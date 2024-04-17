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
    public static class TransformersReader
    {
        public static List<Transformer> GetTransformers()
        {
            var path = @"sources\Transformers.xlsx";

            using var workbook = ReoGridControl.CreateMemoryWorkbook();
            workbook.Load(path, FileFormat.Excel2007);
            var sheet1 = workbook.Worksheets[0];

            List<Transformer> transformers = new List<Transformer>();

            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("en-US");

            for (int i = 11; i < 145; i += 7)
            {

                var transformer = new Transformer
                    (
                        Power: sheet1.GetCellData<double>("A" + i),
                        FirstVoltages: sheet1.GetCellData<string>("A" + (i + 1)).Split(';').Select(x => double.Parse(x.Trim())).ToArray(),
                        SecondVoltages: sheet1.GetCellData<string>("A" + (i + 2)).Split('/').Select(x => double.Parse(x.Trim())).ToArray(),
                        ShortVoltages: sheet1.GetCellData<string>("A" + (i + 3)).Split('/').Select(x => double.Parse(x.Trim())).ToArray(),
                        NoLoadCurrent: sheet1.GetCellData<double>("A" + (i + 4)),
                        NoLoadLoses: sheet1.GetCellData<double>("A" + (i + 5)),
                        ShortLoses: sheet1.GetCellData<string>("A" + (i + 6)).Split('/').Select(x => double.Parse(x.Trim())).ToArray()
                    );

                transformers.Add(transformer);
            }

            return transformers;
        }
    }
}
