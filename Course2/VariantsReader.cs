using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using unvell.ReoGrid.IO;
using unvell.ReoGrid;
using System.Security.AccessControl;
using System.Globalization;

namespace Course2
{
    public static class VariantsReader
    {
        public static List<Variant> GetVariants()
        {
            var path = @"sources\Variants.xlsx";

            using var workbook = ReoGridControl.CreateMemoryWorkbook();
            workbook.Load(path, FileFormat.Excel2007);
            var sheet1 = workbook.Worksheets[0];

            List<Variant> variants = new List<Variant>();

            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("en-US");
            for (int i = 2; i < 32; i++) 
            {

                var variant = new Variant
                    (
                        Num: sheet1.GetCellData<int>("A" + i),
                        ObjName: sheet1.GetCellData<string>("B" + i),
                        PmaxDay: sheet1.GetCellData<double>("C" + i),
                        PmaxEvening: sheet1.GetCellData<double>("D" + i),
                        LineLength: sheet1.GetCellData<double>("E" + i),
                        LightsStep: sheet1.GetCellData<double>("F" + i)
                    );

                variants.Add(variant);
            }

            return variants;
        }
    }
}
