using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Course2
{
    public static class Calculator
    {
        public static async Task<Dictionary<string, string>> GetValuesToReplace(Variant v)
        {
            var d = new Dictionary<string, string>();
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("ru-RU");

            var varN = v.Num.ToString();
            d["{varN}"] = varN;

            var objName = v.ObjName;
            d["{objName}"] = objName;

            var PmaxDay = v.PmaxDay;
            d["{PmaxDay}"] = PmaxDay.ToString();

            var PmaxEvening = v.PmaxEvening;
            d["{PmaxEvening}"] = PmaxEvening.ToString();

            var lineLength = v.LineLength;
            d["{LineLength}"] = lineLength.ToString();

            var lightsSrep = v.LightsStep;
            d["{LightsStep}"] = lightsSrep.ToString();

            //
            var consumerSelector = new ConsumerSelectorForm();
            consumerSelector.Show();
            var con = consumerSelector.SelectedConsumerObject;

            var dayCosFi = con.DayCosFi;
            d["{dayCosFi}"] = dayCosFi.ToString();

            var dayTgFi = con.DayTgFi;
            d["{dayTgFi}"] = con.DayTgFi.ToString();

            var eveningCosFi = con.EveningCosFi;
            d["{eveningCosFi}"] = eveningCosFi.ToString();

            var eveningTgFi = con.EveningTgFi;
            d["{eveningTgFi}"] = eveningTgFi.ToString();

            var conName = con.ObjName;
            d["{conName}"] = conName;

            var Sdaymax = PmaxDay / dayCosFi;
            d["{Sdaymax}"] = Sdaymax.ToString();

            var Seveningmax = PmaxEvening / eveningCosFi;
            d["{Seveningmax}"] = Seveningmax.ToString();

            var Qdaymax = Math.Sqrt((Sdaymax * Sdaymax) - (PmaxDay * PmaxDay));
            d["{Qdaymax}"] = Qdaymax.ToString();

            var Qeveningmax = Math.Sqrt((Seveningmax * Seveningmax) - (PmaxEvening * PmaxEvening));
            d["{Qeveningmax}"] = Qeveningmax.ToString();

            //TODO: formulas and calculations...

            foreach (var key in d.Keys)
            {
                d[key] = ShortenStringNum(d[key]);
            }

            return d;
        }

        private static string ShortenStringNum(string s)
        {
            int pointPosition = 0;
            for (int i = 0; i < s.Length; i++)
            {
                if (s[i] == ',')
                {
                    pointPosition = i;
                    for (int j = 4; j > 0; j--)
                    {
                        if (pointPosition + j <= s.Length)
                        {
                            return s.Substring(0, pointPosition + j);
                        }
                    }
                }
            }
            return s;
        }
    }
}
