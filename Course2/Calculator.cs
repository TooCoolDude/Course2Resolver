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

            var lightsStep = v.LightsStep;
            d["{LightsStep}"] = lightsStep.ToString();

            //select consumer
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

            var Nlights = lineLength / lightsStep;
            d["{Nlights}"] = Nlights.ToString();

            var Plights = 0.125 * Nlights;
            d["{Psum}"] = Plights.ToString();

            var Slights = Plights / 0.92;
            d["{Slights}"] = Slights.ToString();

            var Stpmax = Sdaymax + Slights;
            d["{Stpmax}"] = Sdaymax.ToString();

            var St = Stpmax * 1.3;
            d["{St}"] = St.ToString();

            //select transformer
            var tr = new Transformer(1, new[] { 1.0 } ,new[] { 1.0 }, new[] { 1.0 }, 1, 1, new[]{ 1.0});

            var trSnom = tr.Power;
            d["{trSnom}"] = trSnom.ToString();

            var trUkz = tr.ShortVoltages[1]; // 0 1
            d["{trUkz}"] = trUkz.ToString();

            var trIxx = tr.NoLoadCurrent;
            d["{trIxx}"] = trIxx.ToString();

            var trLoseXX = tr.NoLoadLoses;
            d["{trLoseXX}"] = trLoseXX.ToString();

            var trLoseKZ = tr.ShortLoses[1]; // 0 1
            d["{trLoseKZ}"] = trLoseKZ.ToString();

            var Imax04 = Stpmax / (0.38 * Math.Sqrt(3));
            d["{Imaxtp}"] = Imax04.ToString();

            //select wiresSIP 0.4
            var sip = new WireSIP("",1,1,1,1,1,1,1);

            var sipr0 = sip.ResistancePhase25;
            d["{sipr0}"] = sipr0.ToString();

            var sipx0 = sip.InductancePhase;
            d["{sipx0}"] = sipx0.ToString();

            var Psum = PmaxDay + Plights;
            d["{Psum}"] = Psum.ToString();

            var Ma04 = Psum * lineLength / 1000;
            d["{Ma04}"] = Ma04.ToString();

            var Mp04 = Qdaymax * lineLength / 1000;
            d["{Mp04}"] = Mp04.ToString();

            var Uloses04 = (sipr0 * Ma04 / 0.38) + (sipx0 * Mp04 / 0.38);
            d["{Uloses04}"] = Uloses04.ToString();



            var Imax10 = Stpmax / (10 * Math.Sqrt(3));
            d["{Imax10}"] = Imax10.ToString();

            //select wireAS 10
            var asw = new WireAS(1,1,1,1,1,1,1,1,1,1,1);

            var aswr0 = asw.TypeAS_Resistance;
            d["{awsr0}"] = aswr0.ToString();

            var aswx0 = asw.Inductance6to10;
            d["{aswx0}"] = aswx0.ToString();

            var Ma10 = Psum * 12 / 1000;
            d["{Ma10}"] = Ma04.ToString();

            var Mp10 = Qdaymax * 12 / 1000;
            d["{Mp10}"] = Mp04.ToString();

            var Uloses10 = (aswr0 * Ma10 / 10) +(aswx0 * Mp10 / 10);
            d["{Uloses10}"] = Uloses10.ToString();



            var R10 = aswr0 * 12;
            d["{R10}"] = 



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
