using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Course2
{
    public record Transformer(double Power, double[] FirstVoltages, double[] SecondVoltages, double[] ShortVoltages,
        double NoLoadCurrent, double NoLoadLoses, double[] ShortLoses);
}
