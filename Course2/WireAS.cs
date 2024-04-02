using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Course2
{
    public record WireAS(double WireSize, double TypeA_Mass, double TypeA_Current, double TypeA_Resistance, 
        double TypeAS_Mass, double TypeAS_Current, double TypeAS_Resistance,
        double Inductance1, double Inductance6to10, double Inductance35, double Inductance110);
}
