using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Course2
{
    public record WireSIP(string WiresNumAndSize, double ResistancePhase25, double ResistancePhase50, 
        double ResistanceEarth25, double ResistanceEarth50, double InductancePhase, double InductanceEarth, double Current25);
}
