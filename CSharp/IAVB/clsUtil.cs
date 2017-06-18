
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IAVB
{
    class Util
    {
        public static void Attendre(int iMilliSec = 200)
        {
            System.Threading.Thread.Sleep(iMilliSec);
        }
    }
}
