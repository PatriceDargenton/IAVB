
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IAVB
{
    class clsUtilWPF
    {
        public static void DoEvents()
        {
            System.Windows.Application.Current.Dispatcher.Invoke(
                System.Windows.Threading.DispatcherPriority.Background,
                    new Action(delegate { }));
        }
    }
}
