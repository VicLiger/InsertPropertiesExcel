using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.Navisworks.Api.Plugins;
using Properties.Methods;

namespace Properties
{
    [PluginAttribute("addProperties",
                       "PDMS",
                       DisplayName = "AddProperties")]
    public class ActivePlugin : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            VerifyModel v = new VerifyModel();
            bool TrueOrFalse = v.verify();

            if (TrueOrFalse)
            {
                ReadExcel.readerExcel(@"D:\C.xlsx");
                FindItem.AttItem();
            }
            return 0;
        }
    }
}
