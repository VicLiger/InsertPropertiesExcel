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
                Form2 form = new Form2(); // Forms do caminho do excel
                form.ShowDialog();


                Form4 form4 = new Form4(); // Forms da pasta pra onde vai o Json
                form4.ShowDialog();


                ReadExcel.readerExcel(form.fileName);

                
                string JsonPath = $@"{form4.fileName}";
                JsonConvert json = new JsonConvert(ReadExcel.ListaItens, JsonPath);

                FindItem.AttItem(ReadExcel.ListaItens);
            }
            return 0;
        }
    }
}
