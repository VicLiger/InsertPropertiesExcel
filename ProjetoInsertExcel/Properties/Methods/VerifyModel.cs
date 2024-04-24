using Autodesk.Navisworks.Api;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = Autodesk.Navisworks.Api.Application;

namespace Properties.Methods
{
    class VerifyModel
    {
        
        public bool verify()
        {
            Document doc = Application.ActiveDocument;

            if (doc.Models.Count == 0)
                 return false;
            
            return true;
        }

    }
}
