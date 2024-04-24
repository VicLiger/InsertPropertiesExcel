using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;

namespace Properties.Objects
{
    class Item
    {
        public string categoria;
        public string itemName;
        public Dictionary<string, string> properties;

        public Item()
        {
            // Inicializa o dicionário properties
            properties = new Dictionary<string, string>();
        }
    }
}
