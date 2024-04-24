using Autodesk.Navisworks.Api;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Application = Autodesk.Navisworks.Api.Application;
using Properties.Objects;

namespace Properties.Methods
{
    static class FindItem
    {
        public static void AttItem()
        {
            int cont = 0;
            List<Item> newItens = ReadExcel.Itens;
            Document doc = Application.ActiveDocument;
            Model model = doc.Models[0];
            ModelItemCollection collection = new ModelItemCollection();

            foreach(Item i in newItens)
            {
                string name = i.itemName;
            }

            if (model != null)
                collection.Add(model.RootItem);

            ModelItemEnumerableCollection collectionT = collection.Descendants;

            foreach (var item in newItens)
            {
                Search search = new Search();
                search.Selection.SelectAll();

                SearchCondition condition = SearchCondition.HasPropertyByDisplayName("Item", "Name")
                    .EqualValue(new VariantData(item.itemName));

                search.SearchConditions.Add(condition);

                ModelItem itemMaquete = search.FindFirst(doc, false);

                if (itemMaquete != null)
                {
                    foreach (var kvp in item.properties)
                    {
                        string key = kvp.Key;
                        string value = kvp.Value;

                        PropertiesInsert.PropertyInsertion(item.categoria, key, value, itemMaquete);
                    }
                }
            }
        }
    }
}
