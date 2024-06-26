﻿using Autodesk.Navisworks.Api;
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

        public static List<string> ItensEncontrados;
        public static int CountItensEncontrados;
        public static void AttItem(Dictionary<string, List<Item>> newItens)
        {

            int cont = 0;
            List<string> ItensTest = new List<string>();
            Document doc = Application.ActiveDocument;
            Model model = doc.Models[0];
            ModelItemCollection collection = new ModelItemCollection();


            if (model != null)
                collection.Add(model.RootItem);

            ModelItemEnumerableCollection collectionT = collection.Descendants;
            Form1 form = new Form1();
            form.Show();

            foreach (var item in newItens)
            {
                Search search = new Search();
                search.Selection.SelectAll();

                SearchCondition condition = SearchCondition.HasPropertyByDisplayName("Item", "Name")
                    .EqualValue(new VariantData(item.Key));

                search.SearchConditions.Add(condition);

                ModelItemCollection itemMaquete = search.FindAll(doc, false);

                if (itemMaquete == null)
                {
                    Search searchTag = new Search();
                    searchTag.Selection.SelectAll();

                    SearchCondition conditionTag = SearchCondition.HasPropertyByDisplayName("AutoCad", "Tag")
                    .EqualValue(new VariantData(item.Key));

                    searchTag.SearchConditions.Add(conditionTag);

                    ModelItemCollection itemMaqueteTag = search.FindAll(doc, false);

                    if (itemMaqueteTag != null)
                    {

                        List<Item> listaDeItens = item.Value;

                        foreach (var itens in listaDeItens)
                        {
                            string categoria = itens.categoria;
                            Dictionary<string, string> properties = itens.properties;

                            foreach(ModelItem itemMT in itemMaqueteTag)
                            {
                                foreach (var prop in properties)
                                {
                                    PropertiesInsert.PropertyInsertion(categoria, prop.Key, prop.Value, itemMT);
                                }

                                form.UpdateBarProgress();
                                ItensTest.Add(itemMT.DisplayName);
                                CountItensEncontrados++;
                            }
                          
                        }

                    }
                    else
                    {
                        form.UpdateBarProgress();
                    }

                }
                else
                {
                    List<Item> listaDeItens = item.Value;

                    foreach (var itens in listaDeItens)
                    {
                        string categoria = itens.categoria;
                        Dictionary<string, string> properties = itens.properties;

                        foreach(ModelItem itemM in itemMaquete)
                        {
                            foreach (var prop in properties)
                            {
                                PropertiesInsert.PropertyInsertion(categoria, prop.Key, prop.Value, itemM);
                            }

                            form.UpdateBarProgress();
                            ItensTest.Add(itemM.DisplayName);
                            CountItensEncontrados++;
                        }  
                    }
                }
            }
            form.killForms();
            ItensEncontrados = ItensTest;
        }
    }
}