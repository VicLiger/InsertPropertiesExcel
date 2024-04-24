using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using Properties.Objects;

namespace Properties.Methods
{
    class JsonConvert
    {
        public JsonConvert(Dictionary<string, List<Item>> itens, string filePath)
        {


            JObject jsonItems = new JObject();

            // Itera sobre cada par chave-valor no dicionário de itens
            foreach (var parChaveValor in itens)
            {
                string chave = parChaveValor.Key; // Obtém a chave
                List<Item> listaDeItens = parChaveValor.Value; // Obtém a lista de itens associada a essa chave

                // Cria um array JSON para representar a lista de itens
                JArray jsonItemList = new JArray();

                // Itera sobre cada Item na lista
                foreach (var item in listaDeItens)
                {
                    // Cria um objeto JSON para representar o item
                    JObject jsonItem = new JObject();
                    jsonItem.Add("categoria", item.categoria); // Adiciona a categoria do item ao objeto JSON

                    // Cria um objeto JSON para representar as propriedades do item
                    JObject jsonProperties = new JObject();
                    foreach (var prop in item.properties)
                    {
                        jsonProperties.Add(prop.Key, prop.Value); // Adiciona cada propriedade ao objeto JSON de propriedades
                    }
                    jsonItem.Add("properties", jsonProperties); // Adiciona o objeto JSON de propriedades ao objeto JSON do item

                    jsonItemList.Add(jsonItem); // Adiciona o objeto JSON do item ao array JSON
                }

                jsonItems.Add(chave, jsonItemList); // Adiciona o array JSON de itens ao objeto JSON principal usando a chave
            }

            // Escreve o JSON em um arquivo
            if (filePath != null && filePath != "")
                File.WriteAllText(filePath, jsonItems.ToString());
        }
        // Cria um objeto JObject para representar os itens
    }
}


