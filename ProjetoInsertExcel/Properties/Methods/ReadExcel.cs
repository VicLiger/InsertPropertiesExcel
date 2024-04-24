using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Properties.Objects;

namespace Properties.Methods
{
    static class ReadExcel
    {
        public static int validItens = 0;
        public static int propertiesCount;
        public static List<string> properties;
        public static int rowCounts = 0;
        public static object objeto;
        public static Dictionary<string, List<Item>> ListaItens;
        public static void readerExcel(string caminho)
        {
            try
            {

                string itemName = " ";
                string propValue = " ";
                string path = caminho;
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook excelFile = excelApp.Workbooks.Open(path);
                List<Item> itens = new List<Item>();
                Dictionary<string, List<Item>> listaItens = new Dictionary<string, List<Item>>();
                List<string> props = new List<string>();

                foreach (Excel.Worksheet excelWSCount in excelFile.Worksheets)
                {
                    string sheetName = excelWSCount.Name;
                    Excel.Range excelRange = excelWSCount.UsedRange;

                    int rowCount = excelRange.Rows.Count;
                    int columnCount = excelRange.Columns.Count;

                    rowCounts += rowCount;
                }

                Form3 form3 = new Form3();
                form3.Show();

                // Itera sobre todas as planilhas no arquivo Excel
                foreach (Excel.Worksheet excelWS in excelFile.Worksheets)
                {
                    string sheetName = excelWS.Name;
                    Excel.Range excelRange = excelWS.UsedRange;

                    int rowCount = excelRange.Rows.Count;
                    int columnCount = excelRange.Columns.Count;

                    for (int i = 2; i <= rowCount; i++)
                    {
                        List<Item> novaLista = new List<Item>();
                        Item item = new Item();

                        form3.UpdateProgressBar();

                        item.categoria = excelWS.Name; // Nome da categoria

                        for (int j = 1; j <= columnCount; j++)
                        {
                            object cellValue = excelRange.Cells[i, j].Value2; // Item por linha
                            objeto = cellValue;

                            if (cellValue != null)
                            {
                                if (j.Equals(1)) // Item sempre da primeira coluna ou seja o nome ou tag do item
                                {
                                    itemName = cellValue.ToString();
                                }
                                else // Obter o nome da propriedade e seu valor
                                {
                                    string propName = excelRange.Cells[1, j].Value2.ToString();

                                    if (!string.IsNullOrEmpty(propName) && cellValue != null)
                                    {
                                        propValue = cellValue.ToString();
                                        item.properties.Add(propName, propValue);

                                        if (!props.Contains(propName))
                                            props.Add(propName);
                                    }
                                }
                            }
                        }

                        if (!string.IsNullOrEmpty(itemName)) // Verifique se o itemName não está vazio
                        {
                            if (item.categoria == "DTwin_Equipamento")
                            {

                            }
                            novaLista.Add(item); // Adicione o item à nova lista

                            if (listaItens.ContainsKey(itemName))
                            {
                                listaItens[itemName].AddRange(novaLista);
                            }
                            else
                            {
                                List<Item> novaListaItens = new List<Item>();
                                novaListaItens.Add(item);
                                listaItens.Add(itemName, novaListaItens);

                            }



                            validItens++;
                        }
                    }
                    Marshal.ReleaseComObject(excelRange);
                    Marshal.ReleaseComObject(excelWS);

                }
                form3.killForms();

                ListaItens = listaItens;
                properties = props;


                int countList = itens.Count();

                excelFile.Close(true);
                Marshal.ReleaseComObject(excelFile);
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }
            catch (Exception e)
            {
                MessageBox.Show("Ocorreu um erro ao buscar o excel: " + e.Message);
            }
        }
    }
}
