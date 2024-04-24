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
        public static List<Item> Itens;
        public static void readerExcel(string caminho)
        {
            try
            {
                string propValue = " ";
                string path = caminho;
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook excelFile = excelApp.Workbooks.Open(path);
                List<Item> itens = new List<Item>();

                // Itera sobre todas as planilhas no arquivo Excel
                foreach (Excel.Worksheet excelWS in excelFile.Worksheets)
                {
                    string sheetName = excelWS.Name;
                    Excel.Range excelRange = excelWS.UsedRange;

                    int rowCount = excelRange.Rows.Count;
                    int columnCount = excelRange.Columns.Count;

                    for (int i = 2; i <= rowCount; i++)
                    {
                        Item item = new Item();
                        item.categoria = excelWS.Name;

                        for (int j = 1; j <= columnCount; j++)
                        {
                            object cellValue = excelRange.Cells[i, j].Value2;

                            if (j.Equals(1))
                            {
                                item.itemName = cellValue.ToString();
                            }
                            else
                            {
                                string propName = excelRange.Cells[1, j].Value2.ToString();

                                if (!string.IsNullOrEmpty(propName) && cellValue != null)
                                {
                                    propValue = cellValue.ToString();
                                    item.properties.Add(propName, propValue);
                                }
                            }
                        }


                        itens.Add(item);
                    }

                    Marshal.ReleaseComObject(excelRange);
                    Marshal.ReleaseComObject(excelWS);
                }

                Itens = itens;

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
