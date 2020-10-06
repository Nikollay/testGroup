using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.ComponentModel;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace testGroup
{
    class Program
    {
        static void Main(string[] args)
        {
            string filename;
            filename = "D:\\PDM\\Проект\\САВиП\\КШУ\\ПАКБ.469256.048 Камера широкоугольная (КШУ).xml";
            if (String.IsNullOrWhiteSpace(filename)) { return; }
            //Excel.Application xlApp = new Excel.Application();
            //xlApp.Visible = false;
            //xlApp.DisplayAlerts = false;

            XDocument doc = XDocument.Load(filename);
            //Excel.Workbook wb = GetfromXDocument(doc, xlApp);
            GetfromXDocument(doc);
            //xlApp.DisplayAlerts = true;
            //filename = filename.Substring(0, filename.Length - 4);
            //wb.SaveAs(filename + "SP" + ".xlsx");
            //wb.Close();
            //xlApp = null;
        }
        public static void GetfromXDocument(XDocument doc)
        {
            //"Документация" "Сборочные единицы" "Стандартные изделия" "Прочие изделия" "Материалы" "Комплекты" 

            IEnumerable<XElement> elements1, elements2;
            //Excel.Worksheet wh, wh1, wh2;
            //Excel.Range wc;
            //Excel.Workbook wb = xlApp.Workbooks.Add("D:\\PDM\\EPDM_LIBRARY\\EPDM_Specification\\sp.xls");
            XElement tmpXEl;
            string designation;
            ////Заполняем шапку
            //wh = (Excel.Worksheet)wb.Worksheets[1];
            elements1 = doc.Root.Element("transaction").Element("project").Element("configurations").Element("configuration").Element("graphs").Elements();
            //tmpXEl = elements1.First(item => item.Attribute("name").Value.Equals("Проект"));
            //wh.Cells[1, 1] = tmpXEl.Attribute("value").Value;
            //tmpXEl = elements1.First(item => item.Attribute("name").Value.Equals("Перв.Примен."));
            //wh.Cells[1, 3] = tmpXEl.Attribute("value").Value;
            //wh.Cells[3, 14] = "Документация";
            //wc = (Excel.Range)wh.Cells[3, 14];
            //wc.Font.Underline = true;
            //wc.Font.Underline = true;
            //wc.Font.Bold = true;
            //wc.HorizontalAlignment = -4108; // xlCenter
            //wc.VerticalAlignment = -4108; // xlCenter
            //wh.Cells[5, 4] = "A3";
            //tmpXEl = elements1.First(item => item.Attribute("name").Value.Equals("Обозначение"));
            //designation = tmpXEl.Attribute("value").Value;
            //wh.Cells[5, 9] = tmpXEl.Attribute("value").Value + "СБ";
            //wh.Cells[5, 14] = "Сборочный чертеж";
            //wh.Cells[32, 12] = tmpXEl.Attribute("value").Value;
            //tmpXEl = elements1.First(item => item.Attribute("name").Value.Equals("Наименование"));
            //wh.Cells[35, 12] = tmpXEl.Attribute("value").Value;
            //tmpXEl = elements1.First(item => item.Attribute("name").Value.Equals("п_Разраб"));
            //wh.Cells[35, 8] = tmpXEl.Attribute("value").Value;
            //tmpXEl = elements1.First(item => item.Attribute("name").Value.Equals("п_Пров_P"));
            //wh.Cells[36, 8] = tmpXEl.Attribute("value").Value;
            //tmpXEl = elements1.First(item => item.Attribute("name").Value.Equals("п_Н_контр"));
            //wh.Cells[38, 8] = tmpXEl.Attribute("value").Value;
            //tmpXEl = elements1.First(item => item.Attribute("name").Value.Equals("п_Утв"));
            //wh.Cells[39, 8] = tmpXEl.Attribute("value").Value;

            //Заполняем словарь
            elements1 = doc.Root.Element("transaction").Element("project").Element("configurations").Element("configuration").Element("components").Elements();
            Record component;
            Dictionary<string, Record> dict;
            List<Record> sList;
            SortedDictionary<string, Record> dDocumentation, dAssembly, dParts, dStandard, dOther, dMaterials, dKits, dNone, dictS;
            dDocumentation = new SortedDictionary<string, Record>();
            dAssembly = new SortedDictionary<string, Record>();
            dParts = new SortedDictionary<string, Record>();
            dStandard = new SortedDictionary<string, Record>();
            dOther = new SortedDictionary<string, Record>();
            dMaterials = new SortedDictionary<string, Record>();
            dKits = new SortedDictionary<string, Record>();
            dNone = new SortedDictionary<string, Record>();
            dictS = new SortedDictionary<string, Record>();

            sList =new List<Record>();
            dict = new Dictionary<string, Record>();
            dictS = new SortedDictionary<string, Record>();
            string key;
            foreach (XElement e1 in elements1)
            {
                component = new Record();
                component.quantity = 1;
                elements2 = e1.Element("properties").Elements();
                foreach (XElement e2 in elements2)
                {
                    switch (e2.Attribute("name").Value)
                    {
                        case "Формат":
                            component.format = e2.Attribute("value").Value;
                            break;
                        case "Обозначение":
                            component.designation = e2.Attribute("value").Value;
                            break;
                        case "Наименование":
                            component.title = e2.Attribute("value").Value;
                            break;
                        case "Примечание":
                            component.note = e2.Attribute("value").Value;
                            break;
                        case "Раздел СП":
                            component.chapter = e2.Attribute("value").Value;
                            break;
                    }
                }
                key = component.designation + (char)32 + component.title;
                if (!dictS.ContainsKey(key)) { dictS.Add(key, component); }
                else dictS[key].quantity++;
            }
            
            //Заполнили словарь *******
            //Сортировка
            //string path = "d:\\Домашняя работа\\test.txt";
            //File.WriteAllLines(path, dictS.Select(kvp => string.Format("{0};{1}", kvp.Key, kvp.Value.quantity)));
            //dict = (Dictionary<string, Record>)dictS.GroupBy(g => g.Value.chapter);
            //path = "d:\\Домашняя работа\\test_s.txt";
            //File.WriteAllLines(path, dict.Select(kvp => string.Format("{0};{1}", kvp.Key, kvp.Value.quantity)));
            //MessageBox.Show(dict.Count.ToString());
            string partition = "Документация";
            int j = 6;

            var dO = dictS.GroupBy(g => g.Value.chapter).OrderBy(n => n.Key, new CustomComparer()).ToDictionary(group => group.Key, group => group.ToDictionary(pair => pair.Key, pair => pair.Value));
            //Console.WriteLine(dO.GetType());
            //var d1 = dict.GroupBy(g => g.Value.chapter).ToDictionary(group => group.Key, group => group.ToDictionary(pair => pair.Key, pair => pair.Value));
            //Console.WriteLine(d1.GetType());
            //SortedDictionary<string, object> ds = new SortedDictionary<string, object>(new CustomComparer());
            //var d1 = dO.OrderBy(g => g.K, new CustomComparer()).ToList();

            //Console.WriteLine(dO.GetType());
            //Console.WriteLine(dO.First().Key);
            //Console.WriteLine(dO.Last().Key);

            foreach (var d in dO.Values)
            ////for (int i = 0; i < dO.Count(); i++)
            {
                foreach (var v in d.Values) { sList.Add(v); }

            }
            Console.WriteLine(dictS.Count);
            Console.WriteLine(sList.Count);
            foreach (Record r in sList) { Console.WriteLine("{0} {1}",r.designation, r.title); }
            //wb.Close();
            //xlApp.Quit();
            Console.ReadKey();
            //Заполняем листы
            //foreach (KeyValuePair<string, Record> d in dict)
            //{
            //    if ((j % 4) == 0) { j++; }
            //    if (!d.Value.chapter.Equals(partition))
            //    {
            //        wc = (Excel.Range)wh.Cells[j + 2, 14];
            //        wh.Cells[j + 2, 14] = d.Value.chapter;
            //        wc.Font.Underline = true;
            //        wc.Font.Bold = true;
            //        wc.HorizontalAlignment = -4108; //xlCenter
            //        wc.VerticalAlignment = -4108; //xlCenter
            //        j += 5;
            //        partition = d.Value.chapter;
            //    }

            //    if (j > 26 & wh.Name.Equals(1))
            //    {
            //        wh1 = (Excel.Worksheet)wb.Sheets.get_Item(wb.Worksheets.Count - 1);
            //        wh2 = (Excel.Worksheet)wb.Sheets.get_Item(wb.Worksheets.Count - 2);
            //        wh1.Copy(wh2);
            //        wh = (Excel.Worksheet)wb.Sheets.get_Item(wb.Worksheets.Count - 2);
            //        j = 4;
            //    }

            //    if (j > 33)
            //    {
            //        wh1 = (Excel.Worksheet)wb.Sheets.get_Item(wb.Worksheets.Count - 1);
            //        wh2 = (Excel.Worksheet)wb.Sheets.get_Item(wb.Worksheets.Count - 2);
            //        wh1.Copy(wh2);
            //        wh = (Excel.Worksheet)wb.Sheets.get_Item(wb.Worksheets.Count - 2);
            //        j = 4;
            //    }

            //    wh.Cells[j, 4] = d.Value.format;
            //    wh.Cells[j, 9] = d.Value.designation;
            //    wh.Cells[j, 20] = d.Value.quantity;
            //    wh.Cells[j, 21] = d.Value.note;
            //    wh.Cells[j, 14] = d.Value.title;
            //    //if (d.Value.title.Length < 32) { wh.Cells[j, 14] = d.Value.title; }

                //if (d.Value.title.Length > 32)
                //{
                //    wh.Cells[j, 14] = d.Value.title.Substring(0, 31);
                //    wh.Cells[j + 1, 14] = d.Value.title.Substring(31);
                //    j += 1;
                //}

            }
            //Заполнили
            //wh1 = (Excel.Worksheet)wb.Sheets.get_Item(wb.Worksheets.Count - 1);
            //wh1.Delete();//Удаляем лист шаблон

            //if (wb.Worksheets.Count == 2)
            //{
            //    wh = (Excel.Worksheet)wb.Sheets.get_Item(1);
            //    wh.Cells[36, 19] = "";
            //}
            //if (wb.Worksheets.Count < 4) { wh1 = (Excel.Worksheet)wb.Sheets.get_Item("ЛРИ"); wh1.Delete(); } //Удаляем лист ЛРИ
            //wh = (Excel.Worksheet)wb.Sheets.get_Item(1);
            //wh.Cells[36, 22] = wb.Worksheets.Count;

            //for (int i = 2; i < wb.Worksheets.Count; i++)
            //{
            //    wh = (Excel.Worksheet)wb.Sheets.get_Item(i);
            //    wh.Cells[35, 12] = designation;
            //    if (!wh.Name.Equals("ЛРИ"))
            //    {
            //        wh.Name = i.ToString();
            //        wh.Cells[37, 22] = i;
            //    }
            //    if (wh.Name.Equals("ЛРИ")) { wh.Cells[37, 19] = wb.Worksheets.Count; }
            //}

            //return wb;
        }
        public class CustomComparer : IComparer<string>
        {
            public int Compare(string x, string y)

            {
                if (wt(x) < wt(y))
                    return -1;
                if (wt(x) > wt(y))
                    return 1;
                else return 0;

                //"Документация" "Сборочные единицы" "Стандартные изделия" "Прочие изделия" "Материалы" "Комплекты" 
                // do your own comparison however you like; return a negative value
                // to indicate that x < y, a positive value to indicate that x > y,
                // or 0 to indicate that they are equal.
            }
            private int wt(string arg)

            {
                switch (arg)
                {
                    case "Документация":
                        return 1;
                    case "Сборочные единицы":
                        return 2;
                    case "Детали":
                        return 3;
                    case "Стандартные изделия":
                        return 4;                 
                    case "Прочие изделия":
                        return 5;                
                    case "Материалы":
                        return 6;                
                    case "Комплекты":
                        return 7;                     
                    default:
                        return 8;                
                }
            }
        }
    //}
}
