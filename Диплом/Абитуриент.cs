using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;

namespace Диплом_С_шарп
{

    class Абитуриент
    {
        public string ФИО;
        public int[] Предметы = new int[18];
        public bool вОчереди;
        public List<int> Предпочтения = new List<int> { };
        public int Текущий_приоритет = 0;
        public bool Аттестат;
        public double Балл_аттестата;
        public int Заслуги;
        public bool Запас = false;
        public bool not_stud = false;

        public Абитуриент(int a, Excel.Workbook Ab)
        {
            вОчереди = true;
            Excel.Worksheet sheet1 = (Excel.Worksheet)Ab.Worksheets.get_Item(1);
            ФИО = (sheet1.Cells[a, 2] as Excel.Range).Value2.ToString();
            for (int i = 3; i < 21; i++)
            {
                Предметы[i - 3] = Predmet(a, i, sheet1);
            }


            try
            {
                Балл_аттестата = Convert.ToDouble(sheet1.Cells[a, 21].Value2.ToString());
            }
            catch { Балл_аттестата = 0; }

            try
            {
                Заслуги = Convert.ToInt32(sheet1.Cells[a, 22].Value2.ToString());
            }
            catch { Заслуги = 0; }


            Аттестат = Convert.ToBoolean(sheet1.Cells[a, 1].Value2.ToString());
            for (int i = 23; i < 26; i++)
            {
                try
                {
                    if ((sheet1.Cells[a, i] as Excel.Range).Interior.Color == ColorTranslator.ToOle(Color.FromArgb(0, 0, 255, 0)))
                    {
                        Предпочтения.Insert(0, Convert.ToInt32((sheet1.Cells[a, i] as Excel.Range).Value2.ToString()) - 3);
                    }
                    else Предпочтения.Add(Convert.ToInt32((sheet1.Cells[a, i] as Excel.Range).Value2.ToString()) - 3);
                }
                catch { }
            }

        }


        int Predmet(int Номер_студента, int Номер_предмета, Excel.Worksheet Ab)
        {
            try
            {
                return Convert.ToInt32((Ab.Cells[Номер_студента, Номер_предмета] as Excel.Range).Value2.ToString());
            }
            catch
            {
                return 0;
            }
        }

        public void COUT()
        {
            Console.WriteLine();
            Console.WriteLine(ФИО);
            for (int i = 0; i < 18; i++)
                Console.Write(Предметы[i] + " ");
            Console.WriteLine();
            for (int i = 0; i < Предпочтения.Count; i++)
                Console.Write(Предпочтения[i] + " ");
        }

        public void Plus_Priority()
        {
            if (Текущий_приоритет + 1 < Предпочтения.Count) { Текущий_приоритет++; вОчереди = true; }
            else { вОчереди = false; not_stud = true; }
        }

        void вывод_студента(int a)
        {
            
        }
    }
}
