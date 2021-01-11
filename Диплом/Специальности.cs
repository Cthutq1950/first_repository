using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using Microsoft.VisualBasic;
using Excel = Microsoft.Office.Interop.Excel;


namespace Диплом_С_шарп
{
    class Специальности
    {
        public string Название;
        public int[] Предметы = new int[18];
        List<int> основные_предмет = new List<int> { };
        int[] осн_предметы = new int[3] { -1, -1, -1 };
        public Абитуриент[] Места;
        public int Количество_мест;
        int Баллы = 0;
        int o = 3;

        public Специальности(int a, Excel.Workbook Ab_w)
        {
            Excel.Worksheet Ab = (Excel.Worksheet)Ab_w.Worksheets.get_Item(1);
            Название = (Ab.Cells[a, 2] as Excel.Range).Value2.ToString();
            string s = (Ab.Cells[a, 1] as Excel.Range).Value2.ToString();
            Количество_мест = Convert.ToInt32((Ab.Cells[a, 1] as Excel.Range).Value2.ToString());
            Места = new Абитуриент[Количество_мест];


            for (int i = 3; i < 21; i++)
            {
                Предметы[i - 3] = Predmet(a, i, Ab);
                if (Предметы[i - 3] != 0)
                {

                    // основные_предметы.Add(i - 3);
                    Баллы += Предметы[i - 3];
                }

            }
        }

        int Predmet(int Номер_студента, int Номер_предмета, Excel.Worksheet Ab)
        {
            int a;
            try
            {
                a = Convert.ToInt32((Ab.Cells[Номер_студента, Номер_предмета] as Excel.Range).Value2.ToString());
            }
            catch
            {
                a = 0;
            }


            if (a != 0)
            {
                if ((Ab.Cells[Номер_студента, Номер_предмета] as Excel.Range).Interior.Color == ColorTranslator.ToOle(Color.FromArgb(0, 0, 255, 0)))
                {
                    осн_предметы[0] = Номер_предмета - 3;
                }
                if ((Ab.Cells[Номер_студента, Номер_предмета] as Excel.Range).Interior.Color == ColorTranslator.ToOle(Color.FromArgb(0, 255, 255, 0)))
                {
                    осн_предметы[1] = Номер_предмета - 3;

                }
                if ((Ab.Cells[Номер_студента, Номер_предмета] as Excel.Range).Interior.Color == ColorTranslator.ToOle(Color.FromArgb(0, 255, 0, 0)))
                {
                    осн_предметы[2] = Номер_предмета - 3;

                }
            }
            return a;
        }


        public void COUT()
        {
            Console.WriteLine();
            Console.WriteLine(Название);
            for (int i = 0; i < 18; i++)
                Console.Write(Предметы[i] + " ");
            Console.WriteLine();
            for (int i = 0; i < o; i++)
                Console.Write(осн_предметы[i]);
        }

        public Абитуриент Add(Абитуриент абитуриент)
        {
            if (Места[Количество_мест - 1] != null)
            {
                Места[Количество_мест - 1].Plus_Priority();
            }
            Абитуриент A = Места[Количество_мест - 1];
            Места[Количество_мест - 1] = абитуриент;
            Sort(Места);

            return A;
        }


        public bool Check1(Абитуриент абитуриент)
        {
            for (int i = 0; i < Количество_мест; i++)
            {
                if (Места[i] == null) return true;
            }



            return false;
        }

        public bool Check2(Абитуриент абитуриент)
        {
            if ((Sum(абитуриент) >= Баллы) && (Sum(Места[Количество_мест - 1])) < Sum(абитуриент)) return true;
            else if (Sum(Места[Количество_мест - 1]) == Sum(абитуриент)) { return (find_better(Места[Количество_мест - 1], абитуриент)); }
            return false;
        }

        public bool Check3(Абитуриент абитуриент)
        {
            for (int i = 0; i < o; i++)
            {
                if (абитуриент.Предметы[осн_предметы[i]] < Предметы[осн_предметы[i]]) { return false; }
            }
            return true;
        }

        bool find_better(Абитуриент a, Абитуриент b)
        {
            for (int i = 0; i < o; i++)
            {
                if (осн_предметы[i] != -1)
                    if (a.Предметы[осн_предметы[i]] < b.Предметы[осн_предметы[i]]) { return true; }
                    else if (a.Предметы[осн_предметы[i]] != b.Предметы[осн_предметы[i]]) { return false; }
            }
            if (a.Балл_аттестата < b.Балл_аттестата) return true;
            else if (a.Балл_аттестата != b.Балл_аттестата) return false;

            return false;

        }

        int Sum(Абитуриент абитуриент)
        {
            int sum = 0;
            if (абитуриент != null)
                for (int i = 0; i < o; i++)
                {
                    sum += абитуриент.Предметы[осн_предметы[i]];
                }
            else
            {
                return 0;
            }
            sum += абитуриент.Заслуги;
            return sum;
        }

        Абитуриент[] Sort(Абитуриент[] a)
        {
            int max, place;
            int[] b = new int[Количество_мест];
            for (int i = 0; i < Количество_мест; i++)
            {
                b[i] = Sum(a[i]);
            }

            for (int i = 0; i < Количество_мест; i++)
            {
                place = i; max = b[i];
                for (int j = i; j < Количество_мест; j++)
                {
                    if (max < b[j]) { max = b[j]; place = j; }

                }
                int k1 = b[i]; Абитуриент k2 = a[i];
                b[i] = b[place]; a[i] = a[place];
                b[place] = k1; a[place] = k2;
            }

            return a;
        }

        public void Вывод_запаса(int a, Excel.Worksheet sheet)
        {
            sheet.Cells[a + 1, 1].Value = Название;
            int k = 2;
            for (int i = 0; i < Количество_мест; i++)
            {
                if (Места[i] != null && Места[i].Запас) { sheet.Cells[a + 1, k].Value = Места[i].ФИО; k++; }
            }
        }

        public void Vivod_mest(int a, Excel.Worksheet sheet)
        {

            sheet.Cells[a + 1, 1].Value = Название;
            for (int i = 0; i < Количество_мест; i++)
            {
                if (Места[i] != null) sheet.Cells[a + 1, i + 2].Value = Места[i].ФИО;
            }

        }
    }
}
