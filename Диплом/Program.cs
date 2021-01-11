using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;


namespace Диплом_С_шарп
{
    class Program
    {
        static void Main(string[] args)
        {
            int Количество_абитуриентов;
            Excel.Application Ab = new Microsoft.Office.Interop.Excel.Application();
            Excel.Workbook Ab_w = Ab.Workbooks.Open(@"C:\Users\Сергей\Desktop\ДИПЛОООМ!\Диплом\Абитуриенты.xlsx",
  Type.Missing, Type.Missing, Type.Missing, Type.Missing,
  Type.Missing, Type.Missing, Type.Missing, Type.Missing,
  Type.Missing, Type.Missing, Type.Missing, Type.Missing,
  Type.Missing, Type.Missing);

            Количество_абитуриентов = Convert.ToInt32((Ab.Cells[1, 2] as Excel.Range).Value2.ToString());

            Абитуриент[] абитуриентs = new Абитуриент[Количество_абитуриентов];
            for (int i = 3; i < 3 + Количество_абитуриентов; i++)
            {
                абитуриентs[i - 3] = new Абитуриент(i, Ab_w);
            }

            //for (int i = 0; i < Количество_абитуриентов; i++)
            //    абитуриентs[i].COUT();

            int Количество_специальностей;
            Ab_w.Close(true);
            Ab.Quit();

            Excel.Application Spec = new Microsoft.Office.Interop.Excel.Application();
            Excel.Workbook Spec_w = Spec.Workbooks.Open(@"C:\Users\Сергей\Desktop\ДИПЛОООМ!\Диплом\Специальности.xlsx",
  Type.Missing, Type.Missing, Type.Missing, Type.Missing,
  Type.Missing, Type.Missing, Type.Missing, Type.Missing,
  Type.Missing, Type.Missing, Type.Missing, Type.Missing,
  Type.Missing, Type.Missing);

            Количество_специальностей = Convert.ToInt32((Spec.Cells[1, 2] as Excel.Range).Value2.ToString());

            Специальности[] специальности = new Специальности[Количество_специальностей];
            for (int i = 3; i < 3 + Количество_специальностей; i++)
            {
                специальности[i - 3] = new Специальности(i, Spec_w);
            }

            bool b = true;
            List<Абитуриент> not_stud = new List<Абитуриент> { };
            while (b)
            {
                b = false;
                for (int i = 0; i < Количество_абитуриентов; i++)
                {
                    if (абитуриентs[i].Текущий_приоритет == 0 && абитуриентs[i].Аттестат)
                    {
                        if ((абитуриентs[i].вОчереди && (специальности[абитуриентs[i].Предпочтения[абитуриентs[i].Текущий_приоритет]].Check1(абитуриентs[i])
                            || специальности[абитуриентs[i].Предпочтения[абитуриентs[i].Текущий_приоритет]].Check2(абитуриентs[i])))
                            && специальности[абитуриентs[i].Предпочтения[абитуриентs[i].Текущий_приоритет]].Check3(абитуриентs[i]))
                        {
                            специальности[абитуриентs[i].Предпочтения[абитуриентs[i].Текущий_приоритет]].Add(абитуриентs[i]);
                            абитуриентs[i].вОчереди = false;
                            b = true;
                        }
                        else
                        {
                            if (абитуриентs[i].вОчереди)
                            {
                                абитуриентs[i].Plus_Priority();
                                not_stud.Add(абитуриентs[i]);
                                абитуриентs[i].Запас = true;
                            }
                        }
                    }
                    else if (абитуриентs[i].Аттестат == false) { not_stud.Add(абитуриентs[i]); абитуриентs[i].Запас = true; }
                }
            }




            Spec_w.Close(true);
            Spec.Quit();
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Add();

            Excel.Worksheet sheet = ObjWorkBook.Sheets[1];
            for (int i = 0; i < Количество_специальностей; i++)
            {
                специальности[i].Vivod_mest(i, sheet);
            }
            Console.WriteLine();

            ObjWorkBook.SaveAs(@"C:\Users\Сергей\Desktop\ДИПЛОООМ!\Диплом\Итог.xlsx");
            ObjWorkBook.Close(true);
            ObjWorkExcel.Quit();

            b = true;

            Excel.Application ObjWorkExcel1 = new Excel.Application();
            Excel.Workbook ObjWorkBook1 = ObjWorkExcel1.Workbooks.Add();

            Excel.Worksheet перевод = ObjWorkBook1.Sheets[1];
            Абитуриент A;

            List<Абитуриент> isNot_stud = new List<Абитуриент> { };

            while (b)
            {
                b = false;
                for (int i = 0; i < not_stud.Count; i++)
                {
                    if (not_stud[i].вОчереди && (специальности[not_stud[i].Предпочтения[not_stud[i].Текущий_приоритет]].Check1(not_stud[i])
                       || специальности[not_stud[i].Предпочтения[not_stud[i].Текущий_приоритет]].Check2(not_stud[i]))
                        && (специальности[not_stud[i].Предпочтения[not_stud[i].Текущий_приоритет]].Check3(not_stud[i])))
                    {
                        A = специальности[not_stud[i].Предпочтения[not_stud[i].Текущий_приоритет]].Add(not_stud[i]);
                        if (A != null)
                        {
                            A.Запас = true;
                            not_stud.Add(A);
                        }
                        not_stud[i].вОчереди = false;
                        b = true;
                    }
                    else
                    {
                        if (not_stud[i].вОчереди)
                        {
                            not_stud[i].Plus_Priority();
                            if (not_stud[i].not_stud) isNot_stud.Add(not_stud[i]);
                        }
                    }

                }
            }

            for (int i = 0; i < Количество_специальностей; i++)
            {
                специальности[i].Вывод_запаса(i, перевод);
            }

            Console.WriteLine();

            ObjWorkBook1.SaveAs(@"C:\Users\Сергей\Desktop\ДИПЛОООМ!\Диплом\Запас.xlsx");
            ObjWorkBook1.Close(true);
            ObjWorkExcel1.Quit();

            Excel.Application Вывод_остальных = new Excel.Application();
            Excel.Workbook workbook = Вывод_остальных.Workbooks.Add();
            Excel.Worksheet sheet1 = (Excel.Worksheet)Вывод_остальных.Worksheets.get_Item(1);
            for (int i = 0; i < isNot_stud.Count; i++)
            {
                sheet1.Cells[i + 1, 1] = isNot_stud[i].ФИО;
            }

 
            workbook.SaveAs(@"C:\Users\Сергей\Desktop\ДИПЛОООМ!\Диплом\Кто не прошел.xlsx");
            workbook.Close(true);
            Вывод_остальных.Quit();

            Console.ReadLine();



        }


    }
}