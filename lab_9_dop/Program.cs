using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;
using System.Diagnostics;
using System.Threading.Tasks;

namespace lab_9_dop
{
    class Program
    {
         static void Count(object mas)
        {
            int[] arr = (int[])mas;
            int temp;
            for (int i = 0; i < arr.Length - 1; i++)
            {
                for (int j = i + 1; j <arr.Length; j++)
                {
                    if (arr[i] > arr[j])
                    {
                        temp = arr[i];
                        arr[i] = arr[j];
                        arr[j] = temp;
                    }
                }
            }
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("Вывод отсортированного массива");
            for (int i = 0; i < arr.Length; i++)
            {
                Console.WriteLine(arr[i]);
            }
        }
        static void Main(string[] args)
        {
            Excel.Application ex = new Microsoft.Office.Interop.Excel.Application();//Объявляем приложение
            ex.Visible = true;//Отобразить Excel
            ex.SheetsInNewWorkbook = 1;//Количество листов в рабочей книге
            Excel.Workbook workBook = ex.Workbooks.Add(Type.Missing); //Добавить рабочую книгу
            ex.DisplayAlerts = false;//Отключить отображение окон с сообщениями
            Excel.Worksheet sheet = (Excel.Worksheet)ex.Worksheets.get_Item(1);//Получаем первый лист документа (счет начинается с 1)
            sheet.Name = "отчет";//Название листа (вкладки снизу)
            sheet.Cells[1,1] = String.Format("без потока ");
            sheet.Cells[1,2] = String.Format("с потоком");

            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("введи число испытаний");
            int N = Convert.ToInt32(Console.ReadLine());
            for (int y = 1; y <=N; y++)
            {
                //Console.ForegroundColor = ConsoleColor.Yellow;
                //Console.WriteLine("введи размер массива");
                //int M = Convert.ToInt32(Console.ReadLine());

                int[] mas = new int[y]; 
                Random rand = new Random();
                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine("массив №"+y);
                for (int x = 0; x < y; x++)
                {
                    mas[x] = rand.Next(100);
                    Console.WriteLine(mas[x]);
                }


                Stopwatch stopwatch1 = new Stopwatch();//создаем объект
                stopwatch1.Start();//засекаем время начала операции
                Console.ForegroundColor = ConsoleColor.Green;
                int temp;
                for (int i = 0; i < mas.Length - 1; i++)
                {
                    for (int j = i + 1; j < mas.Length; j++)
                    {
                        if (mas[i] > mas[j])
                        {
                            temp = mas[i];
                            mas[i] = mas[j];
                            mas[j] = temp;
                        }
                    }
                }

                Console.WriteLine("Вывод отсортированного массива");
                for (int i = 0; i < mas.Length; i++)
                {
                    Console.WriteLine(mas[i]);
                }
                stopwatch1.Stop();//останавливаем счётчик
                int Q = Convert.ToInt32(stopwatch1.ElapsedMilliseconds);
                Console.ForegroundColor = ConsoleColor.Magenta;
                Console.WriteLine("Время работы программы без потока " + Q);//смотрим сколько миллисекунд было затрачено на выполнение




                
                //Thread thread = new Thread(Count);
                Stopwatch stopwatch = new Stopwatch();//создаем объект
                stopwatch.Start();//засекаем время начала операции
                Parallel.Invoke(() => Count(mas));
                //thread.Start(mas);
                Thread.Sleep(1000);
                stopwatch.Stop();//останавливаем счётчик
                int R = Convert.ToInt32(stopwatch.ElapsedMilliseconds)-1000;
                Console.ForegroundColor = ConsoleColor.Magenta;
                Console.WriteLine("Время работы программы с потоком "+R);//смотрим сколько миллисекунд было затрачено на выполнение


                sheet.Cells[y+1,1] = Q;
                sheet.Cells[y+1,2] = R;
            }
            Excel.Range chartRange;

            Excel.ChartObjects xlCharts = (Excel.ChartObjects)sheet.ChartObjects(Type.Missing);
            Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(70, 30, 500, 400);
            Excel.Chart chartPage = myChart.Chart;
            chartRange = sheet.get_Range("A1", "B600");//Задаем ячейки данных для графика
            chartPage.SetSourceData(chartRange, Type.Missing);
            chartPage.ChartType = Excel.XlChartType.xlLineStacked;

        }
    }
}
