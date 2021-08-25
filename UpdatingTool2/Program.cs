using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using V83;
using System.Diagnostics;
using System.Threading;


namespace UpdatingTool
{
    class Program
    {
        public static bool timerFlag = true;
        public static DateTime UpdatingStart;
        public static int posX;
        public static int posY;
        static void Main(string[] args)
        {
            int mode = StartInformation();

            TimerCallback tm = new TimerCallback(TimerProcess);

            V83.COMConnector com1s = new V83.COMConnector();
            dynamic Com1sObject = null;

            com1s.PoolCapacity = 10;
            com1s.PoolTimeout = 60;
            com1s.MaxConnections = 2;

            try
            {
                Console.ForegroundColor = ConsoleColor.Magenta;
                Console.Write("@logs> ");
                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine("Connection attempt...");

                Com1sObject = com1s.Connect("Srvr='ws-oit-007';Ref='NZP_local2';Usr='test';Pwd='852456';");

                Console.ForegroundColor = ConsoleColor.Magenta;
                Console.Write("@logs> ");
                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine("The connection was successfully established!");

                if (mode == 1)
                {
                    Console.ForegroundColor = ConsoleColor.Magenta;
                    Console.Write("@logs> ");
                    Console.ForegroundColor = ConsoleColor.Blue;
                    Console.WriteLine("NZP@Barcodes and SS-BD@kzmp_stickers_NiceLabel updating...");

                    posX = Console.CursorLeft;
                    posY = Console.CursorTop;

                    UpdatingStart = DateTime.Now;
                    Timer _timer = new Timer(tm, null, 0, 1000);

                    Com1sObject.СинхронизацияДанных.ОбновитьШтрихкодыНоменклатур();
                    Com1sObject.СинхронизацияДанных.ОбновлениеБазыШтрихкодовMSSQL();

                    _timer.Dispose();
                    Console.WriteLine();

                    Console.ForegroundColor = ConsoleColor.Magenta;
                    Console.Write("@logs> ");
                    Console.ForegroundColor = ConsoleColor.Blue;
                    Console.WriteLine("Successful update!");

                }
                else if (mode == 2)
                {
                    Console.ForegroundColor = ConsoleColor.Magenta;
                    Console.Write("@logs> ");
                    Console.ForegroundColor = ConsoleColor.Blue;
                    Console.WriteLine("NZP@Nomenclature updating...");

                    posX = Console.CursorLeft;
                    posY = Console.CursorTop;

                    UpdatingStart = DateTime.Now;
                    Timer _timer = new Timer(tm, null, 0, 1000);

                    Com1sObject.СинхронизацияДанных.СинхронизацияНоменклатур();
                    Com1sObject.СинхронизацияДанных.ОбновлениеБазыШтрихкодовMSSQL();

                    _timer.Dispose();
                    Console.WriteLine();

                    Console.ForegroundColor = ConsoleColor.Magenta;
                    Console.Write("@logs> ");
                    Console.ForegroundColor = ConsoleColor.Blue;
                    Console.WriteLine("Successful update!");

                }
            }
            catch (Exception ex)
            {
                //Console.ForegroundColor = ConsoleColor.Magenta;
                //Console.Write("@logs> ");
                //Console.ForegroundColor = ConsoleColor.Red;
                //Console.WriteLine("Error: " + ex.Message);

                //Com1sObject = null;
                //com1s = null;

                //Environment.Exit(0);


                var trace = new StackTrace(ex, true);


                foreach (var frame in trace.GetFrames())
                {
                    Console.ForegroundColor = ConsoleColor.Magenta;
                    Console.Write("@logs> ");

                    var sb = new StringBuilder();

                    sb.AppendLine($"Файл: {frame.GetFileName()}");
                    sb.AppendLine($"Строка: {frame.GetFileLineNumber()}");
                    sb.AppendLine($"Столбец: {frame.GetFileColumnNumber()}");
                    sb.AppendLine($"Метод: {frame.GetMethod()}");

                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Error: " + sb);
                }

            }

            Com1sObject = null;
            com1s = null;

            Environment.Exit(0);

        }

        public static int StartInformation()
        {
            Console.SetWindowSize(98, 40);
            Console.ForegroundColor = ConsoleColor.DarkGreen;
            Console.WriteLine("*********************************DATABASE UPDATING TOOL v1.0****************************************");
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.WriteLine("░░░████▌█████▌█░████████▐▀██▀\t" + "░░░████▌█████▌█░████████▐▀██▀\t" + "░░░████▌█████▌█░████████▐▀██▀\n" +
                          "░▄█████░█████▌░█░▀██████▌█▄▄▀▄\t" + "░▄█████░█████▌░█░▀██████▌█▄▄▀▄\t" + "░▄█████░█████▌░█░▀██████▌█▄▄▀▄\n" +
                          "░▌███▌█░▐███▌▌░░▄▄░▌█▌███▐███░▀\t" + "░▌███▌█░▐███▌▌░░▄▄░▌█▌███▐███░▀\t" + "░▌███▌█░▐███▌▌░░▄▄░▌█▌███▐███░▀\n" +
                          "░▐██░░▄▄▐▀█░░░▐▄█▀▌█▐███▐█\t" + "░▐██░░▄▄▐▀█░░░▐▄█▀▌█▐███▐█\t" + "░▐██░░▄▄▐▀█░░░▐▄█▀▌█▐███▐█\n" +
                          "░░███░▌▄█▌░░▀░░▀██░░▀██████\t" + "░░███░▌▄█▌░░▀░░▀██░░▀██████\t" + "░░███░▌▄█▌░░▀░░▀██░░▀██████\n" +
                          "░░░▀█▌▀██▀░▄░░░░░░░░░███▐███\t" + "░░░▀█▌▀██▀░▄░░░░░░░░░███▐███\t" + "░░░▀█▌▀██▀░▄░░░░░░░░░███▐███\n" +
                          "░░░░██▌░░░░░░░░░░░░░▐███████\t" + "░░░░██▌░░░░░░░░░░░░░▐███████\t" + "░░░░██▌░░░░░░░░░░░░░▐███████\n" +
                          "░░░░███░░░░░▀█▀░░░░░▐██▐███▀\t" + "░░░░███░░░░░▀█▀░░░░░▐██▐███▀\t" + "░░░░███░░░░░▀█▀░░░░░▐██▐███▀\n" +
                          "░░░░▌█▌█▄░░░░░░░░░▄▄████▀░▀\t" + "░░░░▌█▌█▄░░░░░░░░░▄▄████▀░▀\t" + "░░░░▌█▌█▄░░░░░░░░░▄▄████▀░▀\n" +
                          "░░░░░░█▀██▄▄▄░▄▄▀▀▒█▀█░\t\t" + "░░░░░░█▀██▄▄▄░▄▄▀▀▒█▀█░\t\t" + "░░░░░░█▀██▄▄▄░▄▄▀▀▒█▀█░\n");

            Console.ForegroundColor = ConsoleColor.DarkGreen;
            Console.WriteLine("*********************************DATABASE UPDATING TOOL v1.0****************************************");
            Console.WriteLine("          -----------------------------------------------------------------------------");
            Console.WriteLine("         |                                #INFO#                                       |");
            Console.WriteLine("         | * Press 1 to update NZP@Barcodes and SS-BD@kzmp_stickers_NiceLabel          |");
            Console.WriteLine("         | * Press 2 to update NZP@Nomenclature and SS-BD@kzmp_stickers_NiceLabel (Warn|");
            Console.WriteLine("         |   ing: high load on the system )                                            |");
            Console.WriteLine("         | * Press 0 to exit                                                           |");
            Console.WriteLine("         |                                                                             |");
            Console.WriteLine("          ----------------------------------------------------------------------------");


            while (true)
            {
                try
                {
                    Console.ForegroundColor = ConsoleColor.Magenta;
                    Console.Write("@robot> ");
                    Console.ForegroundColor = ConsoleColor.Blue;
                    Console.Write("Select the mode: ");
                    int mode = Convert.ToInt32(Console.ReadLine());

                    if (mode == 1 || mode == 2)
                    {
                        return mode;
                    }
                    else if (mode == 0)
                    {
                        Environment.Exit(0);
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Magenta;
                        Console.Write("@robot> ");
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine("Warning: Incorrect input");
                    }
                }
                catch (Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.Magenta;
                    Console.Write("@robot> ");
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Error: " + ex.Message);
                }
            }
        }

        public static void TimerProcess(object obj)
        {
            for (int i = 0; i < 60; i++)
            {
                Console.SetCursorPosition(posX, posY);
                Console.Write("");
            }

            Console.SetCursorPosition(posX, posY);
            TimeSpan diff1 = DateTime.Now.Subtract(UpdatingStart);

            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.Write("@logs> ");
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.Write("Timer of the update process: " + Convert.ToString(diff1));

        }
    }
}
