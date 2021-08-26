using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using V83;
using System.Diagnostics;
using System.Threading;
using System.Runtime.InteropServices;



namespace UpdatingTool
{
    class Program
    {

        [DllImport("user32.dll")]
        public static extern int DeleteMenu(IntPtr hMenu, int nPosition, int wFlags);

        [DllImport("user32.dll")]
        private static extern IntPtr GetSystemMenu(IntPtr hWnd, bool bRevert);

        [DllImport("kernel32.dll", ExactSpelling = true)]
        private static extern IntPtr GetConsoleWindow();

        private const int MF_BYCOMMAND = 0x00000000;
        public const int SC_CLOSE = 0xF060;
        public const int SC_MINIMIZE = 0xF020;
        public const int SC_MAXIMIZE = 0xF030;
        public const int SC_SIZE = 0xF000;

        public static bool timerFlag = true;
        public static DateTime UpdatingStart;
        public static int posX;
        public static int posY;
        static void Main(string[] args)
        {
            Console.BackgroundColor = ConsoleColor.Black;

            IntPtr handle = GetConsoleWindow();
            IntPtr sysMenu = GetSystemMenu(handle, false);

            if (handle != IntPtr.Zero)
            {
                DeleteMenu(sysMenu, SC_MINIMIZE, MF_BYCOMMAND);
                DeleteMenu(sysMenu, SC_MAXIMIZE, MF_BYCOMMAND);
                DeleteMenu(sysMenu, SC_SIZE, MF_BYCOMMAND);
            }

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

                Com1sObject = com1s.Connect("Srvr='vs-nzp';Ref='nzp';Usr='BarCodesUploadUser';Pwd='852456';");

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

                    if (handle != IntPtr.Zero)
                    {
                        DeleteMenu(sysMenu, SC_CLOSE, MF_BYCOMMAND);
                    }

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

                    if (handle != IntPtr.Zero)
                    {
                        DeleteMenu(sysMenu, SC_CLOSE, MF_BYCOMMAND);
                    }

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

            Console.ReadKey();
        }

        public static int StartInformation()
        {
            Console.SetWindowSize(98, 40);
            Console.ForegroundColor = ConsoleColor.DarkGreen;
            Console.WriteLine("*********************************DATABASE UPDATING TOOL v1.0**************************************");
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
            Console.WriteLine("*********************************DATABASE UPDATING TOOL v1.0**************************************");
            Console.WriteLine("          -----------------------------------------------------------------------------");
            Console.WriteLine("         |                                #INFO#                                       |");
            Console.WriteLine("         | * Press 1 to update NZP@Barcodes and SS-BD@kzmp_stickers_NiceLabel          |");
            Console.WriteLine("         | * Press 2 to update NZP@Nomenclature and SS-BD@kzmp_stickers_NiceLabel      |");
            Console.WriteLine("         |   (Warning: mode number 2 causes a heavy load on the system )               |");
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
            Console.SetCursorPosition(posX, posY);
            for (int i = 0; i < 60; i++)
            {
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
