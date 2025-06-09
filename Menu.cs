using System;
using DocumentFormat.OpenXml.Bibliography;

namespace Menu
{
    public class StartMenu
    {
        public int Header(string[] args)
        {
            //op é a váriavel de operação do menu
            int op = 0;
            Console.Clear();
            Console.BackgroundColor = ConsoleColor.DarkGray;
            Console.ForegroundColor = ConsoleColor.White;

            for (int i = 0; i < 28; i++)
            {
                Console.Write("*");
            }
            Console.WriteLine("\n* Listagem de computadores *\n");
            for (int i = 0; i < 28; i++)
            {
                Console.Write("*");
            }

            Console.WriteLine("\n");

            Console.WriteLine("Digite a operação");
            Console.WriteLine("1 - Relação de computadores no excel");
            op = Convert.ToInt32(Console.ReadLine());

            return op;       
        } 

    }
}