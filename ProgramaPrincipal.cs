using System;
using Menu;
using relacaoComputadores;



namespace ProjetoConect
{
    public class Program
    {
        static void Main(string[] args)
        {
            Console.Clear();
            //instancia do objeto menu para poder acessar a classe no outro código
            Excel excel = new Excel();
            StartMenu menu = new StartMenu();
            //captura o valor retornado do outro código
            int op = menu.Header(args);

            switch (op)
            {
                case 1:
                    excel.ExcelComputadores(args);
                    break;

                default:
                    Console.WriteLine("Operação inválida");
                    break;



            }
        }

    }
}