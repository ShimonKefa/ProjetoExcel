using ClosedXML.Excel;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using DocumentFormat.OpenXml.EMMA;

namespace relacaoComputadores
{
    public class Excel
    {
        //define o caminho até o arquivo excel
        private string PATH = "D:\\ComputadoresEXCEL\\RELACAOCOMPUTADORES.xlsx";
        private XLWorkbook WORKBOOK;
        private IXLWorksheet PLAN;
        private int TOTALLINES;
        private List<PARAMETERS> Computadores = new List<PARAMETERS>();

        public void ExcelComputadores(string[] args)
        {
            while (true)
            {
                try
                {
                    Console.Clear();
                    //Cria uma instancia para o arquivo e direciona para o caminho path
                    WORKBOOK = new XLWorkbook(PATH);
                    //cria uma instancia para a planilia que vai ser alterada 
                    PLAN = WORKBOOK.Worksheets.FirstOrDefault(w => w.Name == "Plan1");
                    //faz uma contagem da quantidade de células preencida permitindo que preencha a última 
                    TOTALLINES = PLAN.RowsUsed().Count();


                    Console.WriteLine("RELAÇÃO COMPUTADORES");
                    Console.WriteLine("1 - inserir pessoas na planilia\n2 - procurar Ordem de serviços na planilia\n3 - Sair");
                    var aux1 = Convert.ToInt16(Console.ReadLine());

                    if (PLAN == null)
                    {
                        Console.WriteLine("Planilia não encontrada");
                    }

                    else
                    {

                        if (aux1 == 1)
                        { //inserir pessoas na planilia                
                            Console.WriteLine("Quantos computadores a serem listados");
                            var listComputadores = Convert.ToInt16(Console.ReadLine());
                            //eixo    x  y    
                            PLAN.Cell(1, 1).Value = "Nome";
                            PLAN.Cell(1, 2).Value = "Os";
                            PLAN.Cell(1, 3).Value = "Model";
                            PLAN.Cell(1, 4).Value = "Memory";
                            PLAN.Cell(1, 5).Value = "Storage";


                            for (int i = TOTALLINES + 1; i < TOTALLINES + 1 + listComputadores; i++)
                            {
                                Console.WriteLine($"Nome: {i}");
                                var nome = Console.ReadLine();

                                Console.WriteLine($"OS: {i}");
                                var os = Convert.ToInt16(Console.ReadLine());

                                Console.WriteLine($"model: {i}");
                                var model = Console.ReadLine();

                                Console.WriteLine($"memory: {i}");
                                var memory = Console.ReadLine();

                                Console.WriteLine($"storage: {i}");
                                var storage = Console.ReadLine();

                                PLAN.Cell(i, 1).Value = nome;
                                PLAN.Cell(i, 2).Value = os;
                                PLAN.Cell(i, 3).Value = model;
                                PLAN.Cell(i, 4).Value = memory;
                                PLAN.Cell(i, 5).Value = storage;


                                Console.Clear();

                                PARAMETERS NovoPC = new PARAMETERS { NAME = nome, OS = os, MODEL = model, MEMORY = memory, STORAGE = storage };
                                Computadores.Add(NovoPC);
                            }
                            WORKBOOK.Save();

                            Console.Clear();
                            foreach (var PC in Computadores)
                            {
                                Console.WriteLine($"Nome: {PC.NAME} - OS: {PC.OS} - Modelo: {PC.MODEL} - Memória: {PC.MEMORY} - Armazenamento: {PC.STORAGE}\n");
                            }
                        }

                        else if (aux1 == 2)
                        {
                            Console.WriteLine("Digite o número da OS que você precisa encontrar");
                            //aux3 procura a OS
                            int aux3;
                            while (!int.TryParse(Console.ReadLine(), out aux3))
                            {
                                Console.WriteLine("Digite um número válido para a Busca");
                            }


                            bool encontrado = false;

                            for (int i = 2; i <= TOTALLINES; i++)
                            {
                                var Os = PLAN.Cell($"B{i}").GetValue<int>();
                                if (Os == aux3)
                                {

                                    var Name = PLAN.Cell($"A{i}").GetValue<string>();
                                    var Model = PLAN.Cell($"C{i}").GetValue<string>();
                                    var Memory = PLAN.Cell($"D{i}").GetValue<string>();
                                    var Storage = PLAN.Cell($"E{i}").GetValue<string>();

                                    PARAMETERS novoParametro = new PARAMETERS { NAME = Name, OS = Os, MODEL = Model, MEMORY = Memory, STORAGE = Storage };
                                    Computadores.Add(novoParametro);

                                    foreach (var leitura in Computadores)
                                    {
                                        Console.WriteLine($"Nome: {leitura.NAME} - OS: {leitura.OS} - Modelo: {leitura.MODEL} - Memória: {leitura.MEMORY} - Armazenamento: {leitura.STORAGE}\n");
                                        encontrado = true;
                                    }
                                }
                                else if (!encontrado)
                                {
                                    Console.WriteLine("OS Não encontrada");
                                }
                                Console.WriteLine("precione qualuer tecla");
                                Console.ReadKey();
                                Console.Clear();
                            }
                        }

                        else if (aux1 == 3)
                        {
                            Console.WriteLine("Saindo do Sistema ...");
                            Thread.Sleep(3000);
                            System.Environment.Exit(0);
                        }

                    }

                }
                catch (Exception ex)
                {
                    Console.WriteLine($"ocorreu um erro: {ex.Message}");
                }
            }
        }

    }

    public class PARAMETERS
    {
        public string NAME { get; set; }
        public int OS { get; set; }
        public string MODEL { get; set; }
        public string MEMORY { get; set; }
        public string STORAGE { get; set; }
    }


}