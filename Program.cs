using Log.Convert;

namespace Log
{
    internal class Program
    {
        static void Main(string[] args)
        {
            ChamaTela();
        }

        public static void ChamaTela()
        {
            var opcao = Menu();
            EscolheProjeto(opcao);
        }

        public static void EscolheProjeto(string opcao)
        {
            switch (opcao)
            {
                case "1":
                    string strFileNameEnricher = @"C:\Users\rogarcia\OneDrive - ScanSource, Inc\Documentos\ProductEnricherConsumer.xlsx";
                    using (var reader = new StreamReader(strFileNameEnricher))
                    {
                        ConvertCsv.ConvertErrosProductEnricher(reader.BaseStream);
                        ChamaTela();
                    }
                    break;
                case "2":
                    string strFileNameSync = @"C:\Users\rogarcia\OneDrive - ScanSource, Inc\Documentos\ProtheusProductSyncConsumer.xlsx";
                    using (var reader = new StreamReader(strFileNameSync))
                    {
                        ConvertCsv.ConvertErrosProductSync(reader.BaseStream);
                        ChamaTela();
                    }
                    break;
                case "3":
                    string strFileNameNational = @"C:\Users\rogarcia\OneDrive - ScanSource, Inc\Documentos\ProtheusNationalPurchaseSyncConsumer.xlsx";
                    using (var reader = new StreamReader(strFileNameNational))
                    {
                        ConvertCsv.ConvertErrosNationalPurchaseSync(reader.BaseStream, "ProtheusNationalPurchaseSyncConsumer");
                        ChamaTela();
                    }
                    break;
                case "4":
                    string strFileNameIntangivel = @"C:\Users\rogarcia\OneDrive - ScanSource, Inc\Documentos\ProtheusIntangiblePurchaseSyncConsumer.xlsx";
                    using (var reader = new StreamReader(strFileNameIntangivel))
                    {
                        ConvertCsv.ConvertErrosNationalPurchaseSync(reader.BaseStream, "ProtheusIntangiblePurchaseSyncConsumer");
                        ChamaTela();
                    }
                    break;
                case "5":
                    Environment.Exit(0);
                    break;
                default:
                    break;
            }
        }

        public static string Menu()
        {
            Console.WriteLine("Qual LOG vc quer processar?");
            Console.WriteLine("1 - Enricher Product");
            Console.WriteLine("2 - Sync Product");
            Console.WriteLine("3 - Sync NationalPurchase");
            Console.WriteLine("4 - Sync IntangiblePurchase");
            Console.WriteLine("5 - sair");

            var nome = Console.ReadLine();

            return nome!;
        }
    }
}
