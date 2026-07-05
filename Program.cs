using System;
using System.Collections.Generic;
using System.IO;

namespace NfeXmlToExcel
{
    internal class Program
    {
        static int Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;

            if (args.Length < 2)
            {
                Console.WriteLine("Uso:");
                Console.WriteLine("  NfeXmlToExcel <pasta_dos_xmls> <saida.xlsx>");
                Console.WriteLine();
                Console.WriteLine("Exemplo:");
                Console.WriteLine(@"  NfeXmlToExcel ""C:\XMLS_NFE"" ""C:\saida\nfes.xlsx""");
                return 1;
            }

            var inputFolder = args[0];
            var outputXlsx = args[1];

            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine("❌ Pasta não existe: " + inputFolder);
                return 1;
            }

            var xmlFiles = Directory.GetFiles(inputFolder, "*.xml", SearchOption.AllDirectories);
            Console.WriteLine($"📂 Encontrados {xmlFiles.Length} XML(s).");

            var headers = new List<NfeHeader>();
            var items = new List<NfeItem>();
            var errors = new List<ParseError>();

            int ok = 0, fail = 0;

            foreach (var file in xmlFiles)
            {
                try
                {
                    var doc = NfeParser.ParseFromFile(file);
                    headers.Add(doc.Header);
                    items.AddRange(doc.Items);
                    ok++;
                }
                catch (Exception ex)
                {
                    fail++;
                    errors.Add(new ParseError
                    {
                        FilePath = file,
                        Message = ex.Message
                    });
                }
            }

            Console.WriteLine($"✅ OK: {ok} | ❌ Erros: {fail}");
            Console.WriteLine("🧾 Gerando Excel...");

            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(outputXlsx))!);

            ExcelExporter.Export(outputXlsx, headers, items, errors);

            Console.WriteLine("✅ Excel gerado em: " + Path.GetFullPath(outputXlsx));
            return 0;
        }
    }
}
