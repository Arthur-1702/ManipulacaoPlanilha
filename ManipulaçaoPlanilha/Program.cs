using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace ManipulaçaoPlanilha
{
    internal class Program
    {
        static void Main(string[] args)
        {
            /*
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;            
            ExcelPackage arquivo = new ExcelPackage(new FileInfo("nomeArquivo")); //acessar arquivo
            ExcelWorkbook planilhas = arquivo.Workbook; //acessar o conjunto de planilhas
            var planilha = planilhas.Worksheets["nomePlanilha"]; //acessar planilha (pode usar o índice no lugar do nome)
            var linha = 1;
            var coluna = 1;
            planilha.Cells[linha,coluna].Value = "teste"; //acessar cada célula da planilha com variáveis como coordenadas
            planilha.Cells["A1"].Value = "teste"; // com o "endereço"
            planilha.Cells["A1:C1"].Style.Font.Bold = true; // aplicando em várias (nesse caso negrito na fonte)
            */

            string caminhoPlanilha = "C:\\Users\\arthur.araujo\\OneDrive - Kurier Tecnologia\\Documentos\\AtividadesVB\\ManipulaçaoPlanilha\\planilhas\\Vendas.xlsx";

            Console.WriteLine("Cria planilhas");
            Console.ReadLine();
            CriaPlanilha(caminhoPlanilha);

            Console.WriteLine("Abre planilha");
            Console.ReadLine();
            AbrePlanilha(caminhoPlanilha);
            Console.ReadLine();


        }

        private static void CriaPlanilha(string caminhoPlanilha)
        {
            var Vendas = new[]
            {
                new{ Id = "PE101", Filial="Recife", Vendas = 660},
                new{ Id = "PB101", Filial="João Pessoa", Vendas = 460},
                new{ Id = "RN101", Filial="Natal", Vendas = 560},
                new{ Id = "CE101", Filial="Fortaleza", Vendas = 650}
            };

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage arquivo = new ExcelPackage();

            var planilha = arquivo.Workbook.Worksheets.Add("PlanilhaVendas");

            planilha.TabColor = System.Drawing.Color.Black;
            planilha.DefaultRowHeight = 12;

            planilha.Row(1).Height = 20;
            planilha.Row(1).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            planilha.Row(1).Style.Font.Bold = true;
            planilha.Cells["A1:C1"].Style.Font.Italic = true;

            planilha.Cells[1, 1].Value = "Cod.";
            planilha.Cells[1, 2].Value = "Filial";
            planilha.Cells[1, 3].Value = "Vendas/mil";

            int indice = 2;
            foreach (var venda in Vendas)
            {
                planilha.Cells[indice, 1].Value = venda.Id;
                planilha.Cells[indice, 2].Value = venda.Filial;
                planilha.Cells[indice, 3].Value = venda.Vendas;
                indice ++;
            }

            //ajustar o tamanho das colunas para os valores
            planilha.Column(1).AutoFit();
            planilha.Column(2).AutoFit();
            planilha.Column(3).AutoFit();

            //apagar arquivo para substituir se existir
            if (File.Exists(caminhoPlanilha))
                File.Delete(caminhoPlanilha);

            //cria o arquivo de verdade
            FileStream objFileStrm = File.Create(caminhoPlanilha);
            objFileStrm.Close();

            //escreve do arquivo "virutal" para o de verdade
            File.WriteAllBytes(caminhoPlanilha,arquivo.GetAsByteArray()); //escrever na planilha

            arquivo.Dispose(); //fecha o arquivo virtual

            Console.WriteLine($"Planilha criado em: {caminhoPlanilha}");

        }
        
        private static void AbrePlanilha(string caminhoPlanilha)
        {
            var arquivo = new ExcelPackage(new FileInfo(caminhoPlanilha));
            ExcelWorksheet planilha = arquivo.Workbook.Worksheets.FirstOrDefault();

            //ler celula a celula para mostrar o conteudo
            for(int i = 1; i <= planilha.Dimension.Rows; i++)
            {
                for(int j= 1; j <= planilha.Dimension.Columns; j++)
                {
                    string conteudo = planilha.Cells[i, j].Value.ToString();
                    Console.WriteLine($"{conteudo}");
                }
            }
            

        }
    }
}
