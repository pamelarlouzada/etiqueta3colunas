using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using ClosedXML.Excel;
using PdfSharpCore.Drawing;
using PdfSharpCore.Pdf;
using ZXing;
using ZXing.Common;

class Program
{
    static void Main(string[] args)
    {
        // Caminho do arquivo Excel a ser lido
        string excelFilePath = "C:/Users/pamela.louzada/Desktop/Meus/_Etiquetas/EtiquetasRetang/EtiquetasRetang/planilha/file.xlsx";
        // Caminho onde o arquivo PDF será salvo
        string pdfOutputPath = "C:/Users/pamela.louzada/Desktop/Meus/_Etiquetas/EtiquetasRetang/EtiquetasRetang/etiqueta/output.pdf";

        // Ler os itens do arquivo Excel
        List<Item> items = ReadExcelFile(excelFilePath);
        // Gerar o PDF com base nos itens lidos
        GeneratePdf(items, pdfOutputPath);

        // Mensagem de confirmação no console
        Console.WriteLine("Etiquetas geradas com sucesso.");
    }

    static List<Item> ReadExcelFile(string filePath)
    {
        var items = new List<Item>(); // Lista para armazenar os itens lidos do Excel
        using (var workbook = new XLWorkbook(filePath)) // Abrir o arquivo Excel
        {
            var worksheet = workbook.Worksheet(1); // Selecionar a primeira planilha
            foreach (var row in worksheet.RangeUsed().RowsUsed().Skip(1)) // Iterar pelas linhas usadas, pulando a primeira linha (cabeçalho)
            {
                var item = new Item
                {
                    Code = row.Cell(1).GetString(), // Ler o código da peça
                    Description = row.Cell(2).GetString(), // Ler a descrição da peça
                    Price = row.Cell(3).GetString(), // Ler o preço da peça
                    Quantity = row.Cell(4).GetValue<int>() // Ler a quantidade de peças
                };
                items.Add(item); // Adicionar o item à lista
                // Mensagem de depuração
                Console.WriteLine($"Lido do Excel: {item.Code}, {item.Description}, {item.Price}, {item.Quantity}");
            }
        }
        return items; // Retornar a lista de itens lidos
    }

    static void GeneratePdf(List<Item> items, string outputPath)
    {
        using (var document = new PdfDocument()) // Criar um novo documento PDF
        {
            var page = document.AddPage(); // Adicionar uma nova página ao documento
            page.Width = XUnit.FromCentimeter(10.6); // Definir a largura da página
            page.Height = XUnit.FromCentimeter(2.1); // Definir a altura da página
            var gfx = XGraphics.FromPdfPage(page); // Criar um objeto gráfico para desenhar na página
            var font = new XFont("Arial", 7, XFontStyle.Regular); // Definir a fonte a ser usada

            int column = 0; // Contador de colunas
            double columnWidth = page.Width / 3; // Largura de cada coluna
            double barcodeHeight = page.Height * 0.5; // Altura da linha do código de barras
            double textLineHeight = (page.Height - barcodeHeight) / 3.5; // Altura das outras linhas ajustada para diminuir espaçamento
            double padding = 1; // Padding ajustado para diminuir espaçamento

            foreach (var item in items) // Iterar pelos itens
            {
                for (int i = 0; i < item.Quantity; i++) // Iterar pela quantidade de cada item
                {
                    if (column == 3) // Se a coluna atual for a quarta (índice 3), adicionar uma nova página
                    {
                        column = 0;
                        page = document.AddPage();
                        page.Width = XUnit.FromCentimeter(10.6);
                        page.Height = XUnit.FromCentimeter(2.1);
                        gfx = XGraphics.FromPdfPage(page);
                    }

                    double x = column * columnWidth; // Calcular a posição x da coluna atual
                    double y = 0; // Posição y inicial

                    // Linha 1: Imagem do código de barras
                    var barcodeImage = GenerateBarcode(item.Code); // Gerar a imagem do código de barras
                    var barcodeStream = new MemoryStream(); // Criar um stream de memória
                    barcodeImage.Save(barcodeStream, ImageFormat.Png); // Salvar a imagem no stream
                    barcodeStream.Position = 0; // Reiniciar a posição do stream
                    XImage xImage = XImage.FromStream(() => new MemoryStream(barcodeStream.ToArray())); // Criar uma imagem XImage a partir do stream
                    gfx.DrawImage(xImage, x + padding, y + padding, columnWidth - 2 * padding, barcodeHeight - 2 * padding); // Desenhar a imagem na página

                    // Verificar o tamanho da descrição e ajustar as linhas
                    string descriptionLine1 = item.Description.Length > 20 ? item.Description.Substring(0, 20) : item.Description;
                    string descriptionLine2 = item.Description.Length > 20 ? item.Description.Substring(20, Math.Min(20, item.Description.Length - 20)) : "";

                    // Se a descrição tiver até 20 caracteres, utilizar três linhas
                    if (string.IsNullOrEmpty(descriptionLine2))
                    {
                        // Linha 2: Descrição
                        y += barcodeHeight; // Mover para a próxima linha
                        gfx.DrawString(descriptionLine1, font, XBrushes.Black, new XRect(x + padding, y, columnWidth - 2 * padding, textLineHeight), XStringFormats.Center); // Desenhar a descrição

                        // Linha 3: Código da peça e valor
                        y += textLineHeight; // Mover para a próxima linha
                        gfx.DrawString($"783/{item.Code}/R${item.Price}", font, XBrushes.Black, new XRect(x + padding, y, columnWidth - 2 * padding, textLineHeight), XStringFormats.Center); // Desenhar o código e o valor
                    }
                    else
                    {
                        // Se a descrição tiver mais de 20 caracteres, utilizar quatro linhas
                        // Linha 2: Início da descrição
                        y += barcodeHeight; // Mover para a próxima linha
                        gfx.DrawString(descriptionLine1, font, XBrushes.Black, new XRect(x + padding, y, columnWidth - 2 * padding, textLineHeight), XStringFormats.Center); // Desenhar o início da descrição

                        // Linha 3: Continuação da descrição
                        y += textLineHeight; // Mover para a próxima linha
                        gfx.DrawString(descriptionLine2, font, XBrushes.Black, new XRect(x + padding, y, columnWidth - 2 * padding, textLineHeight), XStringFormats.Center); // Desenhar a continuação da descrição

                        // Linha 4: Código da peça e valor
                        y += textLineHeight; // Mover para a próxima linha
                        gfx.DrawString($"{item.Code}/{item.Price}", font, XBrushes.Black, new XRect(x + padding, y, columnWidth - 2 * padding, textLineHeight), XStringFormats.Center); // Desenhar o código e o valor
                    }

                    column++; // Mover para a próxima coluna
                }
            }

            document.Save(outputPath); // Salvar o documento PDF no caminho especificado
        }
    }

    static Bitmap GenerateBarcode(string code)
    {
        var writer = new BarcodeWriterPixelData // Criar um gerador de código de barras
        {
            Format = BarcodeFormat.CODE_128, // Definir o formato do código de barras
            Options = new EncodingOptions // Definir as opções de codificação
            {
                Width = 100, // Largura do código de barras
                Height = 40 // Altura do código de barras
            }
        };

        var pixelData = writer.Write(code); // Gerar os dados de pixel do código de barras
        var bitmap = new Bitmap(pixelData.Width, pixelData.Height, PixelFormat.Format32bppRgb); // Criar um bitmap a partir dos dados de pixel
        var bitmapData = bitmap.LockBits(new Rectangle(0, 0, pixelData.Width, pixelData.Height),
                                        ImageLockMode.WriteOnly, PixelFormat.Format32bppRgb); // Bloquear os bits do bitmap para escrita
        try
        {
            System.Runtime.InteropServices.Marshal.Copy(pixelData.Pixels, 0, bitmapData.Scan0, pixelData.Pixels.Length); // Copiar os dados de pixel para o bitmap
        }
        finally
        {
            bitmap.UnlockBits(bitmapData); // Desbloquear os bits do bitmap
        }

        return bitmap; // Retornar o bitmap gerado
    }

    // Classe para representar um item
    class Item
    {
        public string Code { get; set; } // Código da peça
        public string Description { get; set; } // Descrição da peça
        public string Price { get; set; } // Preço da peça
        public int Quantity { get; set; } // Quantidade de peças
    }
}