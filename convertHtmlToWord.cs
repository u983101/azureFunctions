using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;

namespace convertToWord
{
    public static class convertToWordFunc
    {
        [FunctionName("convertToWordFunc")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = null)] HttpRequest req, ILogger log)
        {
            try
            {
                log.LogInformation("ConvertToWord HTTP trigger function.");

                string strHtml = await new StreamReader(req.Body).ReadToEndAsync();
                MemoryStream htmlStream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(strHtml));
                MemoryStream outputStream = new MemoryStream();

                using (WordprocessingDocument doc =
                    WordprocessingDocument.Create(outputStream, WordprocessingDocumentType.Document))
                {
                    string altChunkId = "myId";
                    MainDocumentPart mainPart = doc.AddMainDocumentPart();

                    mainPart.Document = new Document();
                    Body body = mainPart.Document.AppendChild(new Body());

                    AlternativeFormatImportPart formatImportPart = mainPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Html, altChunkId);

                    formatImportPart.FeedData(htmlStream);

                    AltChunk altChunk = new AltChunk();
                    altChunk.Id = altChunkId;

                    mainPart.Document.Body.Append(altChunk);
                }
                log.LogInformation("Conversion Completed.");
                var finalStream = new MemoryStream(outputStream.ToArray());
                finalStream.Position = 0;
                return new OkObjectResult(finalStream);
            } catch(Exception e)
            {
                log.LogError(e, "Unable to convert");
                return new OkObjectResult(e.Message);
            }
        }
    }
}
