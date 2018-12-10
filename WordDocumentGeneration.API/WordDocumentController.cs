using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace WordDocumentGeneration.API
{
    public class WordDocumentController : ApiController
    {
        [HttpGet]
        [Route("api/worddocument")]
        public HttpResponseMessage DownloadPdfFile()
        {
            try
            {
                var documentManager = new WordDocumentManager();

                var bytes = documentManager.GetDocument(new GenerationData());

                var result = Request.CreateResponse(HttpStatusCode.OK);
                result.Content = new ByteArrayContent(bytes);
                result.Content.Headers.ContentDisposition =
                    new System.Net.Http.Headers.ContentDispositionHeaderValue(
                            "attachment")
                        { FileName = "Test" + ".docx" };

                return result;
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.Gone);
            }
        }
    }
}
