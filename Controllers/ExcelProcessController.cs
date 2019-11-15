using ExcelProcessingService.Extensions;
using ExcelProcessingService.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;

namespace ExcelProcessingService.Controllers
{
    [RoutePrefix("api")]
    public class ExcelProcessController : ApiController
    {
       string UNSUPPORTED_CORRUPTED_DATA = "Unsupported File or File contains corrupted data.";
       public async Task<HttpResponseMessage> ProcessExcel(HttpRequestMessage request)
        {
            string _jsonResult = string.Empty;
            HttpResponseMessage response = new HttpResponseMessage();
            if (!request.Content.IsMimeMultipartContent())
            {
                response.Content = new StringContent(UNSUPPORTED_CORRUPTED_DATA);
                throw new HttpResponseException(HttpStatusCode.UnsupportedMediaType);
            }

            try
            {
                HttpPostedData postedData = await Request.Content.ParseMultipartAsync();
                ProcessFile postedFile = new ProcessFile(postedData);
                _jsonResult = postedFile.Process();
                response.StatusCode = HttpStatusCode.OK;
                response.Content = new StringContent(_jsonResult);
            }
            catch(Exception ex)
            {
                //Lof of scope to improve error handling. Will improve in future when refactor this service for .NET Core.
                if (ex.Message.Contains(UNSUPPORTED_CORRUPTED_DATA))
                {
                    response.StatusCode = HttpStatusCode.UnsupportedMediaType;
                }
                else
                {
                    response.StatusCode = HttpStatusCode.BadRequest;
                }

                response.Content = new StringContent(ex.Message.ToString());
            }

            return response;
        }
    }
}
