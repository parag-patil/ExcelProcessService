using ExcelProcessingService.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web;

namespace ExcelProcessingService.Extensions
{
    public static class HttpContentExtension
    {
        public static async Task<HttpPostedData> ParseMultipartAsync(this HttpContent postedContent)
        {
            HttpPostedData postedData = new HttpPostedData();
            try
            {
                var provider = await postedContent.ReadAsMultipartAsync();
                foreach(var content in provider.Contents)
                {   //If no file is uploaded from client, this property will be null
                    if (!string.IsNullOrWhiteSpace(content.Headers.ContentDisposition.FileName))
                    {
                        postedData.File = await content.ReadAsByteArrayAsync();
                    }
                    else
                    {
                        var attributeName = content.Headers.ContentDisposition.Name.Trim('"');
                        var value = await content.ReadAsStringAsync();
                        _setAttribute(attributeName, value, postedData);
                    }
                }

                return postedData;
            }
            catch(Exception ex)
            {
                throw ex;
            }
        }

        private static void _setAttribute(string attributeName,string value,HttpPostedData postedDataObject)
        {
            switch (attributeName)
            {
                case "HeaderLineNumber":
                    int headerLineNumber=1; //set default to 1.
                    int.TryParse(value, out headerLineNumber);
                    postedDataObject.HeaderLineNumber = headerLineNumber;
                    break;
                case "DetailLineNumber":
                    int detailLineNumber = 1; //set default to 1.
                    int.TryParse(value, out detailLineNumber);
                    postedDataObject.DetailLineNumber = detailLineNumber;
                    break;
            }
        }
    }
}