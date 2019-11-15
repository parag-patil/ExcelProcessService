using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using ClosedXML.Excel;
using Newtonsoft.Json;

namespace ExcelProcessingService.Models
{
    public class ProcessFile
    {
        private readonly HttpPostedData _postedData;
        public ProcessFile(HttpPostedData postedData)
        {
            _postedData = postedData;
        }

        public string Process()
        {
            var workBook = GetXLWorkbook(_postedData.File);
            var workSheet = workBook.Worksheets.FirstOrDefault();   //This service limits processing of first sheet. Can be improved in future.
            var firstDetailRowAddress = workSheet.Row(_postedData.DetailLineNumber).FirstCell().Address;
            var lastDetailRowAddress = workSheet.LastCellUsed().Address;
            var detailTableRange = workSheet.Range(firstDetailRowAddress, lastDetailRowAddress).RangeUsed();
            var detailTable = detailTableRange.AsTable().AsNativeDataTable();

            var _jsonSettings = new JsonSerializerSettings();
            _jsonSettings.DateFormatString = "MM/dd/yyyy hh:mm:ss tt";
            return JsonConvert.SerializeObject(detailTable, _jsonSettings);
        }

        private static XLWorkbook GetXLWorkbook(byte[] byteData)
        {
            Stream vStream = new MemoryStream();
            vStream.Write(byteData, 0, Convert.ToInt32(byteData.Length));
            XLWorkbook book = new XLWorkbook(vStream, XLEventTracking.Disabled);        //Disabling XLEventTracking will improve the processing time.
            return book;
        }
    }
}