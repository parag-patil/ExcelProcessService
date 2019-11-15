using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExcelProcessingService.Models
{
    public class HttpPostedData
    {
        public byte[] File { get; set; }
        //Used for future provision to handle master detail excel file
        private int _headerLineNumber = 0;
        public int HeaderLineNumber
        {
            get
            {
                return this._headerLineNumber;
            }
            set
            {
                this._headerLineNumber = value;
            }
        }


        private int _detailLineNumber = 1;
        public int DetailLineNumber
        {
            get
            {
                return this._detailLineNumber;
            }
            set
            {
                this._detailLineNumber = value;
            }
        }
    }
}