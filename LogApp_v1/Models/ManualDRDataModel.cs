using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LogApp_v1.Models
{
    public class ManualDRDataModel
    {
        public bool ischecked { get; set; }
        public string Facility { get; set; }
        public string Partnumber { get; set; }
        public string Pull_Qty { get; set; }
        public string Pull_Ticker_Number { get; set;}
        public string Line { get; set;}
        public string Remarks { get; set;}
        public string Date_Added { get; set; }

        public ManualDRDataModel(){ }
        public ManualDRDataModel( bool ischecked, string Facility, string Partnumber, string Pull_Qty, string Pull_Ticket_Number, string Line, string Remarks, String Date_Added)
        {
            this.ischecked = ischecked;
            this.Facility = Facility;
            this.Partnumber = Partnumber;
            this.Pull_Qty = Pull_Qty;
            this.Pull_Ticker_Number = Pull_Ticket_Number;
            this.Line = Line;
            this.Remarks = Remarks;
            this.Date_Added = Date_Added;
        }
    }
}
