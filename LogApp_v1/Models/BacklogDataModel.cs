using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LogApp_v1.Models
{
    public class BacklogDataModel
    {

        public bool isChecked {  get; set; }
        public string Del_Date {  get; set; }
        public string Del_Time {  get; set; }
        public string Facility {  get; set; }
        public string Partnumber {  get; set; }
        public string Balance {  get; set; }
        public string Remarks {  get; set; }
        public string PullTicketNumber {  get; set; }
        public string Line {  get; set; }
        public string Qty_Del {  get; set; }
        public string Original_Pull {  get; set; }
        public string BackLogType {  get; set; }
        public string History {  get; set; }


        public BacklogDataModel() { }

        public BacklogDataModel(
            bool isChecked,
            string Del_Date, 
            string Del_Time, 
            string Facility, 
            string Partnumber, 
            string Balance, 
            string Remarks, 
            string PullTicketNumber,
            string Line,
            string Qty_Del, 
            string Original_Pull,
            string BackLogType,
            string History
            ) {
            this.isChecked = isChecked;
            this.Del_Date = Del_Date;
            this.Facility = Facility;
            this.Partnumber = Partnumber;
            this.Balance = Balance;
            this.Remarks = Remarks;
            this.PullTicketNumber = PullTicketNumber;
            this.Line = Line;  
            this.Qty_Del = Qty_Del;
            this.Original_Pull = Original_Pull;
            this.BackLogType = BackLogType;
            this.History = History;
           
        }

    }
}
