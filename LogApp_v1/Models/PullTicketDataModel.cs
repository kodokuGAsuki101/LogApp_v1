using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LogApp_v1.Models
{
    public class PullTicketDataModel
    {
        public PullTicketDataModel()
        {
        }

        public string prodDate{ get; set; }
        public string prodTime{ get; set; }
        public string delDate{ get; set; }
        public string delTime{ get; set; }
        public string jobNo{ get; set; }
        public string facility{ get; set; }
        public string partNo{ get; set; }
        public string pullQty{ get; set; }
        public string vendorName{ get; set; }
        public string skuAssembly{ get; set; }
        public string cellNo{ get; set; }
        public string remarks{ get; set; }
        public string pullTicketNo{ get; set; }
        public string line{ get; set; }
        public string fileUploadDate{ get; set; }
        public string vendorAcknowledgement{ get; set; }
        public string acknowledgementDate{ get; set; }
        public string acknowledgementRemarksForVendor{ get; set; }
        public string qtyDelivered{ get; set; }
        public string dlVarience{ get; set; }
        public string hitmiss{ get; set; }
        public string status{ get; set; }

        public PullTicketDataModel(
            string prodDate,
            string prodTime,
            string delDate,
            string delTime,
            string jobNo,
            string facility,
            string partNo,
            string pullQty,
            string vendorName,
            string skuAssembly,
            string cellNo,
            string remarks,
            string pullTicketNo,
            string line,
            string fileUploadDate,
            string vendorAcknowledgement,
            string acknowledgementDate,
            string acknowledgementRemarksForVendor,
            string qtyDelivered,
            string dlVarience,
            string hitmiss,
            string status
            )
        {
            this.prodDate = prodDate;
            this.prodTime = prodTime;
            this.delDate = delDate;
            this.jobNo = jobNo;
            this.facility = facility;
            this.partNo = partNo;
            this.pullQty = pullQty;
            this.vendorName = vendorName;
            this.skuAssembly = skuAssembly;
            this.cellNo = cellNo;
            this.remarks = remarks;
            this.pullTicketNo = pullTicketNo;
            this.line = line;
            this.fileUploadDate = fileUploadDate;
            this.vendorAcknowledgement = vendorAcknowledgement;
            this.acknowledgementDate = acknowledgementDate;
            this.acknowledgementRemarksForVendor = acknowledgementRemarksForVendor;
            this.qtyDelivered = qtyDelivered;
            this.dlVarience = dlVarience;
            this.hitmiss = hitmiss;
            this.status = status;
            this.delTime = delTime;
        }
    }
}
