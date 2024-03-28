using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace FileHandler.Models
{
    public class Customer
    {
        public int Id { get; set; }
        public string CustomerName { get; set; }
        public string CustomerCode { get; set; }
        public string Address1 { get; set; }
        public string Address2 { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string Pin { get; set; }
        public string MobileNo { get; set; }
    }
}