//------------------------------------------------------------------------------
// <auto-generated>
//    This code was generated from a template.
//
//    Manual changes to this file may cause unexpected behavior in your application.
//    Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace Alarms
{
    using System;
    using System.Collections.Generic;
    
    public partial class ServiceorderCalculateDetails
    {
        public int ID { get; set; }
        public int Serviceorder { get; set; }
        public int Invoicing { get; set; }
        public int WhoToInvoice { get; set; }
        public int HourGroup { get; set; }
        public int TaskGroup { get; set; }
        public int Type { get; set; }
        public int Traveling { get; set; }
        public Nullable<decimal> CalloutValue { get; set; }
        public string CallOutText { get; set; }
        public Nullable<decimal> KmCharge { get; set; }
        public string OtherText { get; set; }
        public bool DoNotChangeAnymore { get; set; }
        public Nullable<int> NumberOfCallouts { get; set; }
        public int numberOfKm { get; set; }
        public decimal TravelTime { get; set; }
    }
}