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
    
    public partial class CostDetails
    {
        public int ID { get; set; }
        public int ImportID { get; set; }
        public string ImportIDString { get; set; }
        public int CostCategory { get; set; }
        public string Detail { get; set; }
        public Nullable<System.DateTime> DateFrom { get; set; }
        public Nullable<System.DateTime> DateTo { get; set; }
        public Nullable<decimal> Amount { get; set; }
        public Nullable<int> Currency { get; set; }
        public Nullable<int> PaymentMethod { get; set; }
        public Nullable<int> TicketNumber { get; set; }
        public Nullable<int> EnvelopeNumber { get; set; }
        public bool AllAssigned { get; set; }
        public Nullable<System.DateTime> DateCreated { get; set; }
        public Nullable<System.DateTime> DateModified { get; set; }
        public Nullable<int> FSE { get; set; }
        public string Department { get; set; }
    }
}
