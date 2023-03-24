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
    
    public partial class ServiceReport
    {
        public int ID { get; set; }
        public int ImportID { get; set; }
        public string ImportIDString { get; set; }
        public Nullable<int> ServiceOrder { get; set; }
        public string ProblemCode { get; set; }
        public string SolutionCode { get; set; }
        public string ProblemDescription { get; set; }
        public string ActionUndertaken { get; set; }
        public bool CustomerInformed { get; set; }
        public bool ProblemSolved { get; set; }
        public string ProblemNotSolved { get; set; }
        public bool SafetyNOK { get; set; }
        public bool SafetyOK { get; set; }
        public string SafetyNOKDetails { get; set; }
        public bool Signed { get; set; }
        public Nullable<System.DateTime> DateSigned { get; set; }
        public string SignedBy { get; set; }
        public bool Requests { get; set; }
        public string RequestDetails { get; set; }
        public Nullable<System.DateTime> DateCreated { get; set; }
        public Nullable<System.DateTime> DateModified { get; set; }
        public string TechReport { get; set; }
        public Nullable<int> FSE { get; set; }
    }
}
