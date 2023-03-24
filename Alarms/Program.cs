using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using PdfSharp.Pdf;


//TO DO:
//Add critical calls : TO DO
//Special procedure for long IBS: TO DO

namespace Alarms
{
    class Program
    {
        static void Main(string[] args)
        {

            //Load all serviceorders
            using (ERPLNDataEntities myERPLNData = new ERPLNDataEntities())
            {
                using (InterventionsEntities myInterventionsData = new InterventionsEntities())
                {

                    #region Update Strippit codes
                    //var StrippitInterventions = (from s in myInterventionsData.ServiceOrder
                    //                             where s.SrvOrder >= 110000000 && s.SrvOrder <= 119999999
                    //                             select s).ToList();
                    //foreach (var item in StrippitInterventions)
                    //{
                    //    string CoverageType = "131";
                    //    string CoverageDescription = "Service - machine out of warra";
                    //    string srvtype = "1";
                    //    if (item.Omschr8 != null)
                    //    {
                    //        switch (item.Omschr8.Substring(0, 3).ToLower())
                    //        {
                    //            case "nor":
                    //                {
                    //                    srvtype = "131";
                    //                    CoverageType = "1";
                    //                    CoverageDescription = "Service Chargeable";
                    //                    break;
                    //                }
                    //            case "cty":
                    //                {
                    //                    srvtype = "507";
                    //                    CoverageType = "5";
                    //                    CoverageDescription = "Service Not Chargeable";
                    //                    break;
                    //                }
                    //            case "evp":
                    //                {
                    //                    srvtype = "136";
                    //                    CoverageType = "3";
                    //                    CoverageDescription = "Contract";
                    //                    break;
                    //                }
                    //            case "int":
                    //                {
                    //                    srvtype = "205";
                    //                    CoverageType = "5";
                    //                    CoverageDescription = "Service Not Chargeable";
                    //                    break;
                    //                }
                    //            case "pd":
                    //                {
                    //                    srvtype = "131";
                    //                    CoverageType = "1";
                    //                    CoverageDescription = "Service Chargeable";
                    //                    break;
                    //                }
                    //            case "pm":
                    //                {
                    //                    srvtype = "156";
                    //                    CoverageType = "1";
                    //                    CoverageDescription = "Service Chargeable";
                    //                    break;
                    //                }
                    //            case "rep":
                    //                {
                    //                    srvtype = "131";
                    //                    CoverageType = "1";
                    //                    CoverageDescription = "Service Chargeable";
                    //                    break;
                    //                }
                    //            case "rtf":
                    //                {
                    //                    srvtype = "186";
                    //                    CoverageType = "1";
                    //                    CoverageDescription = "Service Chargeable";
                    //                    break;
                    //                }
                    //            case "rwa":
                    //                {
                    //                    srvtype = "305";
                    //                    CoverageType = "5";
                    //                    CoverageDescription = "Service Not Chargeable";
                    //                    break;
                    //                }
                    //            case "war":
                    //                {
                    //                    srvtype = "304";
                    //                    CoverageType = "4";
                    //                    CoverageDescription = "Warranty";
                    //                    break;
                    //                }
                    //            default:
                    //                {
                    //                    srvtype = "131";
                    //                    CoverageType = "1";
                    //                    CoverageDescription = "Service Chargeable";
                    //                    break;
                    //                }
                    //        }
                    //        item.SrvType = Convert.ToInt32(srvtype);
                    //        item.CoverageType = Convert.ToInt32(CoverageType);
                    //        item.Omschr7 = CoverageDescription;
                    //    }
                    //}
                    //myInterventionsData.SaveChanges();

                    //var StrippitActivities = (from s in myInterventionsData.AlarmActivities
                    //                          where s.ActivityType == "Intervention" && s.Department == "STRIPPIT"
                    //                          select s).ToList();
                    //foreach (var item in StrippitActivities)
                    //{
                    //    var findCode = (from s in myInterventionsData.ServiceOrder
                    //                    where s.SrvOrder == item.ActivityID
                    //                    select new
                    //                    {
                    //                        activityCode = s.SrvType
                    //                    }).FirstOrDefault();
                    //    if (findCode != null)
                    //    {
                    //        item.ActivityCode = findCode.activityCode.ToString();
                    //    }
                    //    myInterventionsData.SaveChanges();
                    //}

                    #endregion

                    #region Update Technologies
                    //var alarmToUpdate = (from s in myInterventionsData.Alarms
                    //                    where s.Technology == null || s.Technology == ""
                    //                    select s).ToList();
                    //if (alarmToUpdate!=null)
                    //{
                    //    foreach (var item in alarmToUpdate)
                    //    {
                    //        var thisMachine = (from s in myInterventionsData.Machines
                    //                           where s.SerialNumber == item.SerialNumber
                    //                           select s).FirstOrDefault();
                    //        if (thisMachine!=null)
                    //        {
                    //            item.Technology = thisMachine.Technology;
                    //        }
                    //    }
                    //    myInterventionsData.SaveChanges();
                    //}
                    //var alarmHistoryToUpdate = (from s in myInterventionsData.AlarmHistory
                    //                     where s.Technology == null || s.Technology == ""
                    //                     select s).ToList();
                    //if (alarmHistoryToUpdate != null)
                    //{
                    //    foreach (var item in alarmHistoryToUpdate)
                    //    {
                    //        var thisMachine = (from s in myInterventionsData.Machines
                    //                           where s.SerialNumber == item.SerialNumber
                    //                           select s).FirstOrDefault();
                    //        if (thisMachine != null)
                    //        {
                    //            item.Technology = thisMachine.Technology;
                    //        }
                    //    }
                    //    myInterventionsData.SaveChanges();
                    //}

                    #endregion

                    #region declarations

                    List<string> warrantyCodes = new List<string>();
                    List<string> ibsCodes = new List<string>();
                    List<string> codesToBeExcluded = new List<string>();

                    var codes = from s in myInterventionsData.TypeCode
                                select s;
                    if (codes != null)
                    {
                        foreach (var item in codes)
                        {
                            if (item.Warranty && item.Warning)
                            {
                                if (!warrantyCodes.Contains(item.TypeInt.ToString()))
                                {
                                    warrantyCodes.Add(item.TypeInt.ToString());
                                }
                            }
                            if (item.IBS && item.Warning)
                            {
                                if (!ibsCodes.Contains(item.TypeInt.ToString()))
                                {
                                    ibsCodes.Add(item.TypeInt.ToString());
                                }
                            }
                            if (item.Warning == false)
                            {
                                if (!codesToBeExcluded.Contains(item.TypeInt.ToString()))
                                {
                                    codesToBeExcluded.Add(item.TypeInt.ToString());
                                }
                            }
                        }
                    }


                    bool[] combinedLevel = { false, false, false };
                    int[] numberOfPartsordersLevel = { 999, 999, 999 };
                    int[] numberOfInterventionsLevel = { 999, 999, 999 };
                    int[] numberOfCallsLevel = { 999, 999, 999 };
                    int[] levelMonths = { 6, 2, 1 };
                    bool[] includeIBS = { false, false, false };
                    bool[] allInterventions = { true, true, true };
                    bool[] allParts = { false, false, false };

                    var alarmLevels = (from s in myInterventionsData.AlarmLevels
                                       where s.Warranty == true
                                       select s).ToList();
                    if (alarmLevels != null)
                    {
                        foreach (var alarmLevel in alarmLevels)
                        {
                            if (alarmLevel.AlarmLevel > 0 && alarmLevel.AlarmLevel <= 3)
                            {
                                numberOfPartsordersLevel[alarmLevel.AlarmLevel - 1] = alarmLevel.PartsOrders;
                                numberOfInterventionsLevel[alarmLevel.AlarmLevel - 1] = alarmLevel.Interventions;
                                numberOfCallsLevel[alarmLevel.AlarmLevel - 1] = alarmLevel.Calls;
                                combinedLevel[alarmLevel.AlarmLevel - 1] = alarmLevel.AllConditionsShouldBeMet;
                                levelMonths[alarmLevel.AlarmLevel - 1] = alarmLevel.MonthsToEvaluate;
                                includeIBS[alarmLevel.AlarmLevel - 1] = alarmLevel.IncludeIBS;
                                allInterventions[alarmLevel.AlarmLevel - 1] = alarmLevel.AllINterventions;
                                allParts[alarmLevel.AlarmLevel - 1] = alarmLevel.AllParts;
                            }
                        }
                    }
                    #endregion

                    #region Import new parts orders
                    int lastPartsOrder = 700000000;
                    //Check for the Oldest ServiceOrder Imported in the Alarmactivities
                    var OldestOrder = (from s in myInterventionsData.AlarmActivities
                                       where s.ActivityType == "Partsorder" &&
                                             s.ActivityID <= 799999999 &&
                                             s.ActivityID >= 700000000
                                       orderby s.ActivityID descending
                                       select s).FirstOrDefault();
                    if (OldestOrder != null)
                    {
                        lastPartsOrder = OldestOrder.ActivityID;
                    }
                    //Get all the new orders
                    var toAddOrders = (from s in myERPLNData.Srv_ServiceOrders
                                       join t in myERPLNData.Srv_Calls on s.Serviceorder equals t.ServiceOrder
                                       where s.Serviceorder <= 799999999 &&
                                             s.Serviceorder > lastPartsOrder &&
                                             s.Status != "Geannuleerd" &&
                                             t.SerianNumberMachine != "" &&
                                             (s.Type == "SL1" || s.Type == "SL3" || s.Type == "SL4" || s.Type == "SL5")
                                       orderby s.Serviceorder
                                       select new
                                       {
                                           Type = s.Type,
                                           Serviceorder = s.Serviceorder,
                                           Country = s.Country,
                                           CustomerName = s.CustomerName,
                                           OrderDate = s.orderDate,
                                           Department = s.Department,
                                           MachineType = t.DetailMachine,
                                           SerialNumber = t.SerianNumberMachine

                                       }).ToList();
                    if (toAddOrders != null)
                    {
                        foreach (var item in toAddOrders)
                        {
                            AlarmActivities newActivity = new AlarmActivities();
                            newActivity.ActivityCode = item.Type;
                            newActivity.ActivityID = item.Serviceorder;
                            newActivity.ActivityType = "Partsorder";
                            newActivity.Country = item.Country;
                            newActivity.CustomerName = item.CustomerName;
                            newActivity.DateCreated = (DateTime)item.OrderDate;
                            newActivity.DateModified = (DateTime)item.OrderDate;
                            newActivity.Department = item.Department;
                            newActivity.EndDate = (DateTime)item.OrderDate;
                            newActivity.MachineType = item.MachineType;
                            newActivity.SerialNumber = item.SerialNumber;
                            newActivity.StartDate = (DateTime)item.OrderDate;
                            if (item.Type == "SL1")
                            {
                                newActivity.Warranty = false;
                            }
                            else
                            {
                                newActivity.Warranty = true;
                            }
                            myInterventionsData.AlarmActivities.Add(newActivity);
                        }
                        myInterventionsData.SaveChanges();
                    }
                    #endregion

                    #region Import new Interventions Strippit
                    //This is done in exportserviceclient
                    //for (int i = 0; i < 10; i++)
                    //{
                    //    int lastIntervention = Convert.ToInt32("11"+i.ToString()+"000000");
                    //    int maxIntervention = Convert.ToInt32("11" + i.ToString() + "999000");
                    //    //Check for the Oldest ServiceOrder Imported in the Alarmactivities
                    //    var OldestInterventionStrippit = (from s in myInterventionsData.AlarmActivities
                    //                                      where s.ActivityType == "Intervention" &&
                    //                                            s.ActivityID <= maxIntervention &&
                    //                                            s.ActivityID >= lastIntervention
                    //                                      orderby s.ActivityID descending
                    //                                      select s).FirstOrDefault();
                    //    if (OldestInterventionStrippit != null)
                    //    {
                    //        lastIntervention = OldestInterventionStrippit.ActivityID;
                    //    }
                    //    //Get all the new orders
                    //    var toAddInterventions = (from s in myInterventionsData.ServiceOrder
                    //                              where s.SrvOrder <= maxIntervention &&
                    //                                    s.SrvOrder > lastIntervention &&
                    //                                    s.Verwerking != "Cancelled" &&
                    //                                    s.SerialNr != "" 
                    //                              orderby s.SrvOrder
                    //                              select new
                    //                              {
                    //                                  Type = s.SrvType,
                    //                                  Serviceorder = s.SrvOrder,
                    //                                  Country = s.Omschr10,
                    //                                  CustomerName = s.Omschr20,
                    //                                  OrderDate = s.CreatedDate,
                    //                                  Department = s.SrvCenter,
                    //                                  MachineType = s.ProdDesc,
                    //                                  SerialNumber = s.SerialNr,
                    //                                  warranty = s.Omschr7

                    //                              }).ToList();
                    //    if (toAddInterventions != null)
                    //    {
                    //        foreach (var item in toAddInterventions)
                    //        {
                    //            AlarmActivities newActivity = new AlarmActivities();
                    //            newActivity.ActivityCode = item.Type.ToString();
                    //            newActivity.ActivityID = (int)item.Serviceorder;
                    //            newActivity.ActivityType = "Intervention";
                    //            newActivity.Country = item.Country;
                    //            newActivity.CustomerName = item.CustomerName;
                    //            newActivity.DateCreated = (DateTime)item.OrderDate;
                    //            newActivity.DateModified = (DateTime)item.OrderDate;
                    //            newActivity.Department = item.Department;
                    //            newActivity.EndDate = (DateTime)item.OrderDate;
                    //            newActivity.MachineType = item.MachineType;
                    //            newActivity.SerialNumber = item.SerialNumber;
                    //            newActivity.StartDate = (DateTime)item.OrderDate;
                    //            if (item.warranty.ToUpper() == "Warranty")
                    //            {
                    //                newActivity.Warranty = true;
                    //            }
                    //            else
                    //            {
                    //                newActivity.Warranty = false;
                    //            }
                    //            myInterventionsData.AlarmActivities.Add(newActivity);
                    //        }
                    //        myInterventionsData.SaveChanges();
                    //    }
                    //}
                    #endregion

                    #region Reset existing alarms if needed
                    //For machines with alarm level 1, 2 or 3 check the status for level 1.
                    //If this isn't valid anymore then reset alarm.

                    DateTime periodToEvaluate = DateTime.Now.AddMonths(-levelMonths[0]);

                    var machinesWithAlarms = (from s in myInterventionsData.Alarms
                                              where s.AlarmLevel > 0
                                              orderby s.SerialNumber
                                              select s).ToList();
                    if (machinesWithAlarms != null)
                    {
                        foreach (var item in machinesWithAlarms)
                        {

                            int numberOfInterventions = 0;
                            int numberOfCalls = 0;
                            int numberOfPartsOrders = 0;

                            //If there is no alarm active, then check if a new alarm hast to be generated.
                            //Get a list of activities 

                            for (int i = item.AlarmLevel - 1; i >= 0; i--)
                            {

                                DateTime startDateToInvestigate = DateTime.Now.AddMonths((levelMonths[i]) * (-1));

                                bool _includeIBS = includeIBS[i];
                                bool _allInterventions = allInterventions[i];
                                var InterventionsForThisMachine = (from s in myInterventionsData.AlarmActivities
                                                                   where s.SerialNumber == item.SerialNumber &&
                                                                   s.StartDate >= startDateToInvestigate &&
                                                                   s.ActivityType == "Intervention" &&
                                                                       (!(ibsCodes.Contains(s.ActivityCode)) || _includeIBS) &&
                                                                       (warrantyCodes.Contains(s.ActivityCode) || _allInterventions) &&
                                                                        (!(codesToBeExcluded.Contains(s.ActivityCode)))
                                                                   orderby s.StartDate
                                                                   select s).ToList();

                                bool keepAlarm = false;
                                if (InterventionsForThisMachine != null && InterventionsForThisMachine.Count() > 0)
                                {
                                    if (InterventionsForThisMachine.Count() >= numberOfInterventionsLevel[i])
                                    {
                                        //Check and see if all are valid
                                        //Consecutive interventions and interventions at the same time as one
                                        int interventionCounter = CalculateInterventions(InterventionsForThisMachine);
                                        if (interventionCounter >= numberOfInterventionsLevel[i])
                                        {
                                            numberOfInterventions = interventionCounter;
                                        }
                                        else
                                        {
                                            numberOfInterventions = 0;
                                        }

                                    }
                                    else
                                    {
                                        numberOfInterventions = 0;
                                    }

                                }

                                //Only warrantyParts
                                var partsordersForThisMachine = (from s in myInterventionsData.AlarmActivities
                                                                 where s.SerialNumber == item.SerialNumber &&
                                                                 s.StartDate >= startDateToInvestigate &&
                                                                     //  s.StartDate <= DateTime.Now &&
                                                                 s.ActivityType == "Partsorder" &&
                                                                 s.ActivityCode != "SL1"
                                                                 orderby s.StartDate
                                                                 select s).ToList();


                                if (partsordersForThisMachine != null && partsordersForThisMachine.Count() > 0)
                                {
                                    if (partsordersForThisMachine.Count() >= numberOfPartsordersLevel[i])
                                    {
                                        DateTime previousOrderDate = new DateTime(2000, 1, 1);
                                        int partsorderCounter = 0;
                                        foreach (var intervention in partsordersForThisMachine)
                                        {
                                            if (previousOrderDate.Date != intervention.StartDate.Date)
                                            {
                                                partsorderCounter++;
                                            }
                                            previousOrderDate = intervention.StartDate;
                                        }
                                        if (partsorderCounter >= numberOfPartsordersLevel[i])
                                        {
                                            numberOfPartsOrders = partsorderCounter;
                                        }
                                        else
                                        {
                                            numberOfPartsOrders = 0;
                                        }
                                    }
                                    else
                                    {
                                        numberOfPartsOrders = 0;
                                    }
                                }

                                var callsForThisMachine = (from s in myInterventionsData.AlarmActivities
                                                           where s.SerialNumber == item.SerialNumber &&
                                                           s.StartDate >= startDateToInvestigate &&
                                                           s.ActivityType == "Call"
                                                           orderby s.StartDate
                                                           select s).ToList();
                                if (callsForThisMachine != null && callsForThisMachine.Count() > 0)
                                {
                                    if (callsForThisMachine.Count() >= numberOfCallsLevel[i])
                                    {
                                        DateTime previousCallDate = new DateTime(2000, 1, 1);
                                        int callCounter = 0;
                                        foreach (var intervention in callsForThisMachine)
                                        {
                                            if (previousCallDate.Date != intervention.StartDate.Date)
                                            {
                                                callCounter++;
                                            }
                                            previousCallDate = intervention.StartDate;
                                        }
                                        if (callCounter >= numberOfCallsLevel[i])
                                        {
                                            numberOfCalls = callCounter;
                                        }
                                        else
                                        {
                                            numberOfCalls = 0;
                                        }
                                    }
                                    else
                                    {
                                        numberOfCalls = 0;
                                    }
                                }

                                if (combinedLevel[i])
                                {
                                    if ((numberOfCalls + numberOfInterventions + numberOfPartsOrders) >= (numberOfCallsLevel[i] + numberOfInterventionsLevel[i] + numberOfPartsordersLevel[i]))
                                    {
                                        keepAlarm = true;
                                    }
                                }
                                else
                                {
                                    if (numberOfCalls >= numberOfCallsLevel[i] || numberOfInterventions >= numberOfInterventionsLevel[i] || numberOfPartsOrders >= numberOfPartsordersLevel[i])
                                    {
                                        keepAlarm = true;
                                    }
                                }
                                if (!keepAlarm)
                                {
                                    item.AlarmLevel = i;
                                    if (item.AlarmLevel < 0)
                                    {
                                        item.AlarmLevel = 0;
                                    }
                                    item.DateOfAlarm = DateTime.Now;

                                }
                                else
                                {
                                    break;
                                }
                            }
                            myInterventionsData.SaveChanges();
                        }
                    }

                    #endregion

                    #region Generate alarms

                    //Create a list with machines which have had interventions in the last 6 months
                    DateTime initialDateToVerify = DateTime.Now.AddYears(-3);
                    periodToEvaluate = DateTime.Now.AddMonths(-levelMonths[0]);
                    DateTime alarmPeriodToEvaluate = DateTime.Now.AddMonths(-1);

                    var listWithMachines = (from s in myInterventionsData.AlarmActivities
                                            join t in myInterventionsData.Machines on s.SerialNumber equals t.SerialNumber
                                            where s.StartDate >= periodToEvaluate
                                            orderby s.SerialNumber
                                            select new
                                            {
                                                SerialNumber = s.SerialNumber,
                                                InstallationDate = t.InstallationDate,
                                                technology = t.Technology
                                            }).Distinct().ToList().OrderBy(x => x.SerialNumber);



                    if (listWithMachines != null)
                    {
                        //For each machine:
                        //Check if there is an active alarm (alarm that is not more than one month old)
                        //If there is an alarm, then what level?
                        int level = 1;

                        foreach (var machine in listWithMachines)
                        {

                            string tempString = "";
                            try
                            {
                                tempString = machine.SerialNumber.Substring(0, 5);
                            }
                            catch
                            {
                                tempString = "";
                            }
                            if (tempString != "99999")
                            {
                                var machineWithAlarm = (from s in myInterventionsData.Alarms
                                                        where s.SerialNumber == machine.SerialNumber
                                                        orderby s.DateOfAlarm descending
                                                        select s).FirstOrDefault();
                                if (machineWithAlarm == null || machineWithAlarm.AlarmLevel == 0)
                                {
                                    #region Level 1
                                    //If there is no alarm active, then check if a new alarm hast to be generated.
                                    level = 1;
                                    //Get a list of activities 
                                    DateTime startDateToInvestigate = DateTime.Now.AddMonths((levelMonths[level - 1]) * (-1));
                                    bool _includeIBS = includeIBS[level - 1];
                                    bool _allInterventions = allInterventions[level - 1];
                                    var InterventionsForThisMachine = (from s in myInterventionsData.AlarmActivities
                                                                       where s.SerialNumber == machine.SerialNumber &&
                                                                       s.StartDate >= startDateToInvestigate &&
                                                                        s.StartDate <= DateTime.Now &&
                                                                       s.ActivityType == "Intervention" &&
                                                                           (!(ibsCodes.Contains(s.ActivityCode)) || _includeIBS) &&
                                                                           (warrantyCodes.Contains(s.ActivityCode) || _allInterventions) &&
                                                                    (!(codesToBeExcluded.Contains(s.ActivityCode)))
                                                                       orderby s.StartDate
                                                                       select s).ToList();

                                    Alarms newAlarm = new Alarms();
                                    bool addAlarm = false;
                                    if (InterventionsForThisMachine != null && InterventionsForThisMachine.Count() > 0)
                                    {
                                        if (InterventionsForThisMachine.Count() >= numberOfInterventionsLevel[level - 1])
                                        {
                                            //Check and see if all are valid
                                            //Consecutive interventions and interventiosn at the same time as one
                                            DateTime previousStartDate = new DateTime(2000, 1, 1);
                                            DateTime previousEndDate = new DateTime(2000, 1, 1);
                                            int interventionCounter = 0;
                                            foreach (var intervention in InterventionsForThisMachine)
                                            {
                                                //Eliminate the weekend in between
                                                if (previousEndDate.DayOfWeek == DayOfWeek.Friday)
                                                {
                                                    previousEndDate = previousEndDate.AddDays(1);
                                                }
                                                if (previousEndDate.DayOfWeek == DayOfWeek.Saturday)
                                                {
                                                    previousEndDate = previousEndDate.AddDays(1);
                                                }
                                                if (previousEndDate.DayOfWeek == DayOfWeek.Sunday)
                                                {
                                                    previousEndDate = previousEndDate.AddDays(1);
                                                }
                                                if (previousStartDate.Date != intervention.StartDate.Date && previousEndDate.Date < intervention.StartDate.Date)
                                                {
                                                    interventionCounter++;
                                                }
                                                previousStartDate = intervention.StartDate;
                                                previousEndDate = intervention.EndDate;
                                            }
                                            if (interventionCounter >= numberOfInterventionsLevel[level - 1])
                                            {
                                                newAlarm.NumberOfInterventions = interventionCounter;
                                            }
                                            else
                                            {
                                                newAlarm.NumberOfInterventions = 0;
                                            }

                                        }
                                        else
                                        {
                                            newAlarm.NumberOfInterventions = 0;
                                        }

                                    }

                                    //Only warrantyParts
                                    var partsordersForThisMachine = (from s in myInterventionsData.AlarmActivities
                                                                     where s.SerialNumber == machine.SerialNumber &&
                                                                     s.StartDate >= startDateToInvestigate &&
                                                                      s.StartDate <= DateTime.Now &&
                                                                     s.ActivityType == "Partsorder" &&
                                                                     s.ActivityCode != "SL1"
                                                                     orderby s.StartDate
                                                                     select s).ToList();


                                    if (partsordersForThisMachine != null && partsordersForThisMachine.Count() > 0)
                                    {
                                        if (partsordersForThisMachine.Count() >= numberOfPartsordersLevel[level - 1])
                                        {
                                            DateTime previousOrderDate = new DateTime(2000, 1, 1);
                                            int partsorderCounter = 0;
                                            foreach (var intervention in partsordersForThisMachine)
                                            {
                                                if (previousOrderDate.Date != intervention.StartDate.Date)
                                                {
                                                    partsorderCounter++;
                                                }
                                                previousOrderDate = intervention.StartDate;
                                            }
                                            if (partsorderCounter >= numberOfPartsordersLevel[level - 1])
                                            {
                                                newAlarm.NumberOfPartsOrders = partsorderCounter;
                                            }
                                            else
                                            {
                                                newAlarm.NumberOfPartsOrders = 0;
                                            }
                                        }
                                        else
                                        {
                                            newAlarm.NumberOfPartsOrders = 0;
                                        }
                                    }

                                    var callsForThisMachine = (from s in myInterventionsData.AlarmActivities
                                                               where s.SerialNumber == machine.SerialNumber &&
                                                               s.StartDate >= startDateToInvestigate &&
                                                                s.StartDate <= DateTime.Now &&
                                                               s.ActivityType == "Call"
                                                               orderby s.StartDate
                                                               select s).ToList();
                                    if (callsForThisMachine != null && callsForThisMachine.Count() > 0)
                                    {
                                        if (callsForThisMachine.Count() >= numberOfCallsLevel[level - 1])
                                        {
                                            DateTime previousCallDate = new DateTime(2000, 1, 1);
                                            int callCounter = 0;
                                            foreach (var intervention in callsForThisMachine)
                                            {
                                                if (previousCallDate.Date != intervention.StartDate.Date)
                                                {
                                                    callCounter++;
                                                }
                                                previousCallDate = intervention.StartDate;
                                            }
                                            if (callCounter >= numberOfCallsLevel[level - 1])
                                            {
                                                newAlarm.NumberOfCalls = callCounter;
                                            }
                                            else
                                            {
                                                newAlarm.NumberOfCalls = 0;
                                            }
                                        }
                                        else
                                        {
                                            newAlarm.NumberOfCalls = 0;
                                        }
                                    }

                                    if (combinedLevel[level - 1])
                                    {
                                        if ((newAlarm.NumberOfCalls + newAlarm.NumberOfInterventions + newAlarm.NumberOfPartsOrders) >= (numberOfCallsLevel[level - 1] + numberOfInterventionsLevel[level - 1] + numberOfPartsordersLevel[level - 1]))
                                        {
                                            addAlarm = true;
                                        }
                                    }
                                    else
                                    {
                                        if (newAlarm.NumberOfCalls >= numberOfCallsLevel[level - 1] || newAlarm.NumberOfInterventions >= numberOfInterventionsLevel[level - 1] || newAlarm.NumberOfPartsOrders >= numberOfPartsordersLevel[level - 1])
                                        {
                                            addAlarm = true;
                                        }
                                    }
                                    if (addAlarm)
                                    {
                                        newAlarm.AlarmLevel = level;
                                        newAlarm.DateOfAlarm = DateTime.Now;
                                        newAlarm.MailsSent = false;
                                        newAlarm.SerialNumber = machine.SerialNumber;
                                        newAlarm.Technology = machine.technology;
                                        var alarmForThisMachine = (from s in myInterventionsData.Alarms
                                                                   where s.SerialNumber == machine.SerialNumber
                                                                   select s).FirstOrDefault();
                                        if (alarmForThisMachine != null)
                                        {
                                            alarmForThisMachine.AlarmLevel = newAlarm.AlarmLevel;
                                            alarmForThisMachine.DateOfAlarm = newAlarm.DateOfAlarm;
                                            alarmForThisMachine.MailsSent = newAlarm.MailsSent;
                                            alarmForThisMachine.NumberOfCalls = newAlarm.NumberOfCalls;
                                            alarmForThisMachine.NumberOfInterventions = newAlarm.NumberOfInterventions;
                                            alarmForThisMachine.NumberOfPartsOrders = newAlarm.NumberOfPartsOrders;
                                            alarmForThisMachine.SerialNumber = newAlarm.SerialNumber;
                                        }
                                        else
                                        {
                                            myInterventionsData.Alarms.Add(newAlarm);
                                        }


                                        AlarmHistory newHistory = new AlarmHistory();
                                        newHistory.AlarmLevel = level;
                                        newHistory.DateOfAlarm = DateTime.Now;
                                        newHistory.NumberOfCalls = newAlarm.NumberOfCalls;
                                        newHistory.NumberOfInterventions = newAlarm.NumberOfInterventions;
                                        newHistory.NumberOfPartsOrders = newAlarm.NumberOfPartsOrders;
                                        newHistory.SerialNumber = machine.SerialNumber;
                                        newHistory.Technology = machine.technology;
                                        myInterventionsData.AlarmHistory.Add(newHistory);

                                        //Add the machine to the mailList
                                        var isMachineInMailList = (from s in myInterventionsData.AlarmSalesMail
                                                                   where s.SerialNumber == machine.SerialNumber
                                                                   select s).FirstOrDefault();
                                        if (isMachineInMailList == null)
                                        {
                                            AlarmSalesMail newEntry = new AlarmSalesMail();
                                            newEntry.SerialNumber = machine.SerialNumber;
                                            newEntry.LocalSalesMail = "";
                                            myInterventionsData.AlarmSalesMail.Add(newEntry);

                                        }
                                        myInterventionsData.SaveChanges();
                                    }

                                    #endregion
                                }
                                else
                                {
                                    if (machineWithAlarm.AlarmLevel == 1)
                                    {
                                        #region level 2
                                        //Machine is in level 1
                                        level = 2;

                                        //Check if the number of interventions in the period of level 2 is not bigger than allowed 
                                        DateTime startDateToInvestigate = machineWithAlarm.DateOfAlarm;
                                        if (machineWithAlarm.DateOfAlarm <= DateTime.Now.AddMonths(((-1) * levelMonths[level - 1])))
                                        {
                                            startDateToInvestigate = DateTime.Now.AddMonths(((-1) * levelMonths[level - 1]));
                                        }
                                        DateTime endDateToInvestigate = startDateToInvestigate.AddMonths((levelMonths[level - 1]));

                                        //Current date must be between Start and enddate 
                                        if (DateTime.Now >= startDateToInvestigate && DateTime.Now <= endDateToInvestigate)
                                        {
                                            bool _includeIBS = includeIBS[level - 1];
                                            bool _allInterventions = allInterventions[level - 1];
                                            var InterventionsForThisMachine = (from s in myInterventionsData.AlarmActivities
                                                                               where s.SerialNumber == machine.SerialNumber &&
                                                                               s.StartDate >= startDateToInvestigate &&
                                                                               s.StartDate <= endDateToInvestigate &&
                                                                               s.ActivityType == "Intervention" &&
                                                                               (!(ibsCodes.Contains(s.ActivityCode)) || _includeIBS) &&
                                                                               (warrantyCodes.Contains(s.ActivityCode) || _allInterventions) &&
                                                                    (!(codesToBeExcluded.Contains(s.ActivityCode)))
                                                                               orderby s.StartDate
                                                                               select s).ToList();

                                            bool modifyAlarm = false;
                                            machineWithAlarm.NumberOfInterventions = 0;
                                            machineWithAlarm.NumberOfCalls = 0;
                                            machineWithAlarm.NumberOfPartsOrders = 0;
                                            if (InterventionsForThisMachine != null && InterventionsForThisMachine.Count() > 0)
                                            {
                                                if (InterventionsForThisMachine.Count() >= numberOfInterventionsLevel[level - 1])
                                                {
                                                    //Check and see if all are valid
                                                    //Consecutive interventions and interventiosn at the same time as one
                                                    int interventionCounter = CalculateInterventions(InterventionsForThisMachine);
                                                    if (interventionCounter >= numberOfInterventionsLevel[level - 1])
                                                    {
                                                        machineWithAlarm.NumberOfInterventions = interventionCounter;
                                                    }
                                                }
                                            }

                                            //Only warrantyParts
                                            var partsordersForThisMachine = (from s in myInterventionsData.AlarmActivities
                                                                             where s.SerialNumber == machine.SerialNumber &&
                                                                              s.StartDate >= startDateToInvestigate &&
                                                                               s.StartDate <= endDateToInvestigate &&
                                                                             s.ActivityType == "Partsorder" &&
                                                                             s.ActivityCode != "SL1"
                                                                             orderby s.StartDate
                                                                             select s).ToList();


                                            if (partsordersForThisMachine != null && partsordersForThisMachine.Count() > 0)
                                            {
                                                if (partsordersForThisMachine.Count() >= numberOfPartsordersLevel[level - 1])
                                                {
                                                    DateTime previousOrderDate = new DateTime(2000, 1, 1);
                                                    int partsorderCounter = 0;
                                                    foreach (var intervention in partsordersForThisMachine)
                                                    {
                                                        if (previousOrderDate.Date != intervention.StartDate.Date)
                                                        {
                                                            partsorderCounter++;
                                                        }
                                                        previousOrderDate = intervention.StartDate;
                                                    }
                                                    if (partsorderCounter >= numberOfPartsordersLevel[level - 1])
                                                    {
                                                        machineWithAlarm.NumberOfPartsOrders = partsorderCounter;
                                                    }
                                                }
                                            }

                                            var callsForThisMachine = (from s in myInterventionsData.AlarmActivities
                                                                       where s.SerialNumber == machine.SerialNumber &&
                                                                        s.StartDate >= startDateToInvestigate &&
                                                                               s.StartDate <= endDateToInvestigate &&
                                                                       s.ActivityType == "Call"
                                                                       orderby s.StartDate
                                                                       select s).ToList();
                                            if (callsForThisMachine != null && callsForThisMachine.Count() > 0)
                                            {
                                                if (callsForThisMachine.Count() >= numberOfCallsLevel[level - 1])
                                                {
                                                    DateTime previousCallDate = new DateTime(2000, 1, 1);
                                                    int callCounter = 0;
                                                    foreach (var intervention in callsForThisMachine)
                                                    {
                                                        if (previousCallDate.Date != intervention.StartDate.Date)
                                                        {
                                                            callCounter++;
                                                        }
                                                        previousCallDate = intervention.StartDate;
                                                    }
                                                    if (callCounter >= numberOfCallsLevel[level - 1])
                                                    {
                                                        machineWithAlarm.NumberOfCalls = callCounter;
                                                    }
                                                }
                                            }

                                            if (combinedLevel[level - 1])
                                            {
                                                if ((machineWithAlarm.NumberOfCalls + machineWithAlarm.NumberOfInterventions + machineWithAlarm.NumberOfPartsOrders) >= (numberOfCallsLevel[level - 1] + numberOfInterventionsLevel[level - 1] + numberOfPartsordersLevel[level - 1]))
                                                {
                                                    modifyAlarm = true;
                                                }
                                            }
                                            else
                                            {
                                                if (machineWithAlarm.NumberOfCalls >= numberOfCallsLevel[level - 1] || machineWithAlarm.NumberOfInterventions >= numberOfInterventionsLevel[level - 1] || machineWithAlarm.NumberOfPartsOrders >= numberOfPartsordersLevel[level - 1])
                                                {
                                                    modifyAlarm = true;
                                                }
                                            }
                                            if (modifyAlarm)
                                            {
                                                machineWithAlarm.AlarmLevel = level;
                                                machineWithAlarm.DateOfAlarm = DateTime.Now;
                                                machineWithAlarm.MailsSent = false;
                                                machineWithAlarm.Technology = machine.technology;
                                                myInterventionsData.SaveChanges();

                                                AlarmHistory newHistory = new AlarmHistory();
                                                newHistory.AlarmLevel = level;
                                                newHistory.DateOfAlarm = DateTime.Now;
                                                newHistory.NumberOfCalls = machineWithAlarm.NumberOfCalls;
                                                newHistory.NumberOfInterventions = machineWithAlarm.NumberOfInterventions;
                                                newHistory.NumberOfPartsOrders = machineWithAlarm.NumberOfPartsOrders;
                                                newHistory.SerialNumber = machine.SerialNumber;
                                                newHistory.Technology = machine.technology;
                                                myInterventionsData.AlarmHistory.Add(newHistory);
                                                myInterventionsData.SaveChanges();
                                            }
                                        }
                                        #endregion
                                    }
                                    else
                                    {

                                        if (machineWithAlarm.AlarmLevel == 2)
                                        {
                                            #region level 3
                                            //Machine is in level 2
                                            level = 3;
                                            //Check if the number of interventions in the period of level 3 is not bigger than allowed 
                                            DateTime startDateToInvestigate = machineWithAlarm.DateOfAlarm;
                                            if (machineWithAlarm.DateOfAlarm <= DateTime.Now.AddMonths(((-1) * levelMonths[level - 1])))
                                            {
                                                startDateToInvestigate = DateTime.Now.AddMonths(((-1) * levelMonths[level - 1]));
                                            }
                                            //DateTime endDateToInvestigate = machineWithAlarm.DateOfAlarm.AddDays((daysForLevel[level - 1]));
                                            DateTime endDateToInvestigate = startDateToInvestigate.AddMonths((levelMonths[level - 1]));

                                            //Current date must be between Start and enddate 
                                            if (DateTime.Now >= startDateToInvestigate && DateTime.Now <= endDateToInvestigate)
                                            {
                                                bool _includeIBS = includeIBS[level - 1];
                                                bool _allInterventions = allInterventions[level - 1];
                                                var InterventionsForThisMachine = (from s in myInterventionsData.AlarmActivities
                                                                                   where s.SerialNumber == machine.SerialNumber &&
                                                                                   s.StartDate >= startDateToInvestigate &&
                                                                                    s.StartDate <= endDateToInvestigate &&
                                                                                   s.ActivityType == "Intervention" &&
                                                                                   (!(ibsCodes.Contains(s.ActivityCode)) || _includeIBS) &&
                                                                                   (warrantyCodes.Contains(s.ActivityCode) || _allInterventions) &&
                                                                        (!(codesToBeExcluded.Contains(s.ActivityCode)))
                                                                                   orderby s.StartDate
                                                                                   select s).ToList();

                                                bool modifyAlarm = false;
                                                machineWithAlarm.NumberOfInterventions = 0;
                                                machineWithAlarm.NumberOfCalls = 0;
                                                machineWithAlarm.NumberOfPartsOrders = 0;
                                                if (InterventionsForThisMachine != null && InterventionsForThisMachine.Count() > 0)
                                                {
                                                    if (InterventionsForThisMachine.Count() >= numberOfInterventionsLevel[level - 1])
                                                    {
                                                        //Check and see if all are valid
                                                        //Consecutive interventions and interventiosn at the same time as one
                                                        int interventionCounter = CalculateInterventions(InterventionsForThisMachine);
                                                        if (interventionCounter >= numberOfInterventionsLevel[level - 1])
                                                        {
                                                            machineWithAlarm.NumberOfInterventions = interventionCounter;
                                                        }
                                                    }
                                                }

                                                //Only warrantyParts
                                                var partsordersForThisMachine = (from s in myInterventionsData.AlarmActivities
                                                                                 where s.SerialNumber == machine.SerialNumber &&
                                                                                 s.StartDate >= startDateToInvestigate &&
                                                                                   s.StartDate <= endDateToInvestigate &&
                                                                                 s.ActivityType == "Partsorder" &&
                                                                                 s.ActivityCode != "SL1"
                                                                                 orderby s.StartDate
                                                                                 select s).ToList();


                                                if (partsordersForThisMachine != null && partsordersForThisMachine.Count() > 0)
                                                {
                                                    if (partsordersForThisMachine.Count() >= numberOfPartsordersLevel[level - 1])
                                                    {
                                                        DateTime previousOrderDate = new DateTime(2000, 1, 1);
                                                        int partsorderCounter = 0;
                                                        foreach (var intervention in partsordersForThisMachine)
                                                        {
                                                            if (previousOrderDate.Date != intervention.StartDate.Date)
                                                            {
                                                                partsorderCounter++;
                                                            }
                                                            previousOrderDate = intervention.StartDate;
                                                        }
                                                        if (partsorderCounter >= numberOfPartsordersLevel[level - 1])
                                                        {
                                                            machineWithAlarm.NumberOfPartsOrders = partsorderCounter;
                                                        }
                                                    }
                                                }

                                                var callsForThisMachine = (from s in myInterventionsData.AlarmActivities
                                                                           where s.SerialNumber == machine.SerialNumber &&
                                                                           s.StartDate >= startDateToInvestigate &&
                                                                                   s.StartDate <= endDateToInvestigate &&
                                                                           s.ActivityType == "Call"
                                                                           orderby s.StartDate
                                                                           select s).ToList();
                                                if (callsForThisMachine != null && callsForThisMachine.Count() > 0)
                                                {
                                                    if (callsForThisMachine.Count() >= numberOfCallsLevel[level - 1])
                                                    {
                                                        DateTime previousCallDate = new DateTime(2000, 1, 1);
                                                        int callCounter = 0;
                                                        foreach (var intervention in callsForThisMachine)
                                                        {
                                                            if (previousCallDate.Date != intervention.StartDate.Date)
                                                            {
                                                                callCounter++;
                                                            }
                                                            previousCallDate = intervention.StartDate;
                                                        }
                                                        if (callCounter >= numberOfCallsLevel[level - 1])
                                                        {
                                                            machineWithAlarm.NumberOfCalls = callCounter;
                                                        }
                                                    }
                                                }

                                                if (combinedLevel[level - 1])
                                                {
                                                    if ((machineWithAlarm.NumberOfCalls + machineWithAlarm.NumberOfInterventions + machineWithAlarm.NumberOfPartsOrders) >= (numberOfCallsLevel[level - 1] + numberOfInterventionsLevel[level - 1] + numberOfPartsordersLevel[level - 1]))
                                                    {
                                                        modifyAlarm = true;
                                                    }
                                                }
                                                else
                                                {
                                                    if (machineWithAlarm.NumberOfCalls >= numberOfCallsLevel[level - 1] || machineWithAlarm.NumberOfInterventions >= numberOfInterventionsLevel[level - 1] || machineWithAlarm.NumberOfPartsOrders >= numberOfPartsordersLevel[level - 1])
                                                    {
                                                        modifyAlarm = true;
                                                    }
                                                }
                                                if (modifyAlarm)
                                                {
                                                    machineWithAlarm.AlarmLevel = level;
                                                    machineWithAlarm.DateOfAlarm = DateTime.Now;
                                                    machineWithAlarm.MailsSent = false;
                                                    machineWithAlarm.Technology = machine.technology;
                                                    myInterventionsData.SaveChanges();

                                                    AlarmHistory newHistory = new AlarmHistory();
                                                    newHistory.AlarmLevel = level;
                                                    newHistory.DateOfAlarm = DateTime.Now;
                                                    newHistory.NumberOfCalls = machineWithAlarm.NumberOfCalls;
                                                    newHistory.NumberOfInterventions = machineWithAlarm.NumberOfInterventions;
                                                    newHistory.NumberOfPartsOrders = machineWithAlarm.NumberOfPartsOrders;
                                                    newHistory.SerialNumber = machine.SerialNumber;
                                                    newHistory.Technology = machine.technology;
                                                    myInterventionsData.AlarmHistory.Add(newHistory);
                                                    myInterventionsData.SaveChanges();
                                                }
                                            }
                                            #endregion
                                        }

                                    }
                                }
                            }
                        }
                    }
                    #endregion

                    #region send mails

                    var mailsToSend = from s in myInterventionsData.Alarms
                                      where s.MailsSent == false && s.AlarmLevel > 0
                                      select s;
                    if (mailsToSend != null)
                    {
                        foreach (var alarm in mailsToSend)
                        {
                            //Get the mailaddresses for this machine

                            var thisMachineMails = (from s in myInterventionsData.AlarmSalesMail
                                                    where s.SerialNumber == alarm.SerialNumber
                                                    select s).FirstOrDefault();
                            //Check for which department
                            string thisDepartment = "LVDWRL";
                            DateTime NineMonthsBack = DateTime.Now.AddMonths(-9);
                            var thisMachinesActivities = (from s in myInterventionsData.AlarmActivities
                                                          where s.SerialNumber == alarm.SerialNumber &&
                                                          s.Department != "LVDWRL" &&
                                                          s.Department != ""
                                                          && s.StartDate > NineMonthsBack
                                                          orderby s.DateCreated descending
                                                          select s).FirstOrDefault();
                            if (thisMachinesActivities != null)
                            {
                                thisDepartment = thisMachinesActivities.Department;
                            }

                            List<string> mailAddres = new List<string>();
                            bool setmailsAsSent = true;

                            var mailAddresses = (from s in myInterventionsData.AlarmDepartmentMails
                                                 where s.Department == thisDepartment
                                                 select s).FirstOrDefault();
                            if (mailAddresses != null)
                            {
                                #region Mails Level 1
                                if (alarm.AlarmLevel == 1)
                                {
                                    if (thisMachineMails != null)
                                    {
                                        if (thisMachineMails.LocalSalesMail == "")
                                        {
                                            setmailsAsSent = false;
                                        }
                                    }
                                    else
                                    {
                                        setmailsAsSent = false;
                                    }
                                    if (setmailsAsSent)
                                    {
                                        if (mailAddresses.LocalServiceManager.Trim() != "")
                                        {
                                            mailAddres.Add(CleanUp(mailAddresses.LocalServiceManager.Trim()));
                                        }
                                        if (thisMachineMails.LocalSalesMail.Trim() != "")
                                        {
                                            mailAddres.Add(CleanUp(thisMachineMails.LocalSalesMail.Trim()));
                                        }
                                        if (mailAddresses.LocalSalesManager.Trim() != "")
                                        {
                                            mailAddres.Add(CleanUp(mailAddresses.LocalSalesManager.Trim()));
                                        }

                                    }
                                }
                                #endregion

                                #region Mails Level 2             
                                if (alarm.AlarmLevel == 2)
                                {
                                    // 17-01-2018 jbac - added the setmailAsSent check on lvl 2 as well
                                    // the mails will now be re-send each day as long as the localsalesmail is empty !
                                    if (thisMachineMails != null)
                                    {
                                        if (thisMachineMails.LocalSalesMail == "")
                                        {
                                            setmailsAsSent = false;
                                        }
                                    }
                                    else
                                    {
                                        setmailsAsSent = false;
                                    }
                                    // 17-01-2018
                                    if (mailAddresses.LocalGeneralManager.Trim()!="")
                                    {
                                        mailAddres.Add(CleanUp(mailAddresses.LocalGeneralManager.Trim()));
                                    }
                                    if (mailAddresses.CompanyLocalRepresentative.Trim() != "")
                                    {
                                        mailAddres.Add(CleanUp(mailAddresses.CompanyLocalRepresentative.Trim()));
                                    }
                                    if (mailAddresses.CompanyServiceManager.Trim() != "")
                                    {
                                        mailAddres.Add(CleanUp(mailAddresses.CompanyServiceManager.Trim()));
                                    }
                                    if (mailAddresses.LocalServiceManager.Trim() != "")
                                    {
                                        mailAddres.Add(CleanUp(mailAddresses.LocalServiceManager.Trim()));
                                    }

                                }
                                #endregion

                                #region Mails Level 3
                                if (alarm.AlarmLevel == 3)
                                {
                                    if (mailAddresses.LocalGeneralManager.Trim()!="")
                                    {
                                        mailAddres.Add(CleanUp(mailAddresses.LocalGeneralManager.Trim()));
                                    }
                                    if (mailAddresses.CEO.Trim()!="")
                                    {
                                        mailAddres.Add(CleanUp(mailAddresses.CEO.Trim()));
                                    }
                                    if (mailAddresses.LocalServiceManager.Trim()!="")
                                    {
                                        mailAddres.Add(CleanUp(mailAddresses.LocalServiceManager.Trim()));
                                    }
                                }
                                #endregion
                            }
                            #region Individual Mails
                            var individualsToAdd = (from s in myInterventionsData.IndividualAlarmMails
                                                    where s.Alarmlevel == alarm.AlarmLevel && 
                                                          (s.Technology == alarm.Technology || s.Technology == "All")
                                                    select s).ToList();
                            if (individualsToAdd != null)
                            {
                                foreach (var mailAdress in individualsToAdd)
                                {
                                    if (mailAdress.email.Trim() != "")
                                    {
                                        mailAddres.Add(CleanUp(mailAdress.email));
                                    }
                                }
                            }
                            #endregion

                            //Find more information regarding the machine based on the last interventions
                            string customer = "";
                            string machineModel = "";
                            string endOfWarranty = "";
                            string mostRecentIntervention = "";
                            var lastIntervention = (from s in myInterventionsData.ServiceOrder
                                                    where s.SerialNr == alarm.SerialNumber && s.Verwerking != "Cancelled"
                                                    orderby s.ID descending
                                                    select s).Take(5);
                            if (lastIntervention != null)
                            {
                                try
                                {
                                    customer = lastIntervention.FirstOrDefault().Omschr20;
                                }
                                catch { }
                                try
                                {
                                    machineModel = lastIntervention.FirstOrDefault().ProdDesc;
                                }
                                catch { }
                                try
                                {
                                    endOfWarranty = lastIntervention.FirstOrDefault().EindeGarantie.ToString();
                                }
                                catch { }
                                try
                                {
                                    StringBuilder sb = new StringBuilder();
                                    sb.AppendLine("");
                                    foreach (var so in lastIntervention)
                                    {                                       
                                        sb.AppendLine(string.Format("Order : {0} {1} Order Status : {2}  Planned Start Date : ({3})", so.SrvOrder, so.CallTextDetail, so.Verwerking, so.EarliestStartTime));
                                    }
                                    mostRecentIntervention = sb.ToString();
                                }
                                catch {}
                            }

                            try
                            {
                                //Create the PDF report for this machine  
                                int monthsBack = 0;
                                switch (alarm.AlarmLevel)
                                {
                                    case 1:
                                        monthsBack = levelMonths[0];
                                        break;
                                    case 2:
                                        monthsBack = levelMonths[0] + levelMonths[1];
                                        break;
                                    case 3:
                                        monthsBack = levelMonths[0] + levelMonths[1] + levelMonths[2];
                                        break;
                                    default:
                                        monthsBack = levelMonths[0];
                                        break;
                                }

                                StringBuilder newBody = new StringBuilder();
                                newBody.AppendFormat("Warning level {0} has been reached for:", alarm.AlarmLevel);
                                newBody.AppendLine();
                                newBody.AppendFormat("Customer: {0}", customer);
                                newBody.AppendLine();
                                newBody.AppendFormat("Serial Number: {0}", alarm.SerialNumber);
                                newBody.AppendLine();
                                newBody.AppendFormat("Machine: {0}", machineModel);
                                newBody.AppendLine();
                                newBody.AppendFormat("Most recent Interventions: {0}", mostRecentIntervention);
                                newBody.AppendLine();
                                newBody.AppendLine();
                                newBody.Append("More details can be found in the attached file. This file contains information made available by electronic registered service interventions and spareparts orders in the ERP-LN system of LVD and Strippit which have been registered with the serialnumber of this machine");
                                newBody.AppendLine();
                                newBody.Append("This mail has been generated automatically. Do not reply to this mail address");
                                newBody.AppendLine();

                                string fileToAttach = PrintTechnicalReports(alarm.SerialNumber, monthsBack, true, mailAddres.Distinct().FirstOrDefault());
                                LVDMail.LVDMail lvdMail = new LVDMail.LVDMail();
                                lvdMail.MailAccount = "Fieldservice";
                                lvdMail.MailSubject = String.Format("Warning level {0} has been reached for a machine at {1} ", alarm.AlarmLevel, customer);
                                lvdMail.MailBody = newBody.ToString();

                                lvdMail.MailToArray = mailAddres.Distinct().ToArray();

                                if (fileToAttach != "")
                                {
                                    lvdMail.MailAttachments = new string[] { fileToAttach };
                                }
                                string mailResponse = lvdMail.SendMail();
                                if (mailResponse == "")
                                {
                                    alarm.MailsSent = setmailsAsSent;
                                }
                                else
                                {
                                    alarm.MailsSent = false;
                                }


                            }
                            catch {}

                        }
                        myInterventionsData.SaveChanges();
                    }
                    #endregion
                }
            }
        }

        private static string CleanUp(String StringToCleanUp)
        {
            return StringToCleanUp.Replace("\r", "").Replace("\n", "");
        }

        private static int CalculateInterventions(List<AlarmActivities> InterventionsForThisMachine)
        {
            DateTime previousStartDate = new DateTime(2000, 1, 1);
            DateTime previousEndDate = new DateTime(2000, 1, 1);
            int interventionCounter = 0;
            foreach (var intervention in InterventionsForThisMachine)
            {
                //Eliminate the weekend in between
                if (previousEndDate.DayOfWeek == DayOfWeek.Friday)
                {
                    previousEndDate = previousEndDate.AddDays(1);
                }
                if (previousEndDate.DayOfWeek == DayOfWeek.Saturday)
                {
                    previousEndDate = previousEndDate.AddDays(1);
                }
                if (previousEndDate.DayOfWeek == DayOfWeek.Sunday)
                {
                    previousEndDate = previousEndDate.AddDays(1);
                }
                if (previousStartDate.Date != intervention.StartDate.Date && previousEndDate.Date < intervention.StartDate.Date)
                {
                    interventionCounter++;
                }
                previousStartDate = intervention.StartDate;
                previousEndDate = intervention.EndDate;
            }
            return interventionCounter;
        }

        public static string PrintTechnicalReports(string searchSerialNumber = "", int monthsBack = 0, bool customer = false, string thisUser = "")
        {
            string fileName = "";
            using (ERPLNDataEntities myERPLNData = new ERPLNDataEntities())
            {
                using (InterventionsEntities myInterventionsData = new InterventionsEntities())
                {
                    if (searchSerialNumber != "")
                    {
                        DateTime fromdate = DateTime.Now.AddMonths(-1 * monthsBack);

                        string dataPad = @"c:\temp\MedicalFiles";

                        if (!System.IO.Directory.Exists(dataPad))
                        {
                            System.IO.Directory.CreateDirectory(dataPad);
                        }
                        fileName = string.Format("{0}\\Overview_{1}.pdf", dataPad, searchSerialNumber.ToString());
                        try
                        {
                            var interventions = (from s in myInterventionsData.ServiceOrder
                                                 where s.SerialNr == searchSerialNumber &&
                                                 s.PlannedStartTime >= fromdate
                                                 select new
                                                 {
                                                     so = s.SrvOrder
                                                 }).ToList();
                            List<int> listWithInterventions = new List<int>();
                            foreach (var item in interventions)
                            {
                                try
                                {
                                    listWithInterventions.Add((int)item.so);
                                }
                                catch { }
                            }

                            var partsorders = (from s in myERPLNData.ERPLN_ServiceOrders
                                               join t in myERPLNData.Srv_Calls on s.Serviceorder equals t.ServiceOrder
                                               where t.SerianNumberMachine == searchSerialNumber &&
                                               s.orderDate >= fromdate &&
                                               (s.Type == "SL1" || s.Type == "SL2" || s.Type == "SL3" || s.Type == "SL4" || s.Type == "SL5" || s.Type == "SL6")
                                               select new
                                               {
                                                   so = s.Serviceorder
                                               }).ToList();
                            List<int> listWithPartsorders = new List<int>();
                            foreach (var item in partsorders)
                            {
                                listWithPartsorders.Add(item.so);
                            }

                            if (!PrintTechnicalReport(fileName, listWithInterventions, listWithPartsorders, true, thisUser))
                            {
                                fileName = "";
                            }
                        }
                        catch (Exception)
                        {
                            fileName = "";
                        }
                    }
                }
            }
            return fileName;

        }

        public static bool PrintTechnicalReport(string fileName, List<int> interventions, List<int> partsorders, bool customer = false, string thisUser = "")
        {
            if (interventions != null && interventions.Count() > 0)
            {
                bool printHeader = true;
                using (ERPLNDataEntities myERPLNData = new ERPLNDataEntities())
                {
                    using (InterventionsEntities myInterventionsData = new InterventionsEntities())
                    {
                        MigraDoc.DocumentObjectModel.Tables.Table table;
                        Document document = new Document();
                        document.Info.Title = "Technical Reports";
                        document.Info.Subject = "Report for Service with intervention information";
                        document.Info.Author = "FSE";
                        document.DefaultPageSetup.Orientation = Orientation.Portrait;

                        MigraDoc.DocumentObjectModel.Style style = document.Styles["Normal"];
                        style.Font.Name = "Verdana";

                        style = document.Styles[StyleNames.Header];
                        style.ParagraphFormat.AddTabStop("16cm", TabAlignment.Right);

                        style = document.Styles[StyleNames.Footer];
                        style.ParagraphFormat.AddTabStop("8cm", TabAlignment.Center);

                        style = document.Styles.AddStyle("Table", "Normal");
                        style.Font.Name = "Verdana";
                        style.Font.Name = "Times New Roman";
                        style.Font.Size = 9;

                        style = document.Styles[StyleNames.Header];
                        style.ParagraphFormat.AddTabStop("16cm", TabAlignment.Right);

                        style = document.Styles[StyleNames.Footer];
                        style.ParagraphFormat.AddTabStop("8cm", TabAlignment.Center);

                        style = document.Styles.AddStyle("Reference", "Normal");
                        style.ParagraphFormat.SpaceBefore = "5mm";
                        style.ParagraphFormat.SpaceAfter = "5mm";
                        style.ParagraphFormat.TabStops.AddTabStop("16cm", TabAlignment.Right);

                        //style = document.Styles.AddStyle("Reference", "Normal");
                        //style.ParagraphFormat.SpaceBefore = "5mm";
                        //style.ParagraphFormat.SpaceAfter = "5mm";
                        //style.ParagraphFormat.TabStops.AddTabStop("16cm", TabAlignment.Right);

                        Int32 totalRows = 0;
                        Int32 blockRows = 0;
                        Int32 soRows = 0;

                        Section section = document.AddSection();

                        Paragraph paragraph = section.AddParagraph();
                        paragraph.Format.SpaceBefore = "1cm";
                        paragraph.Style = "Reference";
                        if (customer)
                        {
                            paragraph.AddFormattedText("OVERVIEW INTERVENTIONS", TextFormat.Bold);
                        }
                        else
                        {
                            paragraph.AddFormattedText("OVERVIEW INTERVENTIONS (internal Report)", TextFormat.Bold);
                        }

                        Paragraph par = new Paragraph();
                        par.AddText(string.Format("Printed by {0}", thisUser));
                        par.AddTab();
                        par.AddPageField();

                        section.Footers.Primary.Add(par);
                        section.Footers.EvenPage.Add(par.Clone());

                        // Create the item table
                        table = section.AddTable();
                        table.Style = "Table";
                        table.Rows.LeftIndent = 0;

                        // define the columns
                        Column column = table.AddColumn("2.2cm");
                        column.Format.Alignment = ParagraphAlignment.Center;

                        column = table.AddColumn("1.4cm");
                        column.Format.Alignment = ParagraphAlignment.Right;

                        column = table.AddColumn("3.7cm");
                        column.Format.Alignment = ParagraphAlignment.Right;

                        column = table.AddColumn("0.3cm");
                        column.Format.Alignment = ParagraphAlignment.Right;

                        column = table.AddColumn("2.0cm");
                        column.Format.Alignment = ParagraphAlignment.Center;

                        column = table.AddColumn("3.2cm");
                        column.Format.Alignment = ParagraphAlignment.Right;

                        column = table.AddColumn("1.6cm");
                        column.Format.Alignment = ParagraphAlignment.Right;

                        column = table.AddColumn("1.6cm");
                        column.Format.Alignment = ParagraphAlignment.Right;

                        Row row;

                        foreach (var _serviceOrder in interventions)
                        {
                            #region loop
                            var thisTechReport = (from s in myInterventionsData.ServiceReport
                                                  where s.ServiceOrder == _serviceOrder
                                                  join t in myInterventionsData.ServiceOrder on s.ServiceOrder equals t.SrvOrder into t0
                                                  from t1 in t0.DefaultIfEmpty()
                                                  select new
                                                  {
                                                      id = s.ID,
                                                      serviceorder = t1.SrvOrder,
                                                      customer = t1.Omschr20,
                                                      adres1 = t1.Adres1,
                                                      adres2 = t1.Adres2,
                                                      adres3 = t1.Adres3,
                                                      adres4 = t1.Adres4,
                                                      adres5 = t1.Adres5,
                                                      activityLine = t1.Omschr9,
                                                      serialItem = t1.Item,
                                                      serialNumber = t1.SerialNr,
                                                      item = t1.Item,
                                                      bpReference = t1.VerkopenAanRel,
                                                      cluster = t1.Cluster,
                                                      tel = t1.Tel,
                                                      contact = t1.Contact,
                                                      department = t1.Omschr6,
                                                      fseNr = t1.Werknemer,
                                                      fseNaam = t1.Naam1,
                                                      problemCode = t1.Problem,
                                                      problem = t1.Omschr22,
                                                      problemSolved = s.ProblemSolved,
                                                      problemNotSolved = s.ProblemNotSolved,
                                                      garantie = t1.EindeGarantie,
                                                      call = t1.CallTextDetail,
                                                      callNumber = t1.CallText,
                                                      problemDetail = s.ProblemDescription,
                                                      action = s.ActionUndertaken,
                                                      techReport = s.TechReport,
                                                      requests = s.Requests,
                                                      requestsDetails = s.RequestDetails,
                                                      safetyOK = s.SafetyOK,
                                                      safetyNOK = s.SafetyNOK,
                                                      safetyNOKDetails = s.SafetyNOKDetails,
                                                      status = t1.Verwerking,
                                                      srvType = t1.SrvType,
                                                      srvTypeOmschrijving = t1.Omschr8,
                                                      prodDesc = t1.ProdDesc,
                                                      DateSigned = s.DateSigned,
                                                      SignedBy = s.SignedBy,
                                                      EarliestStartTime = t1.EarliestStartTime,
                                                  }).FirstOrDefault();

                            if (thisTechReport != null)
                            {
                                if (printHeader)
                                {
                                    #region General Header

                                    printHeader = false;
                                    // Create the header of the Fist section of the table

                                    blockRows = 0;
                                    row = table.AddRow();
                                    totalRows++;
                                    blockRows++;
                                    row.Height = "0.5cm";
                                    row.HeadingFormat = false;
                                    row.Format.Alignment = ParagraphAlignment.Center;
                                    row.Format.Font.Bold = true;
                                    row.Shading.Color = TableGray;
                                    row.Cells[0].AddParagraph("CUSTOMER");
                                    row.Cells[0].Format.Font.Bold = true;
                                    row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
                                    row.Cells[0].VerticalAlignment = VerticalAlignment.Center;
                                    row.Cells[0].MergeRight = 2;
                                    row.Cells[4].AddParagraph("MACHINE");
                                    row.Cells[4].Format.Alignment = ParagraphAlignment.Left;
                                    row.Cells[4].VerticalAlignment = VerticalAlignment.Center;
                                    row.Cells[4].MergeRight = 3;

                                    row = table.AddRow();
                                    totalRows++;
                                    blockRows++;
                                    row.HeadingFormat = false;
                                    row.Format.Alignment = ParagraphAlignment.Left;
                                    row.Format.Font.Size = "7";
                                    row.Cells[0].AddParagraph("Businesspartner");
                                    row.Cells[0].Format.Font.Bold = true;
                                    row.Cells[1].AddParagraph(thisTechReport.customer);
                                    row.Cells[1].Format.Font.Bold = false;
                                    row.Cells[1].MergeRight = 1;
                                    row.Cells[4].AddParagraph("");
                                    row.Cells[4].Format.Font.Bold = true;
                                    row.Cells[5].AddParagraph("");
                                    row.Cells[5].Format.Font.Bold = false;
                                    row.Cells[5].MergeRight = 2;

                                    row = table.AddRow();
                                    totalRows++;
                                    blockRows++;
                                    row.HeadingFormat = false;
                                    row.Format.Alignment = ParagraphAlignment.Left;
                                    row.Format.Font.Size = "7";
                                    row.Cells[0].AddParagraph("BPreference");
                                    row.Cells[0].Format.Font.Bold = true;
                                    row.Cells[1].AddParagraph(thisTechReport.bpReference.ToString());
                                    row.Cells[1].Format.Font.Bold = false;
                                    row.Cells[1].MergeRight = 1;
                                    row.Cells[4].AddParagraph("Product");
                                    row.Cells[4].Format.Font.Bold = true;
                                    if (thisTechReport.prodDesc == null)
                                    {
                                        row.Cells[5].AddParagraph("");
                                    }
                                    else
                                    {
                                        row.Cells[5].AddParagraph(thisTechReport.prodDesc);
                                    }
                                    row.Cells[5].Format.Font.Bold = false;
                                    row.Cells[5].MergeRight = 2;

                                    row = table.AddRow();
                                    totalRows++;
                                    blockRows++;
                                    row.HeadingFormat = false;
                                    row.Format.Alignment = ParagraphAlignment.Left;
                                    row.Format.Font.Size = "7";
                                    row.Cells[0].AddParagraph("Cluster");
                                    row.Cells[0].Format.Font.Bold = true;
                                    row.Cells[1].AddParagraph(thisTechReport.cluster);
                                    row.Cells[1].Format.Font.Bold = false;
                                    row.Cells[1].MergeRight = 1;
                                    row.Cells[4].AddParagraph("Serialiseditem");
                                    row.Cells[4].Format.Font.Bold = true;
                                    row.Cells[5].AddParagraph(thisTechReport.item);
                                    row.Cells[5].Format.Font.Bold = false;
                                    row.Cells[5].MergeRight = 2;

                                    row = table.AddRow();
                                    totalRows++;
                                    blockRows++;
                                    row.HeadingFormat = false;
                                    row.Format.Alignment = ParagraphAlignment.Left;
                                    row.Format.Font.Size = "7";
                                    row.Cells[0].AddParagraph("Location Adress");
                                    row.Cells[0].Format.Font.Bold = true;
                                    row.Cells[0].MergeDown = 4;
                                    row.Cells[1].AddParagraph(thisTechReport.adres1);
                                    row.Cells[1].Format.Font.Bold = false;
                                    row.Cells[1].MergeRight = 1;
                                    row.Cells[4].AddParagraph("Serial Number");
                                    row.Cells[4].Format.Font.Bold = true;
                                    row.Cells[5].AddParagraph(thisTechReport.serialNumber);
                                    row.Cells[5].Format.Font.Bold = false;
                                    row.Cells[5].MergeRight = 2;

                                    row = table.AddRow();
                                    totalRows++;
                                    blockRows++;
                                    row.HeadingFormat = false;
                                    row.Format.Alignment = ParagraphAlignment.Left;
                                    row.Format.Font.Size = "7";
                                    row.Cells[1].AddParagraph(thisTechReport.adres2);
                                    row.Cells[1].Format.Font.Bold = false;
                                    row.Cells[1].MergeRight = 1;
                                    row.Cells[4].AddParagraph("Warranty until");
                                    row.Cells[4].Format.Font.Bold = true;
                                    row.Cells[4].MergeDown = 1;
                                    row.Cells[5].AddParagraph(thisTechReport.garantie.ToString());
                                    row.Cells[5].Format.Font.Bold = false;
                                    row.Cells[5].MergeRight = 2;
                                    row.Cells[5].MergeDown = 1;

                                    row = table.AddRow();
                                    totalRows++;
                                    blockRows++;
                                    row.HeadingFormat = false;
                                    row.Format.Alignment = ParagraphAlignment.Left;
                                    row.Format.Font.Size = "7";
                                    row.Cells[1].AddParagraph(thisTechReport.adres3);
                                    row.Cells[1].Format.Font.Bold = false;
                                    row.Cells[1].MergeRight = 1;

                                    row = table.AddRow();
                                    totalRows++;
                                    blockRows++;
                                    row.HeadingFormat = false;
                                    row.Format.Alignment = ParagraphAlignment.Left;
                                    row.Format.Font.Size = "7";
                                    row.Cells[1].AddParagraph(thisTechReport.adres4);
                                    row.Cells[1].Format.Font.Bold = false;
                                    row.Cells[1].MergeRight = 1;
                                    row.Cells[4].AddParagraph("");
                                    row.Cells[4].Format.Font.Bold = true;
                                    row.Cells[5].AddParagraph("");
                                    row.Cells[5].Format.Font.Bold = false;
                                    row.Cells[5].MergeRight = 2;

                                    row = table.AddRow();
                                    totalRows++;
                                    blockRows++;
                                    row.HeadingFormat = false;
                                    row.Format.Alignment = ParagraphAlignment.Left;
                                    row.Format.Font.Size = "7";
                                    row.Cells[1].AddParagraph(thisTechReport.adres5);
                                    row.Cells[1].Format.Font.Bold = false;
                                    row.Cells[1].MergeRight = 1;
                                    row.Cells[4].AddParagraph("");
                                    row.Cells[4].Format.Font.Bold = true;
                                    row.Cells[5].AddParagraph();
                                    row.Cells[5].Format.Font.Bold = false;
                                    row.Cells[5].MergeRight = 2;

                                    row = table.AddRow();
                                    totalRows++;
                                    blockRows++;
                                    row.HeadingFormat = false;
                                    row.Format.Alignment = ParagraphAlignment.Left;
                                    row.Format.Font.Size = "7";
                                    row.Cells[0].AddParagraph("Telephone");
                                    row.Cells[0].Format.Font.Bold = true;
                                    row.Cells[1].AddParagraph(thisTechReport.tel);
                                    row.Cells[1].Format.Font.Bold = false;
                                    row.Cells[1].MergeRight = 1;
                                    row.Cells[4].AddParagraph("");
                                    row.Cells[4].Format.Font.Bold = true;
                                    row.Cells[4].MergeDown = 1;
                                    row.Cells[5].AddParagraph("");
                                    row.Cells[5].Format.Font.Bold = false;
                                    row.Cells[5].MergeRight = 2;
                                    row.Cells[5].MergeDown = 1;

                                    row = table.AddRow();
                                    totalRows++;
                                    blockRows++;
                                    row.HeadingFormat = false;
                                    row.Format.Alignment = ParagraphAlignment.Left;
                                    row.Format.Font.Size = "7";
                                    row.Cells[0].AddParagraph("Contact");
                                    row.Cells[0].Format.Font.Bold = true;
                                    row.Cells[1].AddParagraph(thisTechReport.contact);
                                    row.Cells[1].Format.Font.Bold = false;
                                    row.Cells[1].MergeRight = 1;
                                    row.Cells[4].MergeRight = 3;
                                    row.Cells[4].MergeDown = 2;

                                    table.SetEdge(0, totalRows - blockRows, 8, 1, Edge.Box, MigraDoc.DocumentObjectModel.BorderStyle.Single, 0.75, Color.Empty);
                                    table.SetEdge(0, (totalRows - blockRows) + 1, 8, blockRows - 1, Edge.Box, MigraDoc.DocumentObjectModel.BorderStyle.Single, 0.75, Color.Empty);

                                    row = table.AddRow();
                                    totalRows++;

                                    #endregion
                                }

                                #region Serviceorder Header
                                blockRows = 0;
                                soRows = 0;

                                row = table.AddRow();
                                totalRows++;
                                blockRows++;
                                soRows++;
                                row.Height = "0.5cm";
                                row.HeadingFormat = true;
                                row.HeadingFormat = false;
                                row.Format.Alignment = ParagraphAlignment.Center;
                                row.Format.Font.Bold = true;
                                row.Shading.Color = TableGray;
                                row.Cells[0].AddParagraph(string.Format("SERVICEORDER: {0}", _serviceOrder.ToString()));
                                row.Cells[0].Format.Font.Bold = true;
                                row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
                                row.Cells[0].VerticalAlignment = VerticalAlignment.Center;
                                row.Cells[0].MergeRight = 2;
                                row.Cells[4].AddParagraph("");
                                row.Cells[4].Format.Alignment = ParagraphAlignment.Left;
                                row.Cells[4].VerticalAlignment = VerticalAlignment.Center;
                                row.Cells[4].MergeRight = 3;

                                row = table.AddRow();
                                totalRows++;
                                blockRows++;
                                soRows++;
                                row.HeadingFormat = false;
                                row.Format.Alignment = ParagraphAlignment.Left;
                                row.Format.Font.Size = "7";
                                row.Cells[0].AddParagraph("Serv. Department");
                                row.Cells[0].Format.Font.Bold = true;
                                row.Cells[1].AddParagraph(thisTechReport.department);
                                row.Cells[1].Format.Font.Bold = false;
                                row.Cells[1].MergeRight = 1;
                                row.Cells[4].AddParagraph("Activityline");
                                row.Cells[4].Format.Font.Bold = true;
                                row.Cells[5].AddParagraph(thisTechReport.activityLine);
                                row.Cells[5].Format.Font.Bold = false;
                                row.Cells[5].MergeRight = 2;

                                row = table.AddRow();
                                totalRows++;
                                blockRows++;
                                soRows++;
                                row.HeadingFormat = false;
                                row.Format.Alignment = ParagraphAlignment.Left;
                                row.Format.Font.Size = "7";
                                row.Cells[0].AddParagraph("Service Engineer");
                                row.Cells[0].Format.Font.Bold = true;
                                row.Cells[1].AddParagraph(thisTechReport.fseNr.ToString());
                                row.Cells[1].Format.Font.Bold = false;
                                row.Cells[2].AddParagraph(thisTechReport.fseNaam);
                                row.Cells[2].Format.Font.Bold = false;
                                row.Cells[4].AddParagraph("Call");
                                row.Cells[4].Format.Font.Bold = true;
                                row.Cells[5].AddParagraph(thisTechReport.callNumber);
                                row.Cells[5].Format.Font.Bold = false;
                                row.Cells[5].MergeRight = 2;

                                row = table.AddRow();
                                totalRows++;
                                blockRows++;
                                soRows++;
                                row.HeadingFormat = false;
                                row.Format.Alignment = ParagraphAlignment.Left;
                                row.Format.Font.Size = "7";
                                row.Cells[0].Format.Font.Bold = true;
                                row.Cells[1].Format.Font.Bold = false;
                                try
                                {
                                    row.Cells[0].AddParagraph("Report signed");
                                    row.Cells[1].AddParagraph(string.Format("{0} by {1}", ((DateTime)thisTechReport.DateSigned).ToShortDateString(), thisTechReport.SignedBy));
                                }
                                catch
                                {
                                    if (thisTechReport.EarliestStartTime != null)
                                    {
                                        row.Cells[0].AddParagraph("Date");
                                        row.Cells[1].AddParagraph(((DateTime)thisTechReport.EarliestStartTime).ToShortDateString());
                                    }
                                }
                                row.Cells[1].MergeRight = 2;
                                row.Cells[4].AddParagraph("Call Text");
                                row.Cells[4].Format.Font.Bold = true;
                                row.Cells[5].AddParagraph(thisTechReport.call);
                                row.Cells[5].Format.Font.Bold = false;
                                row.Cells[5].MergeRight = 2;

                                try
                                {
                                    table.SetEdge(0, totalRows - blockRows, 8, 1, Edge.Box, MigraDoc.DocumentObjectModel.BorderStyle.Single, 0.50, Color.Empty);
                                    table.SetEdge(0, (totalRows - blockRows) + 1, 8, blockRows - 1, Edge.Box, MigraDoc.DocumentObjectModel.BorderStyle.Single, 0.50, Color.Empty);
                                }
                                catch
                                {

                                }
                                row = table.AddRow();
                                totalRows++;
                                soRows++;

                                #endregion

                                string typeCode = "";

                                var thisServiceOrderCalculateDetails = (from s in myInterventionsData.ServiceorderCalculateDetails
                                                                        join h in myInterventionsData.HourGroups on s.HourGroup equals h.ID into hc
                                                                        from h1 in hc.DefaultIfEmpty()
                                                                        join m in myInterventionsData.Invoicing on s.Invoicing equals m.ID into mc
                                                                        from m1 in mc.DefaultIfEmpty()
                                                                        join t in myInterventionsData.TaskGroups on s.TaskGroup equals t.ID into tc
                                                                        from t1 in tc.DefaultIfEmpty()
                                                                        join r in myInterventionsData.Traveling on s.Traveling equals r.ID into rc
                                                                        from r1 in rc.DefaultIfEmpty()
                                                                        where s.Serviceorder == _serviceOrder
                                                                        select new
                                                                        {
                                                                            hourGroup = h1.Description,
                                                                            invoicing = m1.Description,
                                                                            taskGroup = t1.Description,
                                                                            traveling = r1.Detail,
                                                                            callOutValue = s.CalloutValue,
                                                                            callOutText = s.CallOutText,
                                                                            kmCharge = s.KmCharge,
                                                                            numberOfCallouts = s.NumberOfCallouts,
                                                                            otherText = s.OtherText,
                                                                            type = s.Type,
                                                                            doNotChangeAnymore = s.DoNotChangeAnymore,
                                                                            numberOfKm = s.numberOfKm,
                                                                            travelTime = s.TravelTime

                                                                        }).FirstOrDefault();
                                if (thisServiceOrderCalculateDetails != null)
                                {
                                    typeCode = thisServiceOrderCalculateDetails.type.ToString();
                                }
                                var thisType = (from c in myERPLNData.TypeCode
                                                where c.Type == typeCode
                                                select c).FirstOrDefault();
                                if (thisType != null)
                                {
                                    string _typeCode = string.Format("{0} {1}", typeCode, thisTechReport.srvTypeOmschrijving);
                                    typeCode = _typeCode;
                                }

                                #region Status

                                blockRows = 0;

                                row = table.AddRow();
                                totalRows++;
                                blockRows++;
                                soRows++;
                                row.Height = "0.5cm";
                                row.HeadingFormat = true;
                                row.Format.Alignment = ParagraphAlignment.Center;
                                row.Format.Font.Bold = true;
                                row.Shading.Color = TableGray;
                                if (typeCode != "")
                                {
                                    row.Cells[0].AddParagraph(typeCode);
                                }
                                else
                                {
                                    row.Cells[0].AddParagraph(string.Format("{0} {1}", thisTechReport.srvType.ToString(), thisTechReport.srvTypeOmschrijving));
                                }
                                row.Cells[0].Format.Font.Bold = true;
                                row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
                                row.Cells[0].VerticalAlignment = VerticalAlignment.Center;
                                row.Cells[0].MergeRight = 7;

                                row = table.AddRow();
                                totalRows++;
                                blockRows++;
                                soRows++;
                                row.HeadingFormat = false;
                                row.Format.Alignment = ParagraphAlignment.Center;
                                row.Format.Font.Bold = false;
                                row.Cells[0].AddParagraph("STATUS:");
                                row.Cells[0].Format.Font.Bold = true;
                                row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
                                row.Cells[0].VerticalAlignment = VerticalAlignment.Top;
                                row.Cells[0].MergeRight = 1;
                                row.Cells[2].AddParagraph(thisTechReport.status);
                                row.Cells[2].Format.Font.Bold = false;
                                row.Cells[2].Format.Alignment = ParagraphAlignment.Left;
                                row.Cells[2].VerticalAlignment = VerticalAlignment.Top;
                                row.Cells[2].MergeRight = 5;

                                if (!thisTechReport.problemSolved)
                                {
                                    row = table.AddRow();
                                    totalRows++;
                                    blockRows++;
                                    soRows++;
                                    row.HeadingFormat = false;
                                    row.Format.Alignment = ParagraphAlignment.Center;
                                    row.Format.Font.Bold = false;
                                    row.Cells[0].AddParagraph("NOT SOLVED:");
                                    row.Cells[0].Format.Font.Bold = true;
                                    row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
                                    row.Cells[0].VerticalAlignment = VerticalAlignment.Top;
                                    row.Cells[0].MergeRight = 1;
                                    row.Cells[2].AddParagraph(thisTechReport.problemNotSolved);
                                    row.Cells[2].Format.Font.Bold = false;
                                    row.Cells[2].Format.Alignment = ParagraphAlignment.Left;
                                    row.Cells[2].VerticalAlignment = VerticalAlignment.Top;
                                    row.Cells[2].MergeRight = 5;

                                }
                                if (thisTechReport.safetyNOK)
                                {
                                    row = table.AddRow();
                                    totalRows++;
                                    blockRows++;
                                    soRows++;
                                    row.HeadingFormat = false;
                                    row.Format.Alignment = ParagraphAlignment.Center;
                                    row.Format.Font.Bold = false;
                                    row.Cells[0].AddParagraph("SAFETIES NOT OK:");
                                    row.Cells[0].Format.Font.Bold = true;
                                    row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
                                    row.Cells[0].VerticalAlignment = VerticalAlignment.Top;
                                    row.Cells[0].MergeRight = 1;
                                    row.Cells[2].AddParagraph(thisTechReport.safetyNOKDetails);
                                    row.Cells[2].Format.Font.Bold = false;
                                    row.Cells[2].Format.Alignment = ParagraphAlignment.Left;
                                    row.Cells[2].VerticalAlignment = VerticalAlignment.Top;
                                    row.Cells[2].MergeRight = 5;


                                }

                                try
                                {
                                    table.SetEdge(0, totalRows - blockRows, 8, 1, Edge.Box, MigraDoc.DocumentObjectModel.BorderStyle.Single, 0.50, Color.Empty);
                                    table.SetEdge(0, (totalRows - blockRows) + 1, 8, blockRows - 1, Edge.Box, MigraDoc.DocumentObjectModel.BorderStyle.Single, 0.50, Color.Empty);
                                }
                                catch (Exception)
                                {


                                }
                                row.HeadingFormat = false;
                                #endregion

                                #region Problem
                                blockRows = 0;

                                row = table.AddRow();
                                totalRows++;
                                blockRows++;
                                soRows++;
                                row.Height = "0.5cm";
                                row.HeadingFormat = true;
                                row.Format.Alignment = ParagraphAlignment.Center;
                                row.Format.Font.Bold = true;
                                row.Shading.Color = TableGray;
                                row.Cells[0].AddParagraph("PROBLEM DESCRIPTION");
                                row.Cells[0].Format.Font.Bold = true;
                                row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
                                row.Cells[0].VerticalAlignment = VerticalAlignment.Center;
                                row.Cells[0].MergeRight = 7;

                                List<Int32> linesProblem = splitTextInSectionsLineCount(thisTechReport.problemDetail, 60 - totalRows, 60);
                                List<string> sectiesProblem = splitTextInSections(thisTechReport.problemDetail, 60 - totalRows, 60);

                                Int32 i = 0;
                                foreach (var item in sectiesProblem)
                                {
                                    row = table.AddRow();
                                    totalRows++;
                                    blockRows++;
                                    soRows++;
                                    i++;
                                    row.HeadingFormat = false;
                                    row.Format.Font.Bold = false;
                                    row.Cells[0].AddParagraph(item);
                                    row.Cells[0].Format.Font.Bold = false;
                                    row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
                                    row.Cells[0].VerticalAlignment = VerticalAlignment.Top;
                                    row.Cells[0].MergeRight = 7;
                                    Int32 rowHeight = (Int32)row.Height;
                                }

                                try
                                {
                                    table.SetEdge(0, totalRows - blockRows, 8, 1, Edge.Box, MigraDoc.DocumentObjectModel.BorderStyle.Single, 0.50, Color.Empty);
                                    table.SetEdge(0, (totalRows - blockRows) + 1, 8, blockRows - 1, Edge.Box, MigraDoc.DocumentObjectModel.BorderStyle.Single, 0.50, Color.Empty);
                                }
                                catch (Exception)
                                {


                                }
                                row.Borders.Visible = false;
                                #endregion

                                #region Service Report
                                if (customer)
                                {
                                    blockRows = 0;

                                    row = table.AddRow();
                                    totalRows++;
                                    blockRows++;
                                    soRows++;
                                    row.Height = "0.5cm";
                                    row.HeadingFormat = true;
                                    row.Format.Alignment = ParagraphAlignment.Center;
                                    row.Format.Font.Bold = true;
                                    row.Shading.Color = TableGray;
                                    row.Cells[0].AddParagraph("SERVICE REPORT");
                                    row.Cells[0].Format.Font.Bold = true;
                                    row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
                                    row.Cells[0].VerticalAlignment = VerticalAlignment.Center;
                                    row.Cells[0].MergeRight = 7;
                                    List<Int32> linesTechnicalReport = splitTextInSectionsLineCount(thisTechReport.action, 60 - totalRows, 60);
                                    List<string> sectiesTechnicalReport = splitTextInSections(thisTechReport.action, 60, 60);
                                    i = 0;
                                    foreach (var item in sectiesTechnicalReport)
                                    {
                                        row = table.AddRow();
                                        totalRows++;
                                        blockRows++;
                                        soRows++;
                                        i++;
                                        row.HeadingFormat = false;
                                        row.Format.Font.Bold = false;
                                        row.Cells[0].AddParagraph(item);
                                        row.Cells[0].Format.Font.Bold = false;
                                        row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
                                        row.Cells[0].VerticalAlignment = VerticalAlignment.Top;
                                        row.Cells[0].MergeRight = 7;
                                        Int32 rowHeight = (Int32)row.Height;
                                    }

                                    try
                                    {
                                        table.SetEdge(0, totalRows - blockRows, 8, 1, Edge.Box, MigraDoc.DocumentObjectModel.BorderStyle.Single, 0.50, Color.Empty);
                                        table.SetEdge(0, (totalRows - blockRows) + 1, 8, blockRows - 1, Edge.Box, MigraDoc.DocumentObjectModel.BorderStyle.Single, 0.50, Color.Empty);
                                    }
                                    catch (Exception)
                                    {


                                    }

                                    row.Borders.Visible = false;
                                }
                                #endregion

                                #region technical Report
                                if (!customer)
                                {
                                    blockRows = 0;

                                    row = table.AddRow();
                                    totalRows++;
                                    blockRows++;
                                    soRows++;
                                    row.Height = "0.5cm";
                                    row.HeadingFormat = true;
                                    row.Format.Alignment = ParagraphAlignment.Center;
                                    row.Format.Font.Bold = true;
                                    row.Shading.Color = TableGray;
                                    row.Cells[0].AddParagraph("TECHNICAL REPORT");
                                    row.Cells[0].Format.Font.Bold = true;
                                    row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
                                    row.Cells[0].VerticalAlignment = VerticalAlignment.Center;
                                    row.Cells[0].MergeRight = 7;
                                    List<Int32> linesReport = splitTextInSectionsLineCount(thisTechReport.techReport, 60 - totalRows, 60);
                                    List<string> sectiesReport = splitTextInSections(thisTechReport.techReport, 60 - totalRows, 60);
                                    i = 0;
                                    foreach (var item in sectiesReport)
                                    {
                                        row = table.AddRow();
                                        totalRows++;
                                        blockRows++;
                                        soRows++;
                                        i++;
                                        row.HeadingFormat = false;
                                        row.Format.Font.Bold = false;
                                        row.Cells[0].AddParagraph(item);
                                        row.Cells[0].Format.Font.Bold = false;
                                        row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
                                        row.Cells[0].VerticalAlignment = VerticalAlignment.Top;
                                        row.Cells[0].MergeRight = 7;
                                        Int32 rowHeight = (Int32)row.Height;
                                    }

                                    try
                                    {
                                        table.SetEdge(0, totalRows - blockRows, 8, 1, Edge.Box, MigraDoc.DocumentObjectModel.BorderStyle.Single, 0.50, Color.Empty);
                                        table.SetEdge(0, (totalRows - blockRows) + 1, 8, blockRows - 1, Edge.Box, MigraDoc.DocumentObjectModel.BorderStyle.Single, 0.50, Color.Empty);
                                    }
                                    catch (Exception)
                                    {


                                    }
                                    row.Borders.Visible = false;
                                }
                                #endregion

                                #region Requests
                                blockRows = 0;

                                //nbrUsed = 0;
                                Int32 nbrReturned = 0;
                                row = table.AddRow();
                                totalRows++;
                                blockRows++;
                                soRows++;
                                row.HeadingFormat = true;
                                nbrReturned++;

                                if (thisTechReport.requests)
                                {
                                    row = table.AddRow();
                                    totalRows++;
                                    blockRows++;
                                    soRows++;
                                    row.Height = "0.5cm";
                                    row.HeadingFormat = true;
                                    row.Format.Alignment = ParagraphAlignment.Center;
                                    row.Format.Font.Bold = true;
                                    row.Shading.Color = TableGray;
                                    row.Cells[0].AddParagraph("Requests".ToUpper());
                                    row.Cells[0].Format.Font.Bold = true;
                                    row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
                                    row.Cells[0].VerticalAlignment = VerticalAlignment.Center;
                                    row.Cells[0].MergeRight = 7;

                                    row = table.AddRow();
                                    totalRows++;
                                    blockRows++;
                                    soRows++;
                                    row.HeadingFormat = false;
                                    row.Format.Alignment = ParagraphAlignment.Center;
                                    row.Format.Font.Bold = false;
                                    row.Cells[0].AddParagraph(thisTechReport.requestsDetails);
                                    row.Cells[0].Format.Font.Bold = false;
                                    row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
                                    row.Cells[0].VerticalAlignment = VerticalAlignment.Top;
                                    row.Cells[0].MergeRight = 7;

                                    try
                                    {
                                        table.SetEdge(0, totalRows - blockRows, 8, 1, Edge.Box, MigraDoc.DocumentObjectModel.BorderStyle.Single, 0.50, Color.Empty);
                                        table.SetEdge(0, (totalRows - blockRows) + 1, 8, blockRows - 1, Edge.Box, MigraDoc.DocumentObjectModel.BorderStyle.Single, 0.50, Color.Empty);
                                    }
                                    catch (Exception)
                                    {


                                    }
                                    row.HeadingFormat = true;
                                }
                                #endregion

                                if (customer)
                                {

                                    #region Hours
                                    blockRows = 0;

                                    row.Borders.Visible = false;

                                    var thisOrderHourDetails = from s in myInterventionsData.DayDetails
                                                               join a in myInterventionsData.ActivityTypes on s.ActivityType equals a.ID into ac
                                                               from a1 in ac.DefaultIfEmpty()
                                                               join f in myInterventionsData.TravelLocations on s.TravelFrom equals f.ID into fc
                                                               from f1 in fc.DefaultIfEmpty()
                                                               join t in myInterventionsData.TravelLocations on s.TravelTo equals t.ID into tc
                                                               from t1 in tc.DefaultIfEmpty()
                                                               join o in myInterventionsData.ServiceOrder on s.ServiceOrder equals o.SrvOrder into oc
                                                               from o1 in oc.DefaultIfEmpty()
                                                               orderby s.DateOfDay, s.StartTime
                                                               where s.ServiceOrder == _serviceOrder
                                                               select new
                                                               {
                                                                   Date = s.DateOfDay,
                                                                   Start = s.StartTime,
                                                                   Stop = s.StopTime,
                                                                   Activity = a1.DetailEnglish,
                                                                   TravelFrom = f1.DetailEnglish,
                                                                   TravelTo = t1.DetailEnglish,
                                                                   Distance = s.Distance,

                                                               };

                                    if (thisOrderHourDetails != null && thisOrderHourDetails.Count() > 0)
                                    {
                                        //nbrUsed = 2;
                                        row = table.AddRow();
                                        totalRows++;
                                        blockRows++;
                                        soRows++;
                                        row.Height = "0.5cm";
                                        row.HeadingFormat = true;
                                        row.Format.Alignment = ParagraphAlignment.Center;
                                        row.Format.Font.Bold = true;
                                        row.Shading.Color = TableGray;
                                        row.Cells[0].AddParagraph("HOURS");
                                        row.Cells[0].Format.Font.Bold = true;
                                        row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
                                        row.Cells[0].VerticalAlignment = VerticalAlignment.Center;
                                        row.Cells[0].MergeRight = 7;

                                        row = table.AddRow();
                                        totalRows++;
                                        blockRows++;
                                        soRows++;
                                        row.HeadingFormat = true;
                                        row.Format.Alignment = ParagraphAlignment.Center;
                                        row.Format.Font.Bold = true;
                                        row.Cells[0].AddParagraph("Date");
                                        row.Cells[0].Format.Font.Bold = true;
                                        row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
                                        row.Cells[0].VerticalAlignment = VerticalAlignment.Center;
                                        row.Cells[1].AddParagraph("Activity");
                                        row.Cells[1].Format.Font.Bold = true;
                                        row.Cells[1].Format.Alignment = ParagraphAlignment.Left;
                                        row.Cells[1].VerticalAlignment = VerticalAlignment.Center;
                                        row.Cells[1].MergeRight = 1;
                                        row.Cells[3].AddParagraph("Location");
                                        row.Cells[3].Format.Font.Bold = true;
                                        row.Cells[3].Format.Alignment = ParagraphAlignment.Left;
                                        row.Cells[3].VerticalAlignment = VerticalAlignment.Center;
                                        row.Cells[3].MergeRight = 2;
                                        row.Cells[6].AddParagraph("Start");
                                        row.Cells[6].Format.Font.Bold = true;
                                        row.Cells[6].Format.Alignment = ParagraphAlignment.Center;
                                        row.Cells[6].VerticalAlignment = VerticalAlignment.Center;
                                        row.Cells[7].AddParagraph("Stop");
                                        row.Cells[7].Format.Font.Bold = true;
                                        row.Cells[7].Format.Alignment = ParagraphAlignment.Center;
                                        row.Cells[7].VerticalAlignment = VerticalAlignment.Center;


                                        foreach (var item in thisOrderHourDetails)
                                        {
                                            //nbrUsed++;
                                            row = table.AddRow();
                                            totalRows++;
                                            blockRows++;
                                            soRows++;
                                            row.HeadingFormat = true;
                                            row.Format.Alignment = ParagraphAlignment.Center;
                                            row.Format.Font.Bold = false;
                                            row.Cells[0].AddParagraph(item.Date.ToString("d"));
                                            row.Cells[0].Format.Font.Bold = false;
                                            row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
                                            row.Cells[0].VerticalAlignment = VerticalAlignment.Center;
                                            row.Cells[1].AddParagraph(item.Activity);
                                            row.Cells[1].Format.Font.Bold = false;
                                            row.Cells[1].Format.Alignment = ParagraphAlignment.Left;
                                            row.Cells[1].VerticalAlignment = VerticalAlignment.Center;
                                            row.Cells[1].MergeRight = 1;
                                            if (string.IsNullOrEmpty(item.TravelTo))
                                            {
                                                row.Cells[3].AddParagraph(string.Format("{0}", item.TravelFrom));
                                            }
                                            else
                                            {
                                                row.Cells[3].AddParagraph(string.Format("{0}-{1}", item.TravelFrom, item.TravelTo));
                                            }
                                            row.Cells[3].Format.Font.Bold = false;
                                            row.Cells[3].Format.Alignment = ParagraphAlignment.Left;
                                            row.Cells[3].VerticalAlignment = VerticalAlignment.Center;
                                            row.Cells[3].MergeRight = 2;
                                            row.Cells[6].AddParagraph(item.Start);
                                            row.Cells[6].Format.Font.Bold = false;
                                            row.Cells[6].Format.Alignment = ParagraphAlignment.Center;
                                            row.Cells[6].VerticalAlignment = VerticalAlignment.Center;
                                            row.Cells[7].AddParagraph(item.Stop);
                                            row.Cells[7].Format.Font.Bold = false;
                                            row.Cells[7].Format.Alignment = ParagraphAlignment.Center;
                                            row.Cells[7].VerticalAlignment = VerticalAlignment.Center;

                                        }

                                        try
                                        {
                                            table.SetEdge(0, totalRows - blockRows, 8, 1, Edge.Box, MigraDoc.DocumentObjectModel.BorderStyle.Single, 0.50, Color.Empty);
                                            table.SetEdge(0, (totalRows - blockRows) + 1, 8, blockRows - 1, Edge.Box, MigraDoc.DocumentObjectModel.BorderStyle.Single, 0.50, Color.Empty);
                                        }
                                        catch (Exception)
                                        {


                                        }
                                        row.HeadingFormat = true;
                                    }
                                    #endregion

                                    #region Costs
                                    blockRows = 0;

                                    var costDetails = from s in myInterventionsData.AssignedCosts
                                                      join t in myInterventionsData.CostDetails on s.CostIDString equals t.ImportIDString into tc
                                                      from t1 in tc.DefaultIfEmpty()
                                                      join r in myInterventionsData.CostCategories on t1.CostCategory equals r.ID into rc
                                                      from r1 in rc.DefaultIfEmpty()
                                                      join c in myInterventionsData.Currency on t1.Currency equals c.ID into cc
                                                      from c1 in cc.DefaultIfEmpty()
                                                      join p in myInterventionsData.PaymentMethods on t1.PaymentMethod equals p.ID into pc
                                                      from p1 in pc.DefaultIfEmpty()
                                                      where s.ServiceOrderID == _serviceOrder
                                                      select new
                                                      {
                                                          costCategory = r1.DetailEnglish,
                                                          detail = t1.Detail,
                                                          amount = t1.Amount,
                                                          currency = c1.Currency1,
                                                          loadPercentage = s.LoadPercentage,
                                                          dateFrom = t1.DateFrom,
                                                          dateTo = t1.DateTo,
                                                          includeExclude = s.IncludeExclude
                                                      };
                                    if (costDetails != null && costDetails.Count() > 0)
                                    {
                                        nbrReturned = 2;
                                        row = table.AddRow();
                                        totalRows++;
                                        blockRows++;
                                        soRows++;
                                        row.Height = "0.5cm";
                                        row.HeadingFormat = true;
                                        row.Format.Alignment = ParagraphAlignment.Center;
                                        row.Format.Font.Bold = true;
                                        row.Shading.Color = TableGray;
                                        row.Cells[0].AddParagraph("COSTS");
                                        row.Cells[0].Format.Font.Bold = true;
                                        row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
                                        row.Cells[0].VerticalAlignment = VerticalAlignment.Center;
                                        row.Cells[0].MergeRight = 7;

                                        row = table.AddRow();
                                        totalRows++;
                                        blockRows++;
                                        soRows++;
                                        row.HeadingFormat = true;

                                        row.Format.Alignment = ParagraphAlignment.Center;
                                        row.Format.Font.Bold = true;
                                        row.Cells[0].AddParagraph("Included");
                                        row.Cells[0].Format.Font.Bold = true;
                                        row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
                                        row.Cells[0].VerticalAlignment = VerticalAlignment.Center;
                                        row.Cells[1].AddParagraph("Cost");
                                        row.Cells[1].Format.Font.Bold = true;
                                        row.Cells[1].Format.Alignment = ParagraphAlignment.Left;
                                        row.Cells[1].VerticalAlignment = VerticalAlignment.Center;
                                        row.Cells[1].MergeRight = 1;
                                        row.Cells[3].AddParagraph("Amount");
                                        row.Cells[3].Format.Font.Bold = true;
                                        row.Cells[3].Format.Alignment = ParagraphAlignment.Left;
                                        row.Cells[3].VerticalAlignment = VerticalAlignment.Center;
                                        row.Cells[3].MergeRight = 2;
                                        row.Cells[6].AddParagraph("From");
                                        row.Cells[6].Format.Font.Bold = true;
                                        row.Cells[6].Format.Alignment = ParagraphAlignment.Center;
                                        row.Cells[6].VerticalAlignment = VerticalAlignment.Center;
                                        row.Cells[7].AddParagraph("To");
                                        row.Cells[7].Format.Font.Bold = true;
                                        row.Cells[7].Format.Alignment = ParagraphAlignment.Center;
                                        row.Cells[7].VerticalAlignment = VerticalAlignment.Center;

                                        foreach (var item in costDetails)
                                        {
                                            nbrReturned++;
                                            row = table.AddRow();
                                            totalRows++;
                                            blockRows++;
                                            soRows++;
                                            row.HeadingFormat = true;
                                            row.Format.Alignment = ParagraphAlignment.Center;
                                            row.Format.Font.Bold = false;

                                            if (item.includeExclude.ToUpper() == "INCLUDED")
                                            {
                                                row.Cells[0].AddParagraph("YES");
                                            }
                                            else
                                            {
                                                row.Cells[0].AddParagraph("NO");
                                            }
                                            row.Cells[0].Format.Font.Bold = false;
                                            row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
                                            row.Cells[0].VerticalAlignment = VerticalAlignment.Center;
                                            if (item.costCategory != null)
                                            {
                                                row.Cells[1].AddParagraph(item.costCategory);
                                            }
                                            else
                                            {
                                                row.Cells[1].AddParagraph("ERROR: No CostCategory");
                                            }
                                            row.Cells[1].Format.Font.Bold = false;
                                            row.Cells[1].Format.Alignment = ParagraphAlignment.Left;
                                            row.Cells[1].VerticalAlignment = VerticalAlignment.Center;
                                            row.Cells[1].MergeRight = 1;
                                            row.Cells[3].AddParagraph(string.Format("{0} {1} ({2}%)", item.amount.ToString(), item.currency, item.loadPercentage));
                                            row.Cells[3].Format.Font.Bold = false;
                                            row.Cells[3].Format.Alignment = ParagraphAlignment.Left;
                                            row.Cells[3].VerticalAlignment = VerticalAlignment.Center;
                                            row.Cells[3].MergeRight = 2;
                                            row.Cells[6].AddParagraph(((DateTime)item.dateFrom).ToShortDateString());
                                            row.Cells[6].Format.Font.Bold = false;
                                            row.Cells[6].Format.Alignment = ParagraphAlignment.Center;
                                            row.Cells[6].VerticalAlignment = VerticalAlignment.Center;
                                            row.Cells[7].AddParagraph(((DateTime)item.dateTo).ToShortDateString());
                                            row.Cells[7].Format.Font.Bold = false;
                                            row.Cells[7].Format.Alignment = ParagraphAlignment.Center;
                                            row.Cells[7].VerticalAlignment = VerticalAlignment.Center;

                                        }

                                        try
                                        {
                                            table.SetEdge(0, totalRows - blockRows, 8, 1, Edge.Box, MigraDoc.DocumentObjectModel.BorderStyle.Single, 0.50, Color.Empty);
                                            table.SetEdge(0, (totalRows - blockRows) + 1, 8, blockRows - 1, Edge.Box, MigraDoc.DocumentObjectModel.BorderStyle.Single, 0.50, Color.Empty);
                                        }
                                        catch (Exception)
                                        {


                                        }
                                        row.HeadingFormat = true;
                                    }
                                    #endregion
                                }
                                try
                                {
                                    row = table.AddRow();
                                    totalRows++;
                                    soRows++;
                                    table.SetEdge(0, totalRows - soRows, 8, soRows - 1, Edge.Box, MigraDoc.DocumentObjectModel.BorderStyle.Single, 0.75, Color.Empty);
                                }
                                catch (Exception)
                                {


                                }
                            }

                            #endregion
                        }

                        foreach (var _serviceOrder in partsorders)
                        {
                            bool linesPrinted = false;
                            #region loop
                            var thispartsorder = (from s in myERPLNData.ERPLN_ServiceOrders
                                                  join t in myERPLNData.Srv_Calls on s.Serviceorder equals t.ServiceOrder
                                                  where s.Serviceorder == _serviceOrder
                                                  select new
                                                  {
                                                      Serviceorder = s.Serviceorder,
                                                      Customer = t.SalesName,
                                                      Department = t.ServiceDepartment,
                                                      Type = s.Type + " " + s.TypeDetail,
                                                      Status = s.Status,
                                                      DateOfOrder = s.orderDate
                                                  }).FirstOrDefault();

                            if (thispartsorder != null)
                            {
                                #region orderdetails
                                var partsOrderLines = (from s in myERPLNData.Srv_ServiceorderRegels
                                                       where s.ServiceOrder == _serviceOrder &&
                                                       s.Artikel != "SERTRP" &&
                                                       s.Artikel != "SERPCK" &&
                                                       s.Artikel != "SEROTH"
                                                       orderby s.Regelnummer
                                                       select s).ToList();


                                if (partsOrderLines != null && partsOrderLines.Count() > 0)
                                {
                                    linesPrinted = false;

                                    #region Serviceorder Header
                                    blockRows = 0;
                                    soRows = 0;

                                    row = table.AddRow();
                                    totalRows++;
                                    blockRows++;
                                    soRows++;
                                    row.Height = "0.5cm";
                                    row.HeadingFormat = true;
                                    row.HeadingFormat = false;
                                    row.Format.Alignment = ParagraphAlignment.Center;
                                    row.Format.Font.Bold = true;
                                    row.Shading.Color = TableGray;
                                    row.Cells[0].AddParagraph(string.Format("PARTSORDER: {0}", _serviceOrder.ToString()));
                                    row.Cells[0].Format.Font.Bold = true;
                                    row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
                                    row.Cells[0].VerticalAlignment = VerticalAlignment.Center;
                                    row.Cells[0].MergeRight = 2;
                                    row.Cells[5].AddParagraph("");
                                    row.Cells[5].Format.Alignment = ParagraphAlignment.Left;
                                    row.Cells[5].VerticalAlignment = VerticalAlignment.Center;
                                    //row.Cells[5].MergeRight = 1;
                                    row.Cells[6].AddParagraph(thispartsorder.Department);
                                    row.Cells[6].Format.Alignment = ParagraphAlignment.Left;
                                    row.Cells[6].VerticalAlignment = VerticalAlignment.Center;
                                    row.Cells[6].MergeRight = 1;

                                    row = table.AddRow();
                                    totalRows++;
                                    blockRows++;
                                    soRows++;
                                    row.HeadingFormat = false;
                                    row.Format.Alignment = ParagraphAlignment.Left;
                                    row.Format.Font.Size = "7";
                                    row.Cells[0].AddParagraph(thispartsorder.Type);
                                    row.Cells[0].Format.Font.Bold = false;
                                    row.Cells[0].MergeRight = 2;
                                    try
                                    {
                                        string orderDate = ((DateTime)thispartsorder.DateOfOrder).ToShortDateString();
                                        row.Cells[6].AddParagraph(orderDate);
                                    }
                                    catch (Exception)
                                    {

                                        throw;
                                    }
                                    row.Cells[6].Format.Font.Bold = false;
                                    row.Cells[6].MergeRight = 1;

                                    totalRows++;
                                    soRows++;
                                    blockRows++;

                                    #endregion

                                    row = table.AddRow();
                                    totalRows++;
                                    blockRows++;
                                    soRows++;
                                    row.HeadingFormat = true;
                                    row.Format.Alignment = ParagraphAlignment.Center;
                                    row.Format.Font.Bold = true;
                                    row.Cells[0].AddParagraph("Part");
                                    row.Cells[0].Format.Font.Bold = true;
                                    row.Cells[0].Format.Font.Size = 6;
                                    row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
                                    row.Cells[0].VerticalAlignment = VerticalAlignment.Center;
                                    row.Cells[0].MergeRight = 1;
                                    row.Cells[2].AddParagraph("Details");
                                    row.Cells[2].Format.Font.Bold = true;
                                    row.Cells[2].Format.Font.Size = 6;
                                    row.Cells[2].Format.Alignment = ParagraphAlignment.Left;
                                    row.Cells[2].VerticalAlignment = VerticalAlignment.Center;
                                    row.Cells[2].MergeRight = 2;
                                    row.Cells[5].AddParagraph("Aantal");
                                    row.Cells[5].Format.Font.Bold = true;
                                    row.Cells[5].Format.Font.Size = 6;
                                    row.Cells[5].Format.Alignment = ParagraphAlignment.Right;
                                    row.Cells[5].VerticalAlignment = VerticalAlignment.Center;
                                    row.Cells[6].AddParagraph("Type");
                                    row.Cells[6].Format.Font.Bold = true;
                                    row.Cells[6].Format.Font.Size = 6;
                                    row.Cells[6].Format.Alignment = ParagraphAlignment.Left;
                                    row.Cells[6].VerticalAlignment = VerticalAlignment.Center;
                                    row.Cells[6].MergeRight = 1;

                                    foreach (var item in partsOrderLines)
                                    {
                                        linesPrinted = true;
                                        row = table.AddRow();
                                        totalRows++;
                                        blockRows++;
                                        soRows++;
                                        row.HeadingFormat = true;
                                        row.Format.Alignment = ParagraphAlignment.Center;
                                        row.Format.Font.Bold = false;
                                        row.Cells[0].AddParagraph(string.Format("{0} - {1}", item.Regelnummer.ToString(), item.Artikel));
                                        row.Cells[0].Format.Font.Bold = false;
                                        row.Cells[0].Format.Font.Size = 6;
                                        row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
                                        row.Cells[0].VerticalAlignment = VerticalAlignment.Center;
                                        row.Cells[0].MergeRight = 1;
                                        if (item.ArtikelOmschrijving == null)
                                        {
                                            row.Cells[2].AddParagraph("");
                                        }
                                        else
                                        {
                                            row.Cells[2].AddParagraph(item.ArtikelOmschrijving);
                                        }
                                        row.Cells[2].Format.Font.Bold = false;
                                        row.Cells[2].Format.Font.Size = 6;
                                        row.Cells[2].Format.Alignment = ParagraphAlignment.Left;
                                        row.Cells[2].VerticalAlignment = VerticalAlignment.Center;
                                        row.Cells[2].MergeRight = 2;
                                        row.Cells[5].AddParagraph(item.TotaleHoeveelheid.ToString());
                                        row.Cells[5].Format.Font.Bold = false;
                                        row.Cells[5].Format.Font.Size = 6;
                                        row.Cells[5].Format.Alignment = ParagraphAlignment.Right;
                                        row.Cells[5].VerticalAlignment = VerticalAlignment.Center;
                                        if (item.SoortLevering != null)
                                        {
                                            switch (item.SoortLevering)
                                            {
                                                case "Vanuit magazijn":
                                                    row.Cells[6].AddParagraph("From Warehouse");
                                                    break;
                                                case "Naar magazijn":
                                                    row.Cells[6].AddParagraph("To Warehouse");
                                                    break;
                                                case "Via on-site inkoop":
                                                    row.Cells[6].AddParagraph("No Warehouse");
                                                    break;
                                                default:
                                                    row.Cells[6].AddParagraph(item.SoortLevering);
                                                    break;
                                            }
                                        }
                                        else
                                        {
                                            row.Cells[6].AddParagraph("");
                                        }

                                        row.Cells[6].Format.Font.Bold = false;
                                        row.Cells[6].Format.Font.Size = 6;
                                        row.Cells[6].Format.Alignment = ParagraphAlignment.Left;
                                        row.Cells[6].VerticalAlignment = VerticalAlignment.Center;
                                        row.Cells[6].MergeRight = 1;

                                    }
                                    row.HeadingFormat = true;
                                }

                                try
                                {
                                    table.SetEdge(0, totalRows - blockRows, 8, 1, Edge.Box, MigraDoc.DocumentObjectModel.BorderStyle.Single, 0.50, Color.Empty);
                                    table.SetEdge(0, totalRows - blockRows, 8, 2, Edge.Box, MigraDoc.DocumentObjectModel.BorderStyle.Single, 0.50, Color.Empty);
                                    table.SetEdge(0, (totalRows - blockRows) + 1, 8, blockRows - 2, Edge.Box, MigraDoc.DocumentObjectModel.BorderStyle.Single, 0.50, Color.Empty);

                                }
                                catch (Exception)
                                {


                                }
                                #endregion
                            }
                            if (linesPrinted)
                            {
                                row = table.AddRow();
                                linesPrinted = false;
                            }

                            #endregion
                        }
                        document.UseCmykColor = true;

                        PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(true);
                        pdfRenderer.Document = document;
                        pdfRenderer.RenderDocument();
                        pdfRenderer.Save(fileName);
                    }
                }
            }
            return true;
        }

        private static List<Int32> splitTextInSectionsLineCount(string f, Int32 linesOnFirstPage, Int32 linesOnFullPage)
        {
            List<Int32> secties = new List<Int32>();
            bool eersteSectie = true;
            string newTextToReturn = "";
            Int32 charactersPerLine = 100;
            Int32 lineCount = 1;
            Int32 charCount = 0;
            for (int i = 0; i < f.Length; i++)
            {
                charCount++;
                newTextToReturn = newTextToReturn + f.Substring(i, 1);
                if (f.Substring(i, 1) == "\n" || charCount == charactersPerLine)
                {
                    lineCount++;
                    charCount = 0;
                    switch (eersteSectie)
                    {
                        case true:
                            if (lineCount > linesOnFirstPage)
                            {
                                eersteSectie = false;
                                secties.Add(lineCount);
                                newTextToReturn = "";
                                lineCount = 0;
                                charCount = 0;
                            }
                            break;
                        case false:
                            if (lineCount > linesOnFullPage)
                            {
                                eersteSectie = false;
                                secties.Add(lineCount);
                                newTextToReturn = "";
                                lineCount = 0;
                                charCount = 0;
                            }
                            break;
                        default:
                            break;
                    }
                }
            }
            if (!string.IsNullOrEmpty(newTextToReturn))
            {
                secties.Add(lineCount);
            }
            return secties;
        }

        private static List<string> splitTextInSections(string f, Int32 linesOnFirstPage, Int32 linesOnFullPage)
        {
            List<string> secties = new List<string>();
            bool eersteSectie = true;
            string newTextToReturn = "";
            Int32 charactersPerLine = 100;
            Int32 lineCount = 0;
            Int32 charCount = 0;
            for (int i = 0; i < f.Length; i++)
            {
                charCount++;
                newTextToReturn = newTextToReturn + f.Substring(i, 1);
                if (f.Substring(i, 1) == "\n" || charCount == charactersPerLine)
                {
                    lineCount++;
                    charCount = 0;
                    switch (eersteSectie)
                    {
                        case true:
                            if (lineCount > linesOnFirstPage)
                            {
                                eersteSectie = false;
                                secties.Add(newTextToReturn);
                                newTextToReturn = "";
                                lineCount = 0;
                                charCount = 0;
                            }
                            break;
                        case false:
                            if (lineCount > linesOnFullPage)
                            {
                                eersteSectie = false;
                                secties.Add(newTextToReturn);
                                newTextToReturn = "";
                                lineCount = 0;
                                charCount = 0;
                            }
                            break;
                        default:
                            break;
                    }
                }
            }
            if (!string.IsNullOrEmpty(newTextToReturn))
            {
                secties.Add(newTextToReturn);
            }
            return secties;
        }

        readonly static Color TableBorder = new Color(81, 125, 192);
        readonly static Color TableBlue = new Color(235, 240, 249);
        readonly static Color TableGray = new Color(242, 242, 242);
        readonly static Color TableLightGray = new Color(245, 245, 245);
        readonly static Color TableWhite = new Color(255, 255, 255);
    }
}
