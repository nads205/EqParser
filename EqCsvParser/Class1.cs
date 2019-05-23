using System.Collections.Generic;
using System;
using System.Linq;
using System.Diagnostics;
using System.IO;

namespace EqCsvParser
{
    public class Reader
    {
        const string fileAdHrUsers =                                @"C:\Users\nads205\Desktop\EqG1_Deployment_Schedule_v1.0_work_in_progress.xlsx";
        const string fileAdHrUsersNew =                             @"C:\Users\nads205\Desktop\G1 Deployment Schedule v2.0.xlsx";
        const string fileNew =                                      @"C:\Users\nads205\Desktop\fileNew.csv";
        const string fileUk1UsersAll =                              @"C:\Users\nads205\Desktop\UK1UsersNEW.csv";
        const string fileAdUsers1 =                                 @"C:\Users\nads205\Desktop\EqUK1UsersBusinessAreas.csv";
        const string fileAdUsers2 =                                 @"C:\Users\nads205\Desktop\EqUK1UsersGeneralBusinessUsers.csv";
        const string fileAdUsers3 =                                 @"C:\Users\nads205\Desktop\EqUK1UsersGroupFunctions.csv";        
        const string fileRecordsMissingFromDeploymentSchedule =     @"C:\Users\nads205\Desktop\RecordsMissingFromDeploymentSchedule.csv";        
        const string fileInBothHrAndAd =                            @"C:\Users\nads205\Desktop\InBothHrAndAd.csv";        
        const string fileMissingFromAd =                            @"C:\Users\nads205\Desktop\MissingFromAd.csv";

        readonly List<AdUser> adUsersHr = new List<AdUser>();
        readonly List<AdUser> adUsersUk1 = new List<AdUser>();

        int totalUk1Records = 0;
        int counterUk1 = 0;
        int headerrowUk1 = 0;

        public void ReadAllAdFiles()
        {
            adUsersUk1.Clear();
            ReadAdFiles(fileUk1UsersAll);           
        }

        private void ReadAdFiles(string fileName)
        {
            using (var stream = File.Open(fileName, System.IO.FileMode.Open, System.IO.FileAccess.Read))
            {
                using (var reader = ExcelDataReader.ExcelReaderFactory.CreateCsvReader(stream))
                {
                    var counter = 0;
                    while (reader.Read())
                    {
                        //skip header row
                        if (counter == 0)
                        {
                            counter++;
                            counterUk1++;
                            totalUk1Records++;
                            headerrowUk1++;
                            continue;
                        }

                        counterUk1++;
                        counter++;
                        totalUk1Records++;

                        var firstName = reader[0]?.ToString() ?? null;          //col A  
                        var lastName = reader[1]?.ToString() ?? null;           //col B
                        var samaccountname = reader[2]?.ToString() ?? null;     //col C
                        var emailaddress = reader[3]?.ToString() ?? null;       //col D
                        bool.TryParse(reader.GetString(4), out bool enabled);         //col E
                        //var enabled = b;
                        var phone = reader[5]?.ToString() ?? null;              //col F
                        //skip mob
                        var roleTitle = reader[7]?.ToString() ?? null;          //col H
                        //var lastLogonDate = string.IsNullOrEmpty(reader[8].ToString()) ? new DateTime() : DateTime.ParseExact(reader[8].ToString(),@"dd/MM/yyyy HH:mm:ss", null);  //row I
                        var lastLogonDate1 = reader.GetString(8);
                        DateTime? lastLogonDate = null;
                        if (!string.IsNullOrEmpty(lastLogonDate1)) lastLogonDate = DateTime.ParseExact(lastLogonDate1, @"dd/MM/yyyy HH:mm:ss", null);
                        //DateTime? lastLogonDate = string.IsNullOrEmpty(lastLogonDate1) ? null : DateTime.ParseExact(lastLogonDate1, @"dd/MM/yyyy HH:mm:ss", null);

                        var homedirectory = reader[9]?.ToString() ?? null;      //col J
                        //skip office
                        var organisationalUnit = reader[11]?.ToString() ?? null;  //col J
                        var adUser = new AdUser
                        {                            
                            FirstName = firstName,
                            LastName = lastName,
                            SamAccountName = samaccountname,
                            RoleTitle = roleTitle,
                            EmailAddress = emailaddress,
                            HomeDirectory = homedirectory,
                            Enabled = enabled,
                            LastLogonDate = lastLogonDate,
                            OrganisationalUnit = organisationalUnit
                        };

                        adUsersUk1.Add(adUser);
                    }
                }
            }
        }

        public void ReadHrFile()
        {
            using (var stream = File.Open(fileAdHrUsersNew, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream))
                {
                    var totalRecords = 0;
                    var counter = 0;
                    var headerrow = 0;
                    var skippedrecords = 0;

                    while (reader.Read())
                    {
                        //skip header row
                        if (counter == 0)
                        {
                            counter++;
                            totalRecords++;
                            headerrow++;
                            continue;
                        }
                        counter++;
                        totalRecords++;

                        var samaccountname = reader[5]?.ToString() ?? null;     //col F                   
                                                                                //skip any rows that do not have a SamAccountName
                        if (string.IsNullOrEmpty(samaccountname))
                        {
                            totalRecords++;
                            skippedrecords++;
                            continue;
                        }
                        var displayName = reader[0]?.ToString() ?? null;        //col A
                        var permContract = reader[1]?.ToString() ?? null;       //col B
                        var firstName = reader[2]?.ToString() ?? null;          //col C  
                        var lastName = reader[3]?.ToString() ?? null;           //col D
                        var name = reader[4]?.ToString() ?? null;               //col E
                        //var samaccountname = reader[5]?.ToString() ?? null;     //col F //we do this above
                        var roleTitle = reader[6]?.ToString() ?? null;          //col G
                        var country = reader[7]?.ToString() ?? null;            //col H
                        var location = reader[8]?.ToString() ?? null;           //col I                            
                        var division = reader[9]?.ToString() ?? null;           //col J
                        var department = reader[10]?.ToString() ?? null;        //col K
                        var office = reader[11]?.ToString() ?? null;            //col L
                        var phone = reader[12]?.ToString() ?? null;             //col M
                        var emailaddress = reader[13]?.ToString() ?? null;      //col N
                        var homedirectory = reader[14]?.ToString() ?? null;     //col O 
                        var organisationUnit = reader[15]?.ToString() ?? null;
                        var identity = reader[16]?.ToString() ?? null;
                        var approvedDevice = reader[19].ToString() ?? null;     //col T

                        var adUser = new AdUser
                        {
                            DisplayName = displayName,
                            FirstName = firstName,
                            LastName = lastName,
                            Name = name,
                            SamAccountName = samaccountname,
                            RoleTitle = roleTitle,
                            Location = location,
                            Division = division,
                            Department = department,
                            Office = office,
                            EmailAddress = emailaddress,
                            HomeDirectory = homedirectory,
                            OrganisationalUnit = organisationUnit,
                            Identity = identity,
                            ApprovedDevice = approvedDevice
                        };
                        adUsersHr.Add(adUser);
                    }
                }
            }
        }

        public void CompareRecords()
        {
            var approvedOnly = adUsersHr.Where(c => c.ApprovedDevice == "Yes" || c.ApprovedDevice == "No");
            var notApprovedOnly = adUsersHr.Where(c => c.ApprovedDevice != "Yes" && c.ApprovedDevice != "No");
            Debug.WriteLine($"Total deployment schedule =  {adUsersHr.Count()}");
            Debug.WriteLine($"Approved devices count =  {approvedOnly.Count()}");
            Debug.WriteLine($"Not approved devices count =  {notApprovedOnly.Count()}");

            var newRecords = new List<AdUser>();

            //go through the 1561 that are not approved
            foreach (var user in notApprovedOnly)
            {
                //grab enabled, lastlogondate and distingushedname
                var record = adUsersUk1.SingleOrDefault(c => c.SamAccountName == user.SamAccountName);

                string samAccountName;
                bool? enabled;
                DateTime? lastLogonDate;
                string organisationalUnit;

                if (record == null)
                {
                    samAccountName = user.SamAccountName;
                    enabled = null;
                    lastLogonDate = null;
                    organisationalUnit = "Can't find in AD";
                }
                else
                {
                    samAccountName = record.SamAccountName;
                    enabled = record.Enabled;
                    lastLogonDate = record.LastLogonDate;
                    organisationalUnit = record.OrganisationalUnit;
                }    

                var adUser = new AdUser
                {
                    SamAccountName = samAccountName,
                    Enabled = enabled,
                    LastLogonDate = lastLogonDate,
                    OrganisationalUnit = organisationalUnit
                };
                newRecords.Add(adUser);
            }
       
            var adRecordsMissingFromDeploymentSchedule = new List<AdUser>();

            foreach (var uk1User in adUsersUk1)
            {
                var record = adUsersHr.SingleOrDefault(c => c.SamAccountName == uk1User.SamAccountName);

                if (record == null ) //doesn't exist
                {
                    adRecordsMissingFromDeploymentSchedule.Add(uk1User);
                }

            }

            ////write these to file            
            using (var writer = new StreamWriter(fileNew))
            {
                using (var csv = new CsvHelper.CsvWriter(writer))
                {
                    csv.WriteRecords(newRecords);
                }
            }


            using (var writer = new StreamWriter(fileRecordsMissingFromDeploymentSchedule))
            {
                using (var csv = new CsvHelper.CsvWriter(writer))
                {
                    csv.WriteRecords(adRecordsMissingFromDeploymentSchedule);
                }
            }
        }

    }
 
}
