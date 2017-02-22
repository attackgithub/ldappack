using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using System.Linq;
using Unosquare.Swan;
using Unosquare.Swan.Formatters;

namespace LdapPack
{
    class Program
    {
        private static void GetLdapInfo(Action<SearchResult> callback, string filter, string[] propertiesToLoad)
        {
            foreach (DomainController controller in Domain.GetCurrentDomain().DomainControllers)
            {
                $"Querying: {controller.Name}".Info();

                try
                {
                    using (var directoryEntry = new DirectoryEntry($"LDAP://{controller.Name}"))
                    {
                        using (var searcher = new DirectorySearcher(directoryEntry))
                        {
                            searcher.PageSize = 1000;
                            searcher.Filter = filter;
                            searcher.PropertiesToLoad.AddRange(propertiesToLoad);

                            foreach (SearchResult searchResult in searcher.FindAll())
                            {
                                callback(searchResult);
                            }
                        }

                    }
                }
                catch (Exception ex)
                {
                    ex.Message.Error();
                    $"Domain controller {controller.Name} is not responding.".Error();
                }
            }
        }

        private static void GetStaleUsers()
        {
            var data = new Dictionary<string, DateTime>();

            var callback = new Action<SearchResult>((searchResult) =>
            {
                if (!searchResult.Properties.Contains("lastLogon")) return;

                var lastLogOn = DateTime.FromFileTime((long) searchResult.Properties["lastLogon"][0]);
                var username = searchResult.Properties["distinguishedName"][0].ToString();

                if (data.ContainsKey(username))
                {
                    if (lastLogOn > data[username])
                    {
                        data[username] = lastLogOn;
                    }
                }
                else
                {
                    data.Add(username, lastLogOn);
                }
            });

            GetLdapInfo(callback,
                "(&(objectCategory=person)(objectClass=user)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))",
                new[] {"distinguishedName", "lastLogon" });

            if (data.Count == 0)
            {
                "Found no data".Warn();
                return;
            }

            $"Found {data.Count} records".Info();

            var days = "How many days before?".ReadNumber(30);

            var filtered = data.Where(x => x.Value < DateTime.Now.AddDays(-days)).OrderBy(x => x.Key);

            foreach (var item in filtered)
                $"{item.Key} - {item.Value}".Info();

            $"Filtered {filtered.Count()} records".Info();
            
            if ("Do you want to export to CSV? (y)".ReadKey().Key != ConsoleKey.Y) return;

            var filename = Runtime.GetDesktopFilePath("ldap-stale-users.csv");
            CsvWriter.SaveRecords(filtered, filename);
            $"Export done {filename}".Info();
            
            if ("Do you want to open the file? (y)".ReadKey().Key != ConsoleKey.Y) return;

            (new Process()
            {
                StartInfo = new ProcessStartInfo {FileName = filename}
            }).Start();
        }

        private static void GetStaleComputers()
        {
            var data = new Dictionary<string, Tuple<DateTime, string>>();

            var callback = new Action<SearchResult>((searchResult) =>
            {
                if (!searchResult.Properties.Contains("PwdLastSet")) return;

                var pwdLastSet = DateTime.FromFileTime((long) searchResult.Properties["PwdLastSet"][0]);
                var username = searchResult.Properties["distinguishedName"][0].ToString();
                var operatingsystem = searchResult.Properties.GetValue("operatingsystem");
                
                if (data.ContainsKey(username))
                {
                    if (pwdLastSet > data[username].Item1)
                    {
                        data[username] = new Tuple<DateTime, string>(pwdLastSet, operatingsystem);
                    }
                }
                else
                {
                    data.Add(username, new Tuple<DateTime, string>(pwdLastSet, operatingsystem));
                }
            });

            GetLdapInfo(callback,
                "(&(objectCategory=computer)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(!(isCriticalSystemObject=TRUE)))",
                new[] {"PwdLastSet", "distinguishedName", "operatingsystem", "lastLogonTimeStamp " });

            if (data.Count == 0)
            {
                "Found no data".Warn();
                return;
            }

            $"Found {data.Count} records".Info();
            
            var days = "How many days before?".ReadNumber(30);
            var filtered = data.Select(x => new
            {
                Computer = x.Key,
                OperatingSystem = x.Value.Item2,
                Stale = x.Value.Item1 < DateTime.Now.AddDays(-days) ? "TRUE" : "FALSE"
            }).ToList();

            foreach (var item in filtered)
                $"{item.Computer} - {item.OperatingSystem} - {item.Stale}".Info();

            if ("Do you want to export to CSV? (y)".ReadKey().Key != ConsoleKey.Y) return;

            var filename = Runtime.GetDesktopFilePath("ldap-stale-computers.csv");
            CsvWriter.SaveRecords(filtered, filename);

            $"Export done {filename}".Info();

            if ("Do you want to open the file? (y)".ReadKey().Key != ConsoleKey.Y) return;

            (new Process()
            {
                StartInfo = new ProcessStartInfo { FileName = filename }
            }).Start();
        }

        static void Main(string[] args)
        {
            Runtime.WriteWelcomeBanner();

            var options = new Dictionary<ConsoleKey, string>()
            {
                {ConsoleKey.U, "Get stale users"},
                {ConsoleKey.C, "Get stale computers"}
            };

            var continueLoop = true;

            while (continueLoop)
            {
                var response = "Choose an option".ReadPrompt(options, "Press any key to close...");

                switch (response.Key)
                {
                    case ConsoleKey.U:
                        GetStaleUsers();
                        break;
                    case ConsoleKey.C:
                        GetStaleComputers();
                        break;
                    default:
                        continueLoop = false;
                        break;
                }
            }
        }
    }

    static class Extensions
    {
        public static string GetValue(this ResultPropertyCollection coll, string key)
        {
            return coll[key].Count > 0 ? coll[key][0].ToString() : null;
        }
    }
}