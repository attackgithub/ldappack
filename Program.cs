using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using System.Linq;
using Unosquare.Swan;
using Unosquare.Swan.Formatters;

namespace LdapPack
{
    class Program
    {
        public static Dictionary<string, DateTime> UsersLastLogOnDate()
        {
            var lastLogins = new Dictionary<string, DateTime>();

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
                            searcher.Filter =
                                "(&(objectCategory=person)(objectClass=user)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))";
                            searcher.PropertiesToLoad.AddRange(new[] {"distinguishedName", "lastLogon"});

                            foreach (SearchResult searchResult in searcher.FindAll())
                            {
                                if (!searchResult.Properties.Contains("lastLogon")) continue;

                                var lastLogOn = DateTime.FromFileTime((long) searchResult.Properties["lastLogon"][0]);
                                var username = searchResult.Properties["distinguishedName"][0].ToString();

                                if (lastLogins.ContainsKey(username))
                                {
                                    if (lastLogOn > lastLogins[username])
                                    {
                                        lastLogins[username] = lastLogOn;
                                    }
                                }
                                else
                                {
                                    lastLogins.Add(username, lastLogOn);
                                }
                            }
                        }

                    }
                }
                catch
                {
                    // Domain controller is down or not responding
                    $"Domain controller {controller.Name} is not responding.".Debug();
                }
            }
            return lastLogins;
        }

        static void Main(string[] args)
        {
            var data = UsersLastLogOnDate();
            $"Found {data.Count} records".Info();

            var days = "How many days before?".ReadNumber(30);

            var filtered = data.Where(x => x.Value < DateTime.Now.AddDays(-days)).OrderBy(x => x.Key);

            foreach (var item in filtered)
                $"{item.Key} - {item.Value}".Info();

            $"Filtered {filtered.Count()} records".Info();

            var key = "Do you want to export to CSV (y)".ReadKey();

            if (key.Key == ConsoleKey.Y)
            {
                var filename = Runtime.GetDesktopFilePath("ldap.csv");
                CsvWriter.SaveRecords(filtered, filename);
                $"Exporting done {filename}".Info();
            }

            "Press any key to exit...".Info();
            Terminal.ReadLine();
        }
    }
}