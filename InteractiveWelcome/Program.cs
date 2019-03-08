using System;
using System.DirectoryServices;
using System.Diagnostics;


namespace InteractiveWelcome
{
    class Program
    {
        static void Main(string[] args)
        {
            var doExit = false;
            while (doExit != true)
            {
                string ldapFilter = "";
                string ou = "";
                Console.WriteLine("Choose the OU you want to Search\n");
                Console.WriteLine("1 - Nellis");
                Console.WriteLine("2 - Arcata Way");
                Console.WriteLine("3 - Computer Name");
                Console.WriteLine("S - Shutdown Computer");
                Console.WriteLine("R - Reboot Computer");

                Console.WriteLine("Q - Quit");

                char myChoice = Console.ReadKey(true).KeyChar;

                switch (myChoice)
                {
                    case '1':
                        ou = "LDAP://OU=Arcata Way,OU=NTTR,OU=Workstations,DC=jt3,DC=com";
                        ldapFilter = "(&(objectClass=computer))";
                        break;
                    case '2':
                        ou = "LDAP://OU=Nellis,OU=NTTR,OU=Workstations,DC=jt3,DC=com";
                        ldapFilter = "(&(objectClass=computer))";
                        break;
                    case '3':
                        ldapFilter = "(&(objectClass=computer)(cn=LSV-VM-TRN*))";
                        break;
                    case 'S':
                    case 's':
                        Process.Start("shutdown", @"/s /m \\LSV-VM-TRN1 /t 0");
                        Console.Clear();
                        continue;
                    case 'R':
                    case 'r':
                        Process.Start("shutdown", @"/r /m \\LSV-VM-TRN1 /t 0");
                        Console.Clear();
                        continue;
                    case 'Q':
                    case 'q':
                        doExit = true;
                        continue;
                    default:
                        Console.WriteLine("{0} is not a valid choice", myChoice);
                        continue;
                }
                DirectoryEntry entry = new DirectoryEntry("LDAP://JT3-DC-ARC.JT3.com")
                    {
                        Path = ou,
                        AuthenticationType = AuthenticationTypes.Secure
                    };

                using (DirectorySearcher ds = new DirectorySearcher(entry))
                {
                    ds.PropertiesToLoad.Add("name");
                    ds.Filter = ldapFilter;
                    ds.Sort.Direction = SortDirection.Ascending;
                    ds.Sort.PropertyName = "name";
                    SearchResultCollection results = ds.FindAll();

                    int itemCount = results.Count;
                 
                    foreach (SearchResult result in results)
                    {

                        Console.WriteLine("{0}" , result.Properties["name"][0].ToString());
                    }
                    Console.WriteLine("Computer Count: {0}", itemCount);
                    Console.ReadKey();
                    Console.Clear();
                }
            }
            Environment.Exit(0);
        }
    }
    
}