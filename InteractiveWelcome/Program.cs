using System;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using System.Management;


namespace InteractiveWelcome
{
    class Program
    {
        public static void PasswordEncodingMethod()
        {
            string pass = "";
            do
            {
                ConsoleKeyInfo key = Console.ReadKey(true);
                // Backspace Should Not Work
                if (key.Key != ConsoleKey.Backspace && key.Key != ConsoleKey.Enter)
                {
                    pass += key.KeyChar;
                    Console.Write("*");
                }
                else
                {
                    if (key.Key == ConsoleKey.Backspace && pass.Length > 0)
                    {
                        pass = pass.Substring(0, (pass.Length - 1));
                        Console.Write("\b \b");
                    }
                    else if (key.Key == ConsoleKey.Enter)
                    {
                        break;
                    }
                }
            } while (true);
        }

        public static void DoHostCommand(string compName, string doCommand)
        {
            string adsiPath = string.Format(@"\\{0}\root\cimv2", compName);
            ManagementScope scope = new ManagementScope(adsiPath);
            ManagementPath osPath = new ManagementPath("Win32_OperatingSystem");
            ManagementClass os = new ManagementClass(scope, osPath, null);

            ManagementObjectCollection instances;
            instances = os.GetInstances();
            foreach (ManagementObject instance in instances)
            {
                object result = instance.InvokeMethod(doCommand, new object[] { });
                uint returnValue = (uint) result;
            }
        }
        
        static void Main(string[] args)
        {


            DomainCollection dc = Forest.GetCurrentForest().Domains;
            foreach (Domain d in dc)
            {
                Console.WriteLine(d.Name);
            }
            
            var doExit = false;
            while (doExit != true)
            {
                string ldapFilter = "";
                string ou = "";
                Console.WriteLine("Choose the OU you want to Search or Function.\n");
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
                        DoHostCommand("LSV-VM-TRN1","Shutdown");
                        Console.Clear();
                        continue;
                    case 'R':
                    case 'r':
                        DoHostCommand("LSV-VM-TRN0", "Reboot");
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
                        Console.WriteLine("{0}" , result.Properties["name"][0]);
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