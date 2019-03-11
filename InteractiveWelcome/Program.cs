using System;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using System.Management;
using System.Security.Permissions;


namespace InteractiveWelcome
{
    class Program
    {
        public static string ReadPassword()
        {
            string password = "";
            ConsoleKeyInfo info = Console.ReadKey(true);
            while (info.Key != ConsoleKey.Enter)
            {
                if (info.Key != ConsoleKey.Backspace)
                {
                    Console.Write("*");
                    password += info.KeyChar;
                }
                else if (info.Key == ConsoleKey.Backspace)
                {
                    if (!string.IsNullOrEmpty(password))
                    {
                        password = password.Substring(0, password.Length - 1);
                        int pos = Console.CursorLeft;
                        Console.SetCursorPosition(pos - 1, Console.CursorTop);
                        Console.Write(" ");
                        Console.SetCursorPosition(pos - 1, Console.CursorTop);
                    }
                }
                info = Console.ReadKey(true);
            }
            Console.WriteLine();
            return password;
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
            Console.WriteLine("Pls key in your Login ID");
            var loginid = Console.ReadLine();
            Console.WriteLine("Pls key in your Password");
            var password = ReadPassword();
            Console.Write("Your Password is:" + password);
            Console.ReadLine();

            

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