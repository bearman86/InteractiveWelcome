using System;
using System.DirectoryServices;


namespace InteractiveWelcome
{
    class Program
    {
        static void Main(string[] args)
        {

            while (true)
            {
                string myChoice;
                string ou = "";
                do
                {
                    Console.WriteLine("Choose the OU you want to Search\n");
                    Console.WriteLine("1 - Nellis");
                    Console.WriteLine("2 - Arcata Way");
                    Console.WriteLine("Q - Quit");

                    myChoice = Console.ReadLine();

                    switch (myChoice)
                    {
                        case "1":
                            ou = "LDAP://OU=Arcata Way,OU=NTTR,OU=Workstations,DC=jt3,DC=com";
                            break;
                        case "2":
                            ou = "LDAP://OU=Nellis,OU=NTTR,OU=Workstations,DC=jt3,DC=com";
                            break;
                        case "Q":
                            Environment.Exit(0);
                            break;
                        case "q":
                            Environment.Exit(0);
                            break;
                        default:
                            Console.WriteLine("{0} is not a valid choice", myChoice);
                            break;
                    }
                } while (myChoice != "1" && myChoice != "2");


                Console.Write("Choose the number for the OU you want to search.\n");
                DirectoryEntry entry = new DirectoryEntry("LDAP://JT3-DC-ARC.JT3.com")
                {
                    Path = ou,
                    AuthenticationType = AuthenticationTypes.Secure
                };

                using (DirectorySearcher ds = new DirectorySearcher(entry))
                {
                    ds.PropertiesToLoad.Add("name");
                    ds.Filter = "(&(objectClass=computer))";
                    ds.Sort.Direction = SortDirection.Ascending;
                    ds.Sort.PropertyName = "name";
                    SearchResultCollection results = ds.FindAll();

                    int itemCount = results.Count;

                    foreach (SearchResult result in results)
                    {
                        Console.WriteLine("{0}", result.Properties["name"][0].ToString());
                    }
                    Console.WriteLine("Computer Count: {0}", itemCount);
                    Console.Read();
                    Console.Clear();
                }
            }
            
        }
    }
    
}