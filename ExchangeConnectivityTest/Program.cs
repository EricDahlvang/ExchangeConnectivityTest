using System;

namespace ExchangeConnectivityTest
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("This application uses Microsoft.Exchange.WebServices to test connectivity to an email account.");
            Console.WriteLine("");
            Console.WriteLine("");

            Console.WriteLine("Please enter your Email Address:");
            var emailAddress = Console.ReadLine();
            Console.WriteLine("Please enter your Password:");
            var password = ReadPassword();

            var result = CredentialsTester.TestCredentials(emailAddress, password, true);

            Console.WriteLine("");
            Console.WriteLine("");

            if (result != null)
            {
                Console.WriteLine($"Version: {result.Version}");
                Console.WriteLine("Test Succeeded");
            }
            else
                Console.WriteLine("Test Failed");

            Console.WriteLine("(Hit any key to close)");
            Console.ReadKey();
        }

        //http://rajeshbailwal.blogspot.com/2012/03/password-in-c-console-application.html
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
                        // remove one character from the list of password characters
                        password = password.Substring(0, password.Length - 1);
                        // get the location of the cursor
                        int pos = Console.CursorLeft;
                        // move the cursor to the left by one character
                        Console.SetCursorPosition(pos - 1, Console.CursorTop);
                        // replace it with space
                        Console.Write(" ");
                        // move the cursor to the left by one character again
                        Console.SetCursorPosition(pos - 1, Console.CursorTop);
                    }
                }
                info = Console.ReadKey(true);
            }
            // add a new line because user pressed enter at the end of their password
            Console.WriteLine();
            return password;
        }
    }
}

