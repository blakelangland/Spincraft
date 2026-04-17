using System;
using System.Collections.Generic;
using HorizonScientific;

namespace CredentialManager
{
    internal class Program
    {
        // Maps each known credential key to a human-readable label and whether the
        // value should be masked during input (i.e. passwords).
        private static readonly (string Key, string Label, bool Masked)[] KnownCredentials =
        {
            (HSCredentialStore.Keys.SpincraftServiceAccountId,  "Spincraft Service Account ID",           false),
            (HSCredentialStore.Keys.SpincraftServicePassword,   "Spincraft Service Account Password",     true),
            (HSCredentialStore.Keys.NoReplyLogin,    "No-Reply SMTP Login",    false),
            (HSCredentialStore.Keys.NoReplyPassword, "No-Reply SMTP Password", true),
        };

        static void Main(string[] args)
        {
            Console.Title = "Spincraft WI - Credential Manager";

            while (true)
            {
                Console.Clear();
                PrintHeader();
                PrintCredentialTable();
                PrintMenu();

                string choice = Console.ReadLine()?.Trim().ToUpperInvariant() ?? string.Empty;

                if (choice == "Q")
                {
                    Console.WriteLine("\nExiting.");
                    break;
                }
                else if (choice == "A")
                {
                    SetAllCredentials();
                }
                else if (int.TryParse(choice, out int index) &&
                         index >= 1 && index <= KnownCredentials.Length)
                {
                    var credential = KnownCredentials[index - 1];
                    SetSingleCredential(credential.Key, credential.Label, credential.Masked);
                    Pause();
                }
                else
                {
                    Console.WriteLine("\n  Invalid choice. Press any key to try again.");
                    Console.ReadKey(true);
                }
            }
        }

        // -------------------------------------------------------------------------
        // Display helpers
        // -------------------------------------------------------------------------

        private static void PrintHeader()
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("  =====================================================");
            Console.WriteLine("    Spincraft WI  -  Credential Manager");
            Console.WriteLine("    Credentials are stored in:");
            Console.WriteLine("      %ProgramData%\\Spincraft\\sc.cred");
            Console.WriteLine("    Encrypted with Windows DPAPI (machine-scoped).");
            Console.WriteLine("  =====================================================");
            Console.ResetColor();
            Console.WriteLine();
        }

        private static void PrintCredentialTable()
        {
            Console.WriteLine("  Known credentials:\n");
            Console.WriteLine($"  {"#",-4} {"Label",-42} {"Status",-10}");
            Console.WriteLine($"  {new string('-', 4)} {new string('-', 42)} {new string('-', 10)}");

            for (int i = 0; i < KnownCredentials.Length; i++)
            {
                var (key, label, _) = KnownCredentials[i];
                string status = GetStatus(key);
                ConsoleColor statusColor = status == "SET" ? ConsoleColor.Green : ConsoleColor.Yellow;

                Console.Write($"  [{i + 1}]  {label,-42} ");
                Console.ForegroundColor = statusColor;
                Console.WriteLine(status);
                Console.ResetColor();
            }

            Console.WriteLine();
        }

        private static void PrintMenu()
        {
            Console.WriteLine("  Options:");
            Console.WriteLine("    Enter a number (1-{0}) to update a single credential", KnownCredentials.Length);
            Console.WriteLine("    [A]  Set all credentials");
            Console.WriteLine("    [Q]  Quit");
            Console.WriteLine();
            Console.Write("  Choice: ");
        }

        // -------------------------------------------------------------------------
        // Credential entry
        // -------------------------------------------------------------------------

        private static void SetAllCredentials()
        {
            Console.Clear();
            PrintHeader();
            Console.WriteLine("  Setting all credentials. Press Enter with no input to skip a field.\n");

            bool anySet = false;

            foreach (var (key, label, masked) in KnownCredentials)
            {
                string value = PromptForValue(label, masked);
                if (string.IsNullOrEmpty(value))
                {
                    Console.ForegroundColor = ConsoleColor.DarkGray;
                    Console.WriteLine("    Skipped.\n");
                    Console.ResetColor();
                    continue;
                }

                if (SaveCredential(key, value))
                    anySet = true;
            }

            if (anySet)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("\n  All entered credentials have been saved.");
                Console.ResetColor();
            }
            else
            {
                Console.WriteLine("\n  No credentials were changed.");
            }

            Pause();
        }

        private static void SetSingleCredential(string key, string label, bool masked)
        {
            Console.WriteLine();
            string value = PromptForValue(label, masked);

            if (string.IsNullOrEmpty(value))
            {
                Console.ForegroundColor = ConsoleColor.DarkGray;
                Console.WriteLine("  No value entered — credential unchanged.");
                Console.ResetColor();
                return;
            }

            SaveCredential(key, value);
        }

        // Prompts for a credential value, masking input with '*' when masked = true.
        private static string PromptForValue(string label, bool masked)
        {
            Console.Write("  {0}: ", label);

            if (!masked)
                return Console.ReadLine()?.Trim() ?? string.Empty;

            // Masked input: read character by character, print '*' in place of each.
            var input = new System.Text.StringBuilder();
            while (true)
            {
                ConsoleKeyInfo key = Console.ReadKey(intercept: true);

                if (key.Key == ConsoleKey.Enter)
                {
                    Console.WriteLine();
                    break;
                }
                else if (key.Key == ConsoleKey.Backspace && input.Length > 0)
                {
                    input.Remove(input.Length - 1, 1);
                    Console.Write("\b \b"); // erase the last '*'
                }
                else if (!char.IsControl(key.KeyChar))
                {
                    input.Append(key.KeyChar);
                    Console.Write('*');
                }
            }

            return input.ToString();
        }

        // Calls HSCredentialStore.Set and prints the result. Returns true on success.
        private static bool SaveCredential(string key, string value)
        {
            try
            {
                HSCredentialStore.Set(key, value);
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("  [OK] Saved: {0}", key);
                Console.ResetColor();
                return true;
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("  [ERROR] Failed to save '{0}': {1}", key, ex.Message);
                Console.ResetColor();
                return false;
            }
        }

        // -------------------------------------------------------------------------
        // Status check
        // -------------------------------------------------------------------------

        private static string GetStatus(string key)
        {
            try
            {
                HSCredentialStore.Get(key);
                return "SET";
            }
            catch (InvalidOperationException)
            {
                // Store file does not exist yet — no keys are provisioned.
                return "NOT SET";
            }
            catch (KeyNotFoundException)
            {
                return "NOT SET";
            }
        }

        private static void Pause()
        {
            Console.WriteLine("\n  Press any key to return to the menu...");
            Console.ReadKey(true);
        }
    }
}
