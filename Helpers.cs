using System.Globalization;

namespace GmailAPIExample
{
    public static class Helpers
    {
        public static T? Safe<T>(Func<T> func)
        {
            try
            {
                return func();
            }
            catch (System.Exception e)
            {
                System.Console.WriteLine(e);
            }

            return default(T);
        }

        public static T Retry<T>(Func<T> f)
        {
            while (true)
            {
                try
                {
                    return f();
                }
                catch (Google.GoogleApiException ex)
                {
                    if (ex.HttpStatusCode == System.Net.HttpStatusCode.TooManyRequests)
                    {
                        // Extract the retry time from the exception message
                        Console.WriteLine(ex.Message);
                        string retryAfter = ex.Message.Split('(')[0].Trim().Split(' ').Last();
                        // Create a CultureInfo object with the "en-US" culture
                        CultureInfo culture = new CultureInfo("en-US");

                        // Define the custom date and time format
                        string format = "yyyy-MM-ddTHH:mm:ss.fffZ";

                        DateTime retryTime = DateTime.ParseExact(retryAfter, format, culture);

                        // Calculate the delay until the retry time
                        TimeSpan delay = retryTime - DateTime.Now;

                        if (delay.TotalSeconds > 0)
                        {
                            Console.WriteLine($"User-rate limit exceeded. Retrying after {delay.TotalSeconds} seconds...");

                            // Wait for the specified delay before retrying
                            System.Threading.Thread.Sleep(delay);

                            // Retry the operation
                        }
                        else
                        {
                            Console.WriteLine("User-rate limit exceeded, but the retry time has already passed.");
                        }
                    }
                    else
                    {
                        // Handle other types of GoogleApiException
                        Console.WriteLine($"GoogleApiException occurred: {ex.Message} {ex}");
                        throw;
                    }
                }
                catch (TaskCanceledException)
                {
                    Console.WriteLine($"Got TaskCanceledException: Retrying ...");
                }
            }
        }
    }
}
