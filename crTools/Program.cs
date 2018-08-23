using CRAPI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace crTools
{
    static class Program
    {
        private readonly static string myDevelperKey = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpZCI6MTE4OCwiaWRlbiI6IjQ2NzI0ODM4Nzk2ODg2MDE2MSIsIm1kIjp7fSwidHMiOjE1MzE1Mjc5MTQ4Nzh9.rdg-8SUNxiY12igd_wzrMYeOQFT9d8eU6rS-9rEZ9dM";
        private static Wrapper clashRoyale;
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            // There are async versions of every method here. You can call them with async/await.

            // Create wrapper instance. First parameter is your unique API key. Make sure to not share it with anyone!
            //   Request one at API discord. See https://docs.royaleapi.com/#/authentication?id=key-management
            // Second parameter is if you want wrapper to automatically throttle your requests to satisfy API limit. This is set to True by default.
            //   default. This is here because if you try to exceed API limit, it'll temporarily (10 seconds or something like that) disable your account.
            // Thrid parameter is wrapper cache duration, set to 0 by default. If it is set to 0 or negative number, it is disabled.
            //   Caching is good to use if you process a lot of same requests. Wrapper will save responses to these requests and store it in TEMP folder.
            //   Data will not be always up-to-date, even if you disable this cache, as API refreshes data only once every few minutes.
            clashRoyale = new Wrapper(myDevelperKey);
            Application.Run(new Form1(ref clashRoyale));
        }
    }
}
