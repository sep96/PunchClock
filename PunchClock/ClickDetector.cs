using Gma.System.MouseKeyHook;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PunchClock
{
    class ClickDetector
    {
        public static void ListenForMouseEvents()
        {
            Console.WriteLine("Listening to mouse clicks.");

            //When a mouse button is pressed 
            Hook.GlobalEvents().MouseDown += async (sender, e) =>
            {
                try
                {
                    StreamWriter sw = new StreamWriter(Application.StartupPath + @"\Mouselog.txt", true);
                }
                catch (Exception wss) { }
            };
            //When a double click is made
            Hook.GlobalEvents().MouseDoubleClick += async (sender, e) =>
            {
                try
                {
                    StreamWriter sw = new StreamWriter(Application.StartupPath + @"\Mouselog.txt", true);
                }
                catch (Exception ss) { }
            };
        }
    }
}
