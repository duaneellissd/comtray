using System;
using System.Threading;
using System.Windows.Forms;

namespace comtray
{
    public class ComTrayMain
    {
        static public Mutex single_instance_mutex;

        [STAThread]
        static void Main()
        {
            /* prevent multiple instances */
            const string my_single_mutex_name = "duanes_comtray_F95187C1455E43D1A6CC82DD50726F08";
            bool createdNew;

            single_instance_mutex = new Mutex(true, my_single_mutex_name, out createdNew);
            if (!createdNew)
            {

                MessageBox.Show("ComTray is already running");
                return;
            }

            /* create app, but do not show it */
            var app = new ComTrayGui();
            /* run, with app hidden */
            app.show_form(false);
            Application.Run();
        }
    }
}