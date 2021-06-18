using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace EditDim3
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);

        static SldWorks swApp;
        static IModelDoc2 swModel;
        static ModelView swModelView;
        static Mouse swMouse;
        static SelectionMgr swSelMgr;
        static DrawingDoc swDoc;

        static App app;
        static MainWindow wnd;
        int launchCount = 0;

        System.Threading.Mutex mutex;

        [STAThread]
        public static void Main()
        {
            app = new App();
            app.Run();
        }

        private void App_Startup(object sender, StartupEventArgs e)
        {
            bool createdNew;
            string mutName = "EditDimension_nic";
            mutex = new System.Threading.Mutex(true, mutName, out createdNew);
            if (!createdNew)
            {
                this.Shutdown();
            }
            DoubleClickHandler();
        }

        App()
        {
            InitializeComponent();
            swApp = (SldWorks)GetSwAppFromProcess();
            ActivateExitHandler();
            swApp.ActiveModelDocChangeNotify += this.ActiveModelDocChangeHandler;
            wnd = new MainWindow();
            wnd.Title = "wnd";
            wnd.Show();
        }

        private void showMainWindow()
        {
            launchCount++;
            Application.Current.Dispatcher.Invoke(() =>
            {
                MainWindow taskWindow = new MainWindow();
                taskWindow.Title = "wnd " + launchCount;
                taskWindow.Show();
            });
        }

        private static void ActivateExitHandler()
        {
            swApp.DestroyNotify += exitHandler;
        }

        private static int exitHandler()
        {
            int funcResult = 0;
            appShutdown();
            return funcResult;
        }

        public static void appShutdown()
        {
            Application.Current.Shutdown();
        }


        private  int ActiveModelDocChangeHandler()
        {
            int funcResult = 0;
            swMouse.MouseLBtnDblClkNotify -= this.mouseDbl;
            DoubleClickHandler();
            return funcResult;
        }

        private int DoubleClickHandler()
        {
            int funcResult = 0;
            if (swApp != null)
            {
                swModel = (ModelDoc2)swApp.ActiveDoc;
                swModelView = (ModelView)swModel.GetFirstModelView();
                swMouse = (Mouse)swModelView.GetMouse();
                swSelMgr = (SelectionMgr)swModel.SelectionManager;
                swMouse.MouseLBtnDblClkNotify += this.mouseDbl;
            }
            return funcResult;
        }

 

        private int mouseDbl(int x, int y, int param)
        {
            int funcResult = 0;
            app.showMainWindow();
            return funcResult;
        }


        private static ISldWorks GetSwAppFromProcess()
        {
            Process[] procs = Process.GetProcessesByName("SLDWORKS");

            var myPrc = procs[0];

            var monikerName = "SolidWorks_PID_" + myPrc.Id.ToString();

            IBindCtx context = null;
            IRunningObjectTable rot = null;
            IEnumMoniker monikers = null;

            try
            {
                CreateBindCtx(0, out context);

                context.GetRunningObjectTable(out rot);
                rot.EnumRunning(out monikers);

                var moniker = new IMoniker[1];

                while (monikers.Next(1, moniker, IntPtr.Zero) == 0)
                {
                    var curMoniker = moniker.First();

                    string name = null;

                    if (curMoniker != null)
                    {
                        try
                        {
                            curMoniker.GetDisplayName(context, null, out name);
                        }
                        catch (UnauthorizedAccessException)
                        {
                        }
                    }

                    if (string.Equals(monikerName,
                        name, StringComparison.CurrentCultureIgnoreCase))
                    {
                        object app;
                        rot.GetObject(curMoniker, out app);
                        return app as ISldWorks;
                    }
                }
            }
            finally
            {
                if (monikers != null)
                {
                    Marshal.ReleaseComObject(monikers);
                }

                if (rot != null)
                {
                    Marshal.ReleaseComObject(rot);
                }

                if (context != null)
                {
                    Marshal.ReleaseComObject(context);
                }
            }

            return null;
        }
    }
}
