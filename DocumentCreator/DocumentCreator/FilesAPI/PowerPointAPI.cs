using System;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace DocumentCreator.FilesAPI
{
    class PowerPointAPI
    {
        private static PowerPoint.Application app = new PowerPoint.Application();

        public static PowerPoint.Presentation GetPresentation(string fileName)
        {
            //app.Visible = true;

            PowerPoint.Presentation pre = null;

            try
            {
                pre = app.Presentations.Open(fileName);
            }
            catch (Exception e)
            {
                throw new Exception("Can't open file", e);
            }
            return pre;
        }

        public static PowerPoint.Application GetPowerPointApp()
        {
            return app;
        }

        public static void SaveFile(PowerPoint.Presentation pre, string fileName = "")
        {
            if (string.IsNullOrEmpty(fileName))
            {
                try
                {
                    pre.Save();
                }
                catch (Exception e)
                {
                    throw new Exception("Can't save file", e);
                }
            }
            else
            {
                try
                {
                    pre.SaveAs(fileName);
                }
                catch (Exception e)
                {
                    throw new Exception("Can't save file in " + fileName, e);
                }
            }
        }

        public static void Close(PowerPoint.Presentation pre)
        {
            if (pre != null)
            {
                pre.Close();
                KillPowerPoint();
            }
            else
            {
                throw new NullReferenceException();
            }
        }

        public static void KillPowerPoint()
        {
            System.Diagnostics.Process[] procs = System.Diagnostics.Process.GetProcessesByName("POWERPNT");
            foreach (System.Diagnostics.Process p in procs)
            {
                p.Kill();
            }
        }
    }
}
