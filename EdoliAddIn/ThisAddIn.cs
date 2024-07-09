using System;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.PowerPoint;

namespace EdoliAddIn
{
    public partial class ThisAddIn
    {
        void Application_PresentationNewSlide(PowerPoint.Slide Sld)
        {
        }

        public void WindowSelectionChange(Selection sel)
        {
            var shapes = Util.ListSelectedShapes();
            String name = "";
            if (shapes.Count > 0)
            {
                var shape = shapes[0];
                name = shape.Name;
            }
            Globals.Ribbons.EdoliRibbon.animationName.Text = name;
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            KeyboardHook.SetHook();
            this.Application.WindowSelectionChange += new PowerPoint.EApplication_WindowSelectionChangeEventHandler(WindowSelectionChange);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            KeyboardHook.ReleaseHook();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion

    }
}
