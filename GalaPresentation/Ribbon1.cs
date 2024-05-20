using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace GalaPresentation
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void Run_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.run();
        }
    }
}
