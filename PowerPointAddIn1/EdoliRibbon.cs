using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace PowerPointAddIn1
{
    public partial class EdoliRibbon
    {
        private void EdoliRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void labelBottom_Click(object sender, RibbonControlEventArgs e)
        {
            AlignTool.AlignLabels(ShapeExt.Anchor.Bottom);
        }

        private void labelTop_Click(object sender, RibbonControlEventArgs e)
        {
            AlignTool.AlignLabels(ShapeExt.Anchor.Top);
        }

        private void labelLeft_Click(object sender, RibbonControlEventArgs e)
        {
            AlignTool.AlignLabels(ShapeExt.Anchor.Left);
        }

        private void labelRight_Click(object sender, RibbonControlEventArgs e)
        {
            AlignTool.AlignLabels(ShapeExt.Anchor.Right);
        }

        private void editBoxName_TextChanged(object sender, RibbonControlEventArgs e)
        {
            AnimationTool.SetNameOfActive(this.editBoxName.Text);
        }
    }
}
