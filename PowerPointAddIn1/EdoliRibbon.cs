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

        private void grid_Click(object sender, RibbonControlEventArgs e)
        {
            AlignTool.AlignGrid();
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

        private void transpose_Click(object sender, RibbonControlEventArgs e)
        {
            AlignTool.Transpose();
        }

        private void groupLabel_Click(object sender, RibbonControlEventArgs e)
        {
            AlignTool.GroupLabels();
        }

        private void alignPrevSlide_Click(object sender, RibbonControlEventArgs e)
        {
            AlignTool.AlignWithPreviousSlide();
        }

        private void editBoxName_TextChanged(object sender, RibbonControlEventArgs e)
        {
            AnimationTool.SetNameOfActive(this.editBoxName.Text);
        }

        private void swapCycle_Click(object sender, RibbonControlEventArgs e)
        {
            AlignTool.SwapCycle();
        }

        private void swapCycleReverse_Click(object sender, RibbonControlEventArgs e)
        {
            AlignTool.SwapCycleReverse();
        }

        private void snapUpLeft_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void snapUpRight_Click(object sender, RibbonControlEventArgs e)
        {

        }
    }
}
