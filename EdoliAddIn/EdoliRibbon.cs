using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Scripting.Hosting;

namespace EdoliAddIn
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
            AlignTool.AlignWithSiblingSlide(-1);
        }

        private void alignNextSlide_Click(object sender, RibbonControlEventArgs e)
        {
            AlignTool.AlignWithSiblingSlide(1);
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

        private void snapDownRight_Click(object sender, RibbonControlEventArgs e)
        {
            AlignTool.SnapDownRight();
        }

        private void snapUpRight_Click(object sender, RibbonControlEventArgs e)
        {
            AlignTool.SnapUpRight();
        }

        private void alignGrid_Click(object sender, RibbonControlEventArgs e)
        {
            AlignTool.AlignHorizontalVertical();
        }

        private void beginArrowToggle_Click(object sender, RibbonControlEventArgs e)
        {
            ShapeTool.BeginArrowToggle();
        }

        private void beginArrowChangeSize_Click(object sender, RibbonControlEventArgs e)
        {
            ShapeTool.BeginArrowChangeSize(1);
        }

        private void beginArrowSizeDown_Click(object sender, RibbonControlEventArgs e)
        {
            ShapeTool.BeginArrowChangeSize(-1);
        }

        private void endArrowToggle_Click(object sender, RibbonControlEventArgs e)
        {
            ShapeTool.EndArrowToggle();
        }

        private void endArrowSizeUp_Click(object sender, RibbonControlEventArgs e)
        {
            ShapeTool.EndArrowChangeSize(1);
        }

        private void endArrowSizeDown_Click(object sender, RibbonControlEventArgs e)
        {
            ShapeTool.EndArrowChangeSize(-1);
        }

        private void connectShapeByLine_Click(object sender, RibbonControlEventArgs e)
        {
            ShapeTool.ConnectShapesByLine();
        }

        private void curveOfEquation_Click(object sender, RibbonControlEventArgs e)
        {
            var equationX = Globals.Ribbons.EdoliRibbon.curveOfEquationX.Text;
            var equationY = Globals.Ribbons.EdoliRibbon.curveOfEquationY.Text;
            var startValue = Globals.Ribbons.EdoliRibbon.curveStart.Text;
            var endValue = Globals.Ribbons.EdoliRibbon.curveEnd.Text;
            ShapeTool.AddCurveOfExpression(equationX, equationY, startValue, endValue);
        }

        private void polylineOfEquation_Click(object sender, RibbonControlEventArgs e)
        {
            var equationX = Globals.Ribbons.EdoliRibbon.curveOfEquationX.Text;
            var equationY = Globals.Ribbons.EdoliRibbon.curveOfEquationY.Text;
            var startValue = Globals.Ribbons.EdoliRibbon.curveStart.Text;
            var endValue = Globals.Ribbons.EdoliRibbon.curveEnd.Text;
            ShapeTool.AddPolylineOfExpression(equationX, equationY, startValue, endValue);
        }

        private void invertImage_Click(object sender, RibbonControlEventArgs e)
        {
            ImageTool.InvertImage();
        }

        private void trimImage_Click(object sender, RibbonControlEventArgs e)
        {
            ImageTool.TrimImage();
        }

        private void resizeWidth_Click(object sender, RibbonControlEventArgs e)
        {
            AlignTool.MatchWidth();
        }

        private void resizeHeight_Click(object sender, RibbonControlEventArgs e)
        {
            AlignTool.MatchHeight();
        }

        private void followAnimation_Click(object sender, RibbonControlEventArgs e)
        {
            AnimationTool.FollowAnimation();
        }
    }
}
