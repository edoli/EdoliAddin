namespace EdoliAddIn
{
    partial class EdoliRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public EdoliRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.shape = this.Factory.CreateRibbonGroup();
            this.beginArrowToggle = this.Factory.CreateRibbonButton();
            this.beginArrowSizeUp = this.Factory.CreateRibbonButton();
            this.beginArrowSizeDown = this.Factory.CreateRibbonButton();
            this.endArrowToggle = this.Factory.CreateRibbonButton();
            this.endArrowSizeUp = this.Factory.CreateRibbonButton();
            this.endArrowSizeDown = this.Factory.CreateRibbonButton();
            this.connectShapeByLine = this.Factory.CreateRibbonButton();
            this.invertImage = this.Factory.CreateRibbonButton();
            this.trimImage = this.Factory.CreateRibbonButton();
            this.alignGroup = this.Factory.CreateRibbonGroup();
            this.labelBottom = this.Factory.CreateRibbonButton();
            this.labelTop = this.Factory.CreateRibbonButton();
            this.transpose = this.Factory.CreateRibbonButton();
            this.labelLeft = this.Factory.CreateRibbonButton();
            this.labelRight = this.Factory.CreateRibbonButton();
            this.groupLabel = this.Factory.CreateRibbonButton();
            this.alignPrevSlide = this.Factory.CreateRibbonButton();
            this.swapCycle = this.Factory.CreateRibbonButton();
            this.snapDownRight = this.Factory.CreateRibbonButton();
            this.alignNextSlide = this.Factory.CreateRibbonButton();
            this.swapCycleReverse = this.Factory.CreateRibbonButton();
            this.snapUpRight = this.Factory.CreateRibbonButton();
            this.alignGrid = this.Factory.CreateRibbonButton();
            this.resizeWidth = this.Factory.CreateRibbonButton();
            this.resizeHeight = this.Factory.CreateRibbonButton();
            this.groupGrid = this.Factory.CreateRibbonGroup();
            this.grid = this.Factory.CreateRibbonButton();
            this.gridPadding = this.Factory.CreateRibbonEditBox();
            this.gridNumColumn = this.Factory.CreateRibbonEditBox();
            this.equationGroup = this.Factory.CreateRibbonGroup();
            this.box1 = this.Factory.CreateRibbonBox();
            this.curveOfEquation = this.Factory.CreateRibbonButton();
            this.polylineOfEquation = this.Factory.CreateRibbonButton();
            this.curveStart = this.Factory.CreateRibbonEditBox();
            this.curveEnd = this.Factory.CreateRibbonEditBox();
            this.checkBoxNormalizeEqShape = this.Factory.CreateRibbonCheckBox();
            this.box2 = this.Factory.CreateRibbonBox();
            this.curveOfEquationX = this.Factory.CreateRibbonEditBox();
            this.curveOfEquationY = this.Factory.CreateRibbonEditBox();
            this.animationGroup = this.Factory.CreateRibbonGroup();
            this.editBoxName = this.Factory.CreateRibbonEditBox();
            this.followAnimation = this.Factory.CreateRibbonButton();
            this.alignLines = this.Factory.CreateRibbonButton();
            this.trimLines = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.shape.SuspendLayout();
            this.alignGroup.SuspendLayout();
            this.groupGrid.SuspendLayout();
            this.equationGroup.SuspendLayout();
            this.box1.SuspendLayout();
            this.box2.SuspendLayout();
            this.animationGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabHome";
            this.tab1.Groups.Add(this.shape);
            this.tab1.Groups.Add(this.alignGroup);
            this.tab1.Groups.Add(this.groupGrid);
            this.tab1.Groups.Add(this.equationGroup);
            this.tab1.Groups.Add(this.animationGroup);
            this.tab1.Label = "TabHome";
            this.tab1.Name = "tab1";
            // 
            // shape
            // 
            this.shape.Items.Add(this.beginArrowToggle);
            this.shape.Items.Add(this.beginArrowSizeUp);
            this.shape.Items.Add(this.beginArrowSizeDown);
            this.shape.Items.Add(this.endArrowToggle);
            this.shape.Items.Add(this.endArrowSizeUp);
            this.shape.Items.Add(this.endArrowSizeDown);
            this.shape.Items.Add(this.connectShapeByLine);
            this.shape.Items.Add(this.invertImage);
            this.shape.Items.Add(this.trimImage);
            this.shape.Items.Add(this.alignLines);
            this.shape.Items.Add(this.trimLines);
            this.shape.Label = "Shape";
            this.shape.Name = "shape";
            // 
            // beginArrowToggle
            // 
            this.beginArrowToggle.Image = global::EdoliAddIn.Properties.Resources.icon_begin_arrow_toggle;
            this.beginArrowToggle.Label = "button1";
            this.beginArrowToggle.Name = "beginArrowToggle";
            this.beginArrowToggle.ScreenTip = "Begin arrow toggle";
            this.beginArrowToggle.ShowImage = true;
            this.beginArrowToggle.ShowLabel = false;
            this.beginArrowToggle.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.beginArrowToggle_Click);
            // 
            // beginArrowSizeUp
            // 
            this.beginArrowSizeUp.Image = global::EdoliAddIn.Properties.Resources.icon_begin_arrow_size_up;
            this.beginArrowSizeUp.Label = "button1";
            this.beginArrowSizeUp.Name = "beginArrowSizeUp";
            this.beginArrowSizeUp.ShowImage = true;
            this.beginArrowSizeUp.ShowLabel = false;
            this.beginArrowSizeUp.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.beginArrowChangeSize_Click);
            // 
            // beginArrowSizeDown
            // 
            this.beginArrowSizeDown.Image = global::EdoliAddIn.Properties.Resources.icon_begin_arrow_size_down;
            this.beginArrowSizeDown.Label = "button1";
            this.beginArrowSizeDown.Name = "beginArrowSizeDown";
            this.beginArrowSizeDown.ShowImage = true;
            this.beginArrowSizeDown.ShowLabel = false;
            this.beginArrowSizeDown.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.beginArrowSizeDown_Click);
            // 
            // endArrowToggle
            // 
            this.endArrowToggle.Image = global::EdoliAddIn.Properties.Resources.icon_end_arrow_toggle;
            this.endArrowToggle.Label = "button1";
            this.endArrowToggle.Name = "endArrowToggle";
            this.endArrowToggle.ShowImage = true;
            this.endArrowToggle.ShowLabel = false;
            this.endArrowToggle.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.endArrowToggle_Click);
            // 
            // endArrowSizeUp
            // 
            this.endArrowSizeUp.Image = global::EdoliAddIn.Properties.Resources.icon_end_arrow_size_up;
            this.endArrowSizeUp.Label = "button1";
            this.endArrowSizeUp.Name = "endArrowSizeUp";
            this.endArrowSizeUp.ShowImage = true;
            this.endArrowSizeUp.ShowLabel = false;
            this.endArrowSizeUp.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.endArrowSizeUp_Click);
            // 
            // endArrowSizeDown
            // 
            this.endArrowSizeDown.Image = global::EdoliAddIn.Properties.Resources.icon_end_arrow_size_down;
            this.endArrowSizeDown.Label = "button1";
            this.endArrowSizeDown.Name = "endArrowSizeDown";
            this.endArrowSizeDown.ShowImage = true;
            this.endArrowSizeDown.ShowLabel = false;
            this.endArrowSizeDown.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.endArrowSizeDown_Click);
            // 
            // connectShapeByLine
            // 
            this.connectShapeByLine.Image = global::EdoliAddIn.Properties.Resources.icon_connect_shape_by_lines;
            this.connectShapeByLine.Label = "Connect Line";
            this.connectShapeByLine.Name = "connectShapeByLine";
            this.connectShapeByLine.ScreenTip = "Connect Line";
            this.connectShapeByLine.ShowImage = true;
            this.connectShapeByLine.ShowLabel = false;
            this.connectShapeByLine.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.connectShapeByLine_Click);
            // 
            // invertImage
            // 
            this.invertImage.Image = global::EdoliAddIn.Properties.Resources.icon_image_invert;
            this.invertImage.Label = "Invert";
            this.invertImage.Name = "invertImage";
            this.invertImage.ShowImage = true;
            this.invertImage.ShowLabel = false;
            this.invertImage.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.invertImage_Click);
            // 
            // trimImage
            // 
            this.trimImage.Image = global::EdoliAddIn.Properties.Resources.icon_image_trim;
            this.trimImage.Label = "Trim";
            this.trimImage.Name = "trimImage";
            this.trimImage.ShowImage = true;
            this.trimImage.ShowLabel = false;
            this.trimImage.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.trimImage_Click);
            // 
            // alignGroup
            // 
            this.alignGroup.Items.Add(this.labelBottom);
            this.alignGroup.Items.Add(this.labelTop);
            this.alignGroup.Items.Add(this.transpose);
            this.alignGroup.Items.Add(this.labelLeft);
            this.alignGroup.Items.Add(this.labelRight);
            this.alignGroup.Items.Add(this.groupLabel);
            this.alignGroup.Items.Add(this.alignPrevSlide);
            this.alignGroup.Items.Add(this.swapCycle);
            this.alignGroup.Items.Add(this.snapDownRight);
            this.alignGroup.Items.Add(this.alignNextSlide);
            this.alignGroup.Items.Add(this.swapCycleReverse);
            this.alignGroup.Items.Add(this.snapUpRight);
            this.alignGroup.Items.Add(this.alignGrid);
            this.alignGroup.Items.Add(this.resizeWidth);
            this.alignGroup.Items.Add(this.resizeHeight);
            this.alignGroup.Label = "Align";
            this.alignGroup.Name = "alignGroup";
            // 
            // labelBottom
            // 
            this.labelBottom.Image = global::EdoliAddIn.Properties.Resources.icon_label_bottom;
            this.labelBottom.Label = "LabelBottom";
            this.labelBottom.Name = "labelBottom";
            this.labelBottom.ScreenTip = "Label bottom";
            this.labelBottom.ShowImage = true;
            this.labelBottom.ShowLabel = false;
            this.labelBottom.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.labelBottom_Click);
            // 
            // labelTop
            // 
            this.labelTop.Image = global::EdoliAddIn.Properties.Resources.icon_label_top;
            this.labelTop.Label = "LabelTop";
            this.labelTop.Name = "labelTop";
            this.labelTop.ScreenTip = "Label top";
            this.labelTop.ShowImage = true;
            this.labelTop.ShowLabel = false;
            this.labelTop.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.labelTop_Click);
            // 
            // transpose
            // 
            this.transpose.Image = global::EdoliAddIn.Properties.Resources.icon_transpose;
            this.transpose.Label = "Transpose";
            this.transpose.Name = "transpose";
            this.transpose.ScreenTip = "Transpose";
            this.transpose.ShowImage = true;
            this.transpose.ShowLabel = false;
            this.transpose.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.transpose_Click);
            // 
            // labelLeft
            // 
            this.labelLeft.Image = global::EdoliAddIn.Properties.Resources.icon_label_left;
            this.labelLeft.Label = "LabelLeft";
            this.labelLeft.Name = "labelLeft";
            this.labelLeft.ScreenTip = "Label left";
            this.labelLeft.ShowImage = true;
            this.labelLeft.ShowLabel = false;
            this.labelLeft.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.labelLeft_Click);
            // 
            // labelRight
            // 
            this.labelRight.Image = global::EdoliAddIn.Properties.Resources.icon_label_right;
            this.labelRight.Label = "LabelRight";
            this.labelRight.Name = "labelRight";
            this.labelRight.ScreenTip = "Label right";
            this.labelRight.ShowImage = true;
            this.labelRight.ShowLabel = false;
            this.labelRight.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.labelRight_Click);
            // 
            // groupLabel
            // 
            this.groupLabel.Image = global::EdoliAddIn.Properties.Resources.icon_label_group;
            this.groupLabel.Label = "Group Label";
            this.groupLabel.Name = "groupLabel";
            this.groupLabel.ScreenTip = "Group Label";
            this.groupLabel.ShowImage = true;
            this.groupLabel.ShowLabel = false;
            this.groupLabel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.groupLabel_Click);
            // 
            // alignPrevSlide
            // 
            this.alignPrevSlide.Image = global::EdoliAddIn.Properties.Resources.icon_align_prev_slide;
            this.alignPrevSlide.Label = "Align";
            this.alignPrevSlide.Name = "alignPrevSlide";
            this.alignPrevSlide.ScreenTip = "Align with previous slide";
            this.alignPrevSlide.ShowImage = true;
            this.alignPrevSlide.ShowLabel = false;
            this.alignPrevSlide.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.alignPrevSlide_Click);
            // 
            // swapCycle
            // 
            this.swapCycle.Image = global::EdoliAddIn.Properties.Resources.icon_swap_cycle;
            this.swapCycle.Label = "Swap cycle";
            this.swapCycle.Name = "swapCycle";
            this.swapCycle.ScreenTip = "Swap cycle";
            this.swapCycle.ShowImage = true;
            this.swapCycle.ShowLabel = false;
            this.swapCycle.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.swapCycle_Click);
            // 
            // snapDownRight
            // 
            this.snapDownRight.Image = global::EdoliAddIn.Properties.Resources.icon_snap_diag_downright;
            this.snapDownRight.Label = "Snap down right";
            this.snapDownRight.Name = "snapDownRight";
            this.snapDownRight.ScreenTip = "Snap down right";
            this.snapDownRight.ShowImage = true;
            this.snapDownRight.ShowLabel = false;
            this.snapDownRight.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.snapDownRight_Click);
            // 
            // alignNextSlide
            // 
            this.alignNextSlide.Image = global::EdoliAddIn.Properties.Resources.icon_align_next_slide;
            this.alignNextSlide.Label = "Align next slide";
            this.alignNextSlide.Name = "alignNextSlide";
            this.alignNextSlide.ScreenTip = "Align with next slide";
            this.alignNextSlide.ShowImage = true;
            this.alignNextSlide.ShowLabel = false;
            this.alignNextSlide.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.alignNextSlide_Click);
            // 
            // swapCycleReverse
            // 
            this.swapCycleReverse.Image = global::EdoliAddIn.Properties.Resources.icon_swap_cycle_reverse;
            this.swapCycleReverse.Label = "Swap cycle reverse";
            this.swapCycleReverse.Name = "swapCycleReverse";
            this.swapCycleReverse.ScreenTip = "Swap cycle reverse";
            this.swapCycleReverse.ShowImage = true;
            this.swapCycleReverse.ShowLabel = false;
            this.swapCycleReverse.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.swapCycleReverse_Click);
            // 
            // snapUpRight
            // 
            this.snapUpRight.Image = global::EdoliAddIn.Properties.Resources.icon_snap_diag_upright;
            this.snapUpRight.Label = "Snap up right";
            this.snapUpRight.Name = "snapUpRight";
            this.snapUpRight.ScreenTip = "Snap up right";
            this.snapUpRight.ShowImage = true;
            this.snapUpRight.ShowLabel = false;
            this.snapUpRight.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.snapUpRight_Click);
            // 
            // alignGrid
            // 
            this.alignGrid.Image = global::EdoliAddIn.Properties.Resources.icon_align_grid;
            this.alignGrid.Label = "Align grid";
            this.alignGrid.Name = "alignGrid";
            this.alignGrid.ScreenTip = "Align grid";
            this.alignGrid.ShowImage = true;
            this.alignGrid.ShowLabel = false;
            this.alignGrid.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.alignGrid_Click);
            // 
            // resizeWidth
            // 
            this.resizeWidth.Image = global::EdoliAddIn.Properties.Resources.icon_resize_width;
            this.resizeWidth.Label = "Resize width";
            this.resizeWidth.Name = "resizeWidth";
            this.resizeWidth.ShowImage = true;
            this.resizeWidth.ShowLabel = false;
            this.resizeWidth.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.resizeWidth_Click);
            // 
            // resizeHeight
            // 
            this.resizeHeight.Image = global::EdoliAddIn.Properties.Resources.icon_resize_height;
            this.resizeHeight.Label = "Resize height";
            this.resizeHeight.Name = "resizeHeight";
            this.resizeHeight.ShowImage = true;
            this.resizeHeight.ShowLabel = false;
            this.resizeHeight.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.resizeHeight_Click);
            // 
            // groupGrid
            // 
            this.groupGrid.Items.Add(this.grid);
            this.groupGrid.Items.Add(this.gridPadding);
            this.groupGrid.Items.Add(this.gridNumColumn);
            this.groupGrid.Label = "Grid";
            this.groupGrid.Name = "groupGrid";
            // 
            // grid
            // 
            this.grid.Image = global::EdoliAddIn.Properties.Resources.icon_grid;
            this.grid.Label = "Grid";
            this.grid.Name = "grid";
            this.grid.ScreenTip = "Grid";
            this.grid.ShowImage = true;
            this.grid.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.grid_Click);
            // 
            // gridPadding
            // 
            this.gridPadding.Label = "Padding";
            this.gridPadding.Name = "gridPadding";
            this.gridPadding.ScreenTip = "Padding";
            this.gridPadding.SizeString = "000";
            this.gridPadding.Text = "0";
            // 
            // gridNumColumn
            // 
            this.gridNumColumn.Label = "Column";
            this.gridNumColumn.Name = "gridNumColumn";
            this.gridNumColumn.ScreenTip = "Column";
            this.gridNumColumn.SizeString = "000";
            this.gridNumColumn.Text = "0";
            // 
            // equationGroup
            // 
            this.equationGroup.Items.Add(this.box1);
            this.equationGroup.Items.Add(this.box2);
            this.equationGroup.Label = "Equation";
            this.equationGroup.Name = "equationGroup";
            // 
            // box1
            // 
            this.box1.Items.Add(this.curveOfEquation);
            this.box1.Items.Add(this.polylineOfEquation);
            this.box1.Items.Add(this.curveStart);
            this.box1.Items.Add(this.curveEnd);
            this.box1.Items.Add(this.checkBoxNormalizeEqShape);
            this.box1.Name = "box1";
            // 
            // curveOfEquation
            // 
            this.curveOfEquation.Image = global::EdoliAddIn.Properties.Resources.icon_curve_of_equation;
            this.curveOfEquation.Label = "Curve of equation";
            this.curveOfEquation.Name = "curveOfEquation";
            this.curveOfEquation.ScreenTip = "Curve of equation";
            this.curveOfEquation.ShowImage = true;
            this.curveOfEquation.ShowLabel = false;
            this.curveOfEquation.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.curveOfEquation_Click);
            // 
            // polylineOfEquation
            // 
            this.polylineOfEquation.Image = global::EdoliAddIn.Properties.Resources.icon_polyline_of_equation;
            this.polylineOfEquation.Label = "Polyline of equation";
            this.polylineOfEquation.Name = "polylineOfEquation";
            this.polylineOfEquation.ScreenTip = "Polyline of equation";
            this.polylineOfEquation.ShowImage = true;
            this.polylineOfEquation.ShowLabel = false;
            this.polylineOfEquation.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.polylineOfEquation_Click);
            // 
            // curveStart
            // 
            this.curveStart.Label = "editBox1";
            this.curveStart.Name = "curveStart";
            this.curveStart.ShowLabel = false;
            this.curveStart.SizeString = "00000";
            this.curveStart.Text = "0";
            // 
            // curveEnd
            // 
            this.curveEnd.Label = "-";
            this.curveEnd.Name = "curveEnd";
            this.curveEnd.SizeString = "00000";
            this.curveEnd.Text = "1";
            // 
            // checkBoxNormalizeEqShape
            // 
            this.checkBoxNormalizeEqShape.Label = "Norm";
            this.checkBoxNormalizeEqShape.Name = "checkBoxNormalizeEqShape";
            // 
            // box2
            // 
            this.box2.Items.Add(this.curveOfEquationX);
            this.box2.Items.Add(this.curveOfEquationY);
            this.box2.Name = "box2";
            // 
            // curveOfEquationX
            // 
            this.curveOfEquationX.Label = "X";
            this.curveOfEquationX.Name = "curveOfEquationX";
            this.curveOfEquationX.Text = null;
            // 
            // curveOfEquationY
            // 
            this.curveOfEquationY.Label = "Y";
            this.curveOfEquationY.Name = "curveOfEquationY";
            this.curveOfEquationY.Text = null;
            // 
            // animationGroup
            // 
            this.animationGroup.Items.Add(this.editBoxName);
            this.animationGroup.Items.Add(this.followAnimation);
            this.animationGroup.Label = "Animation";
            this.animationGroup.Name = "animationGroup";
            // 
            // editBoxName
            // 
            this.editBoxName.Label = "Animation Name";
            this.editBoxName.Name = "editBoxName";
            this.editBoxName.ScreenTip = "Animation Name";
            this.editBoxName.ShowLabel = false;
            this.editBoxName.SizeString = "00000000";
            this.editBoxName.Text = null;
            this.editBoxName.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.editBoxName_TextChanged);
            // 
            // followAnimation
            // 
            this.followAnimation.Label = "Follow";
            this.followAnimation.Name = "followAnimation";
            this.followAnimation.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.followAnimation_Click);
            // 
            // alignLines
            // 
            this.alignLines.Image = global::EdoliAddIn.Properties.Resources.icon_image_trim;
            this.alignLines.Label = "AlignLines";
            this.alignLines.Name = "alignLines";
            this.alignLines.ShowImage = true;
            this.alignLines.ShowLabel = false;
            this.alignLines.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.alignLines_Click);
            // 
            // trimLines
            // 
            this.trimLines.Image = global::EdoliAddIn.Properties.Resources.icon_image_trim;
            this.trimLines.Label = "TrimLines";
            this.trimLines.Name = "trimLines";
            this.trimLines.ShowImage = true;
            this.trimLines.ShowLabel = false;
            this.trimLines.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.trimLines_Click);
            // 
            // EdoliRibbon
            // 
            this.Name = "EdoliRibbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.EdoliRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.shape.ResumeLayout(false);
            this.shape.PerformLayout();
            this.alignGroup.ResumeLayout(false);
            this.alignGroup.PerformLayout();
            this.groupGrid.ResumeLayout(false);
            this.groupGrid.PerformLayout();
            this.equationGroup.ResumeLayout(false);
            this.equationGroup.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.box2.ResumeLayout(false);
            this.box2.PerformLayout();
            this.animationGroup.ResumeLayout(false);
            this.animationGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup alignGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton grid;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton labelBottom;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton labelTop;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton labelLeft;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton labelRight;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup animationGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBoxName;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton transpose;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton groupLabel;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton alignPrevSlide;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox gridNumColumn;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox gridPadding;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupGrid;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton swapCycle;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton swapCycleReverse;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton snapDownRight;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton snapUpRight;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton alignNextSlide;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup shape;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton beginArrowToggle;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton beginArrowSizeUp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton beginArrowSizeDown;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton endArrowToggle;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton endArrowSizeUp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton endArrowSizeDown;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton alignGrid;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton connectShapeByLine;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton curveOfEquation;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup equationGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton polylineOfEquation;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton invertImage;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton trimImage;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox curveOfEquationX;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox curveOfEquationY;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox curveStart;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox curveEnd;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box2;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBoxNormalizeEqShape;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton resizeWidth;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton resizeHeight;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton followAnimation;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton alignLines;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton trimLines;
    }

    partial class ThisRibbonCollection
    {
        internal EdoliRibbon EdoliRibbon
        {
            get { return this.GetRibbon<EdoliRibbon>(); }
        }
    }
}
