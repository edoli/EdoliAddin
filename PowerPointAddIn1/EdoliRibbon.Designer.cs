namespace PowerPointAddIn1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(EdoliRibbon));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.shape = this.Factory.CreateRibbonGroup();
            this.beginArrowToggle = this.Factory.CreateRibbonButton();
            this.beginArrowSizeUp = this.Factory.CreateRibbonButton();
            this.beginArrowSizeDown = this.Factory.CreateRibbonButton();
            this.endArrowToggle = this.Factory.CreateRibbonButton();
            this.endArrowSizeUp = this.Factory.CreateRibbonButton();
            this.endArrowSizeDown = this.Factory.CreateRibbonButton();
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
            this.groupGrid = this.Factory.CreateRibbonGroup();
            this.grid = this.Factory.CreateRibbonButton();
            this.gridPadding = this.Factory.CreateRibbonEditBox();
            this.gridNumColumn = this.Factory.CreateRibbonEditBox();
            this.animationGroup = this.Factory.CreateRibbonGroup();
            this.editBoxName = this.Factory.CreateRibbonEditBox();
            this.tab1.SuspendLayout();
            this.shape.SuspendLayout();
            this.alignGroup.SuspendLayout();
            this.groupGrid.SuspendLayout();
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
            this.shape.Label = "Shape";
            this.shape.Name = "shape";
            // 
            // beginArrowToggle
            // 
            this.beginArrowToggle.Image = global::PowerPointAddIn1.Properties.Resources.icon_begin_arrow_toggle;
            this.beginArrowToggle.Label = "button1";
            this.beginArrowToggle.Name = "beginArrowToggle";
            this.beginArrowToggle.ScreenTip = "Begin arrow toggle";
            this.beginArrowToggle.ShowImage = true;
            this.beginArrowToggle.ShowLabel = false;
            this.beginArrowToggle.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.beginArrowToggle_Click);
            // 
            // beginArrowSizeUp
            // 
            this.beginArrowSizeUp.Image = global::PowerPointAddIn1.Properties.Resources.icon_begin_arrow_size_up;
            this.beginArrowSizeUp.Label = "button1";
            this.beginArrowSizeUp.Name = "beginArrowSizeUp";
            this.beginArrowSizeUp.ShowImage = true;
            this.beginArrowSizeUp.ShowLabel = false;
            this.beginArrowSizeUp.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.beginArrowChangeSize_Click);
            // 
            // beginArrowSizeDown
            // 
            this.beginArrowSizeDown.Image = global::PowerPointAddIn1.Properties.Resources.icon_begin_arrow_size_down;
            this.beginArrowSizeDown.Label = "button1";
            this.beginArrowSizeDown.Name = "beginArrowSizeDown";
            this.beginArrowSizeDown.ShowImage = true;
            this.beginArrowSizeDown.ShowLabel = false;
            this.beginArrowSizeDown.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.beginArrowSizeDown_Click);
            // 
            // endArrowToggle
            // 
            this.endArrowToggle.Image = global::PowerPointAddIn1.Properties.Resources.icon_end_arrow_toggle;
            this.endArrowToggle.Label = "button1";
            this.endArrowToggle.Name = "endArrowToggle";
            this.endArrowToggle.ShowImage = true;
            this.endArrowToggle.ShowLabel = false;
            this.endArrowToggle.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.endArrowToggle_Click);
            // 
            // endArrowSizeUp
            // 
            this.endArrowSizeUp.Image = global::PowerPointAddIn1.Properties.Resources.icon_end_arrow_size_up;
            this.endArrowSizeUp.Label = "button1";
            this.endArrowSizeUp.Name = "endArrowSizeUp";
            this.endArrowSizeUp.ShowImage = true;
            this.endArrowSizeUp.ShowLabel = false;
            this.endArrowSizeUp.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.endArrowSizeUp_Click);
            // 
            // endArrowSizeDown
            // 
            this.endArrowSizeDown.Image = global::PowerPointAddIn1.Properties.Resources.icon_end_arrow_size_down;
            this.endArrowSizeDown.Label = "button1";
            this.endArrowSizeDown.Name = "endArrowSizeDown";
            this.endArrowSizeDown.ShowImage = true;
            this.endArrowSizeDown.ShowLabel = false;
            this.endArrowSizeDown.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.endArrowSizeDown_Click);
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
            this.alignGroup.Label = "Align";
            this.alignGroup.Name = "alignGroup";
            // 
            // labelBottom
            // 
            this.labelBottom.Image = ((System.Drawing.Image)(resources.GetObject("labelBottom.Image")));
            this.labelBottom.Label = "LabelBottom";
            this.labelBottom.Name = "labelBottom";
            this.labelBottom.ScreenTip = "Label bottom";
            this.labelBottom.ShowImage = true;
            this.labelBottom.ShowLabel = false;
            this.labelBottom.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.labelBottom_Click);
            // 
            // labelTop
            // 
            this.labelTop.Image = ((System.Drawing.Image)(resources.GetObject("labelTop.Image")));
            this.labelTop.Label = "LabelTop";
            this.labelTop.Name = "labelTop";
            this.labelTop.ScreenTip = "Label top";
            this.labelTop.ShowImage = true;
            this.labelTop.ShowLabel = false;
            this.labelTop.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.labelTop_Click);
            // 
            // transpose
            // 
            this.transpose.Image = ((System.Drawing.Image)(resources.GetObject("transpose.Image")));
            this.transpose.Label = "Transpose";
            this.transpose.Name = "transpose";
            this.transpose.ScreenTip = "Transpose";
            this.transpose.ShowImage = true;
            this.transpose.ShowLabel = false;
            this.transpose.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.transpose_Click);
            // 
            // labelLeft
            // 
            this.labelLeft.Image = ((System.Drawing.Image)(resources.GetObject("labelLeft.Image")));
            this.labelLeft.Label = "LabelLeft";
            this.labelLeft.Name = "labelLeft";
            this.labelLeft.ScreenTip = "Label left";
            this.labelLeft.ShowImage = true;
            this.labelLeft.ShowLabel = false;
            this.labelLeft.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.labelLeft_Click);
            // 
            // labelRight
            // 
            this.labelRight.Image = ((System.Drawing.Image)(resources.GetObject("labelRight.Image")));
            this.labelRight.Label = "LabelRight";
            this.labelRight.Name = "labelRight";
            this.labelRight.ScreenTip = "Label right";
            this.labelRight.ShowImage = true;
            this.labelRight.ShowLabel = false;
            this.labelRight.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.labelRight_Click);
            // 
            // groupLabel
            // 
            this.groupLabel.Image = ((System.Drawing.Image)(resources.GetObject("groupLabel.Image")));
            this.groupLabel.Label = "Group Label";
            this.groupLabel.Name = "groupLabel";
            this.groupLabel.ScreenTip = "Group Label";
            this.groupLabel.ShowImage = true;
            this.groupLabel.ShowLabel = false;
            this.groupLabel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.groupLabel_Click);
            // 
            // alignPrevSlide
            // 
            this.alignPrevSlide.Image = ((System.Drawing.Image)(resources.GetObject("alignPrevSlide.Image")));
            this.alignPrevSlide.Label = "Align";
            this.alignPrevSlide.Name = "alignPrevSlide";
            this.alignPrevSlide.ScreenTip = "Align with previous slide";
            this.alignPrevSlide.ShowImage = true;
            this.alignPrevSlide.ShowLabel = false;
            this.alignPrevSlide.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.alignPrevSlide_Click);
            // 
            // swapCycle
            // 
            this.swapCycle.Image = ((System.Drawing.Image)(resources.GetObject("swapCycle.Image")));
            this.swapCycle.Label = "Swap cycle";
            this.swapCycle.Name = "swapCycle";
            this.swapCycle.ScreenTip = "Swap cycle";
            this.swapCycle.ShowImage = true;
            this.swapCycle.ShowLabel = false;
            this.swapCycle.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.swapCycle_Click);
            // 
            // snapDownRight
            // 
            this.snapDownRight.Image = ((System.Drawing.Image)(resources.GetObject("snapDownRight.Image")));
            this.snapDownRight.Label = "Snap down right";
            this.snapDownRight.Name = "snapDownRight";
            this.snapDownRight.ScreenTip = "Snap down right";
            this.snapDownRight.ShowImage = true;
            this.snapDownRight.ShowLabel = false;
            this.snapDownRight.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.snapDownRight_Click);
            // 
            // alignNextSlide
            // 
            this.alignNextSlide.Image = ((System.Drawing.Image)(resources.GetObject("alignNextSlide.Image")));
            this.alignNextSlide.Label = "Align next slide";
            this.alignNextSlide.Name = "alignNextSlide";
            this.alignNextSlide.ScreenTip = "Align with next slide";
            this.alignNextSlide.ShowImage = true;
            this.alignNextSlide.ShowLabel = false;
            this.alignNextSlide.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.alignNextSlide_Click);
            // 
            // swapCycleReverse
            // 
            this.swapCycleReverse.Image = ((System.Drawing.Image)(resources.GetObject("swapCycleReverse.Image")));
            this.swapCycleReverse.Label = "Swap cycle reverse";
            this.swapCycleReverse.Name = "swapCycleReverse";
            this.swapCycleReverse.ScreenTip = "Swap cycle reverse";
            this.swapCycleReverse.ShowImage = true;
            this.swapCycleReverse.ShowLabel = false;
            this.swapCycleReverse.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.swapCycleReverse_Click);
            // 
            // snapUpRight
            // 
            this.snapUpRight.Image = ((System.Drawing.Image)(resources.GetObject("snapUpRight.Image")));
            this.snapUpRight.Label = "Snap up right";
            this.snapUpRight.Name = "snapUpRight";
            this.snapUpRight.ScreenTip = "Snap up right";
            this.snapUpRight.ShowImage = true;
            this.snapUpRight.ShowLabel = false;
            this.snapUpRight.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.snapUpRight_Click);
            // 
            // alignGrid
            // 
            this.alignGrid.Image = ((System.Drawing.Image)(resources.GetObject("alignGrid.Image")));
            this.alignGrid.Label = "Align grid";
            this.alignGrid.Name = "alignGrid";
            this.alignGrid.ScreenTip = "Align grid";
            this.alignGrid.ShowImage = true;
            this.alignGrid.ShowLabel = false;
            this.alignGrid.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.alignGrid_Click);
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
            this.grid.Image = ((System.Drawing.Image)(resources.GetObject("grid.Image")));
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
            this.gridPadding.Text = "0";
            // 
            // gridNumColumn
            // 
            this.gridNumColumn.Label = "Column";
            this.gridNumColumn.Name = "gridNumColumn";
            this.gridNumColumn.ScreenTip = "Column";
            this.gridNumColumn.Text = "0";
            // 
            // animationGroup
            // 
            this.animationGroup.Items.Add(this.editBoxName);
            this.animationGroup.Label = "Animation";
            this.animationGroup.Name = "animationGroup";
            // 
            // editBoxName
            // 
            this.editBoxName.Label = "Animation Name";
            this.editBoxName.Name = "editBoxName";
            this.editBoxName.ScreenTip = "Animation Name";
            this.editBoxName.ShowLabel = false;
            this.editBoxName.Text = null;
            this.editBoxName.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.editBoxName_TextChanged);
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
    }

    partial class ThisRibbonCollection
    {
        internal EdoliRibbon EdoliRibbon
        {
            get { return this.GetRibbon<EdoliRibbon>(); }
        }
    }
}
