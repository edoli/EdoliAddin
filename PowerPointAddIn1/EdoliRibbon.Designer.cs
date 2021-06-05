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
            this.alignGroup = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.labelBottom = this.Factory.CreateRibbonButton();
            this.labelTop = this.Factory.CreateRibbonButton();
            this.labelLeft = this.Factory.CreateRibbonButton();
            this.labelRight = this.Factory.CreateRibbonButton();
            this.animationGroup = this.Factory.CreateRibbonGroup();
            this.editBoxName = this.Factory.CreateRibbonEditBox();
            this.tab1.SuspendLayout();
            this.alignGroup.SuspendLayout();
            this.animationGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabHome";
            this.tab1.Groups.Add(this.alignGroup);
            this.tab1.Groups.Add(this.animationGroup);
            this.tab1.Label = "TabHome";
            this.tab1.Name = "tab1";
            // 
            // alignGroup
            // 
            this.alignGroup.Items.Add(this.button1);
            this.alignGroup.Items.Add(this.labelBottom);
            this.alignGroup.Items.Add(this.labelTop);
            this.alignGroup.Items.Add(this.labelLeft);
            this.alignGroup.Items.Add(this.labelRight);
            this.alignGroup.Label = "Align";
            this.alignGroup.Name = "alignGroup";
            // 
            // button1
            // 
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.Label = "button1";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.ShowLabel = false;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
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
            this.alignGroup.ResumeLayout(false);
            this.alignGroup.PerformLayout();
            this.animationGroup.ResumeLayout(false);
            this.animationGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup alignGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton labelBottom;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton labelTop;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton labelLeft;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton labelRight;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup animationGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBoxName;
    }

    partial class ThisRibbonCollection
    {
        internal EdoliRibbon EdoliRibbon
        {
            get { return this.GetRibbon<EdoliRibbon>(); }
        }
    }
}
