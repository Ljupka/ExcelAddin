using System.Windows.Forms.Layout;
using System.Diagnostics;

namespace ExcelAddIn2
{
    partial class MyRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MyRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Verwendete Ressourcen bereinigen.
        /// </summary>
        /// <param name="disposing">True, wenn verwaltete Ressourcen gelöscht werden sollen; andernfalls False.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Vom Komponenten-Designer generierter Code

        /// <summary>
        /// Erforderliche Methode für Designerunterstützung -
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InitializeComponent()
        {
            this.sent = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.ClassifySelection = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.Classify = this.Factory.CreateRibbonButton();
            this.Send = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.checkBox2 = this.Factory.CreateRibbonCheckBox();
            this.checkBox3 = this.Factory.CreateRibbonCheckBox();
            this.checkBox4 = this.Factory.CreateRibbonCheckBox();
            this.checkBox5 = this.Factory.CreateRibbonCheckBox();
            this.checkBox1 = this.Factory.CreateRibbonCheckBox();
            this.HelloWorldBtn = this.Factory.CreateRibbonButton();
            this.sent.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // sent
            // 
            this.sent.Groups.Add(this.group1);
            this.sent.Label = "AnalyseCells";
            this.sent.Name = "sent";
            // 
            // group1
            // 
            this.group1.Items.Add(this.Classify);
            this.group1.Items.Add(this.ClassifySelection);
            this.group1.Items.Add(this.Send);
            this.group1.Items.Add(this.button2);
            this.group1.Items.Add(this.checkBox1);
            this.group1.Items.Add(this.checkBox4);
            this.group1.Items.Add(this.checkBox3);
            this.group1.Items.Add(this.checkBox5);
            this.group1.Items.Add(this.checkBox2);
            this.group1.Items.Add(this.button1);
            this.group1.Items.Add(this.HelloWorldBtn);
            this.group1.Label = "First";
            this.group1.Name = "group1";
            // 
            // ClassifySelection
            // 
            this.ClassifySelection.Label = "Classify selection";
            this.ClassifySelection.Name = "ClassifySelection";
            this.ClassifySelection.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // button1
            // 
            this.button1.Label = "";
            this.button1.Name = "button1";
            // 
            // Classify
            // 
            this.Classify.Label = "Classify";
            this.Classify.Name = "Classify";
            this.Classify.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.sendToAPI_Click);
            // 
            // Send
            // 
            this.Send.Label = "Change URLs";
            this.Send.Name = "Send";
            this.Send.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Send_Click);
            // 
            // button2
            // 
            this.button2.Label = "Reset";
            this.button2.Name = "button2";
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click_1);
            // 
            // checkBox2
            // 
            this.checkBox2.Label = "show/hide Header labels";
            this.checkBox2.Name = "checkBox2";
            this.checkBox2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBox2_Click);
            // 
            // checkBox3
            // 
            this.checkBox3.Label = "show/hide Derived labels";
            this.checkBox3.Name = "checkBox3";
            this.checkBox3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBox3_Click);
            // 
            // checkBox4
            // 
            this.checkBox4.Label = "show/hide Metadata labels";
            this.checkBox4.Name = "checkBox4";
            this.checkBox4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBox4_Click);
            // 
            // checkBox5
            // 
            this.checkBox5.Label = "show/hide Attributes labels";
            this.checkBox5.Name = "checkBox5";
            this.checkBox5.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBox5_Click);
            // 
            // checkBox1
            // 
            this.checkBox1.Label = "show/hide Data labels";
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBox1_Click);
            // 
            // HelloWorldBtn
            // 
            this.HelloWorldBtn.Label = "";
            this.HelloWorldBtn.Name = "HelloWorldBtn";
            // 
            // MyRibbon
            // 
            this.Name = "MyRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.sent);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.sent.ResumeLayout(false);
            this.sent.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab sent;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton HelloWorldBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Classify;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Send;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ClassifySelection;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox1;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox2;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox3;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox4;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
    }

    partial class ThisRibbonCollection
    {
        internal MyRibbon Ribbon1
        {
            get { return this.GetRibbon<MyRibbon>(); }
        }
    }
}
