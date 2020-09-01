namespace ribbon_xnk_sunzex
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            this.tab_xnk = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.button5 = this.Factory.CreateRibbonButton();
            this.editBox_hoadonmau = this.Factory.CreateRibbonEditBox();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.button7 = this.Factory.CreateRibbonButton();
            this.editBox_thumucxuat = this.Factory.CreateRibbonEditBox();
            this.group_xnk = this.Factory.CreateRibbonGroup();
            this.button_invoice = this.Factory.CreateRibbonButton();
            this.button_shipping = this.Factory.CreateRibbonButton();
            this.editBox3 = this.Factory.CreateRibbonEditBox();
            this.checkBox_openAfter = this.Factory.CreateRibbonCheckBox();
            this.checkBox_xuatRieng = this.Factory.CreateRibbonCheckBox();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.editBox4 = this.Factory.CreateRibbonEditBox();
            this.editBox5 = this.Factory.CreateRibbonEditBox();
            this.checkBox1 = this.Factory.CreateRibbonCheckBox();
            this.button1 = this.Factory.CreateRibbonButton();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.tab_xnk.SuspendLayout();
            this.group1.SuspendLayout();
            this.group_xnk.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab_xnk
            // 
            this.tab_xnk.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab_xnk.Groups.Add(this.group1);
            this.tab_xnk.Groups.Add(this.group_xnk);
            this.tab_xnk.Groups.Add(this.group2);
            this.tab_xnk.Label = "XNK Sunzex";
            this.tab_xnk.Name = "tab_xnk";
            // 
            // group1
            // 
            this.group1.Items.Add(this.button5);
            this.group1.Items.Add(this.editBox_hoadonmau);
            this.group1.Items.Add(this.separator2);
            this.group1.Items.Add(this.button7);
            this.group1.Items.Add(this.editBox_thumucxuat);
            this.group1.Name = "group1";
            // 
            // button5
            // 
            this.button5.Image = ((System.Drawing.Image)(resources.GetObject("button5.Image")));
            this.button5.Label = "Hóa đơn mẫu";
            this.button5.Name = "button5";
            this.button5.ShowImage = true;
            this.button5.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button5_Click);
            // 
            // editBox_hoadonmau
            // 
            this.editBox_hoadonmau.Enabled = false;
            this.editBox_hoadonmau.Label = " ";
            this.editBox_hoadonmau.Name = "editBox_hoadonmau";
            this.editBox_hoadonmau.ShowLabel = false;
            this.editBox_hoadonmau.Text = null;
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // button7
            // 
            this.button7.Image = ((System.Drawing.Image)(resources.GetObject("button7.Image")));
            this.button7.Label = "Thư mục xuất";
            this.button7.Name = "button7";
            this.button7.ShowImage = true;
            this.button7.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button7_Click);
            // 
            // editBox_thumucxuat
            // 
            this.editBox_thumucxuat.Label = " ";
            this.editBox_thumucxuat.Name = "editBox_thumucxuat";
            this.editBox_thumucxuat.ShowLabel = false;
            this.editBox_thumucxuat.Text = null;
            // 
            // group_xnk
            // 
            this.group_xnk.Items.Add(this.button_invoice);
            this.group_xnk.Items.Add(this.button_shipping);
            this.group_xnk.Items.Add(this.editBox3);
            this.group_xnk.Items.Add(this.checkBox_openAfter);
            this.group_xnk.Items.Add(this.checkBox_xuatRieng);
            this.group_xnk.Label = "Xuất";
            this.group_xnk.Name = "group_xnk";
            // 
            // button_invoice
            // 
            this.button_invoice.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_invoice.Image = ((System.Drawing.Image)(resources.GetObject("button_invoice.Image")));
            this.button_invoice.Label = "Invoice";
            this.button_invoice.Name = "button_invoice";
            this.button_invoice.ShowImage = true;
            this.button_invoice.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_invoice_Click);
            // 
            // button_shipping
            // 
            this.button_shipping.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_shipping.Image = ((System.Drawing.Image)(resources.GetObject("button_shipping.Image")));
            this.button_shipping.Label = "Shipping";
            this.button_shipping.Name = "button_shipping";
            this.button_shipping.ShowImage = true;
            this.button_shipping.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_shipping_Click);
            // 
            // editBox3
            // 
            this.editBox3.Enabled = false;
            this.editBox3.Label = "Packing list";
            this.editBox3.Name = "editBox3";
            this.editBox3.Text = null;
            // 
            // checkBox_openAfter
            // 
            this.checkBox_openAfter.Label = "Mở file sau khi tạo";
            this.checkBox_openAfter.Name = "checkBox_openAfter";
            this.checkBox_openAfter.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CheckBox3_Click);
            // 
            // checkBox_xuatRieng
            // 
            this.checkBox_xuatRieng.Label = "Xuất file riêng";
            this.checkBox_xuatRieng.Name = "checkBox_xuatRieng";
            this.checkBox_xuatRieng.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBox_xuatRieng_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.editBox4);
            this.group2.Items.Add(this.editBox5);
            this.group2.Items.Add(this.checkBox1);
            this.group2.Items.Add(this.button1);
            this.group2.Label = "Xuất từ nhiều TK";
            this.group2.Name = "group2";
            // 
            // editBox4
            // 
            this.editBox4.Enabled = false;
            this.editBox4.Label = "List TK";
            this.editBox4.Name = "editBox4";
            this.editBox4.Text = null;
            // 
            // editBox5
            // 
            this.editBox5.Enabled = false;
            this.editBox5.Label = "Packinglist";
            this.editBox5.Name = "editBox5";
            this.editBox5.Text = null;
            // 
            // checkBox1
            // 
            this.checkBox1.Enabled = false;
            this.checkBox1.Label = "Xuất file riêng";
            this.checkBox1.Name = "checkBox1";
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Enabled = false;
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.Label = "Xuất toàn bộ";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab_xnk);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab_xnk.ResumeLayout(false);
            this.tab_xnk.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group_xnk.ResumeLayout(false);
            this.group_xnk.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab_xnk;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_xnk;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_invoice;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_shipping;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox3;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox_xuatRieng;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox4;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox5;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button5;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox_hoadonmau;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button7;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox_thumucxuat;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox_openAfter;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
