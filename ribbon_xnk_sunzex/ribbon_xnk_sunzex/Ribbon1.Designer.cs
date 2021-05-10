namespace ribbon_xnk_sunzex
{
    partial class xnk_ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public xnk_ribbon()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(xnk_ribbon));
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            this.tab_xnk_tisu = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.button5 = this.Factory.CreateRibbonButton();
            this.editBox_hoadonmau = this.Factory.CreateRibbonEditBox();
            this.combobox1 = this.Factory.CreateRibbonComboBox();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.button7 = this.Factory.CreateRibbonButton();
            this.editBox_thumucxuat = this.Factory.CreateRibbonEditBox();
            this.group_xnk = this.Factory.CreateRibbonGroup();
            this.button_invoice = this.Factory.CreateRibbonButton();
            this.button_shipping = this.Factory.CreateRibbonButton();
            this.checkBox_openAfter = this.Factory.CreateRibbonCheckBox();
            this.checkBox_xuatRieng = this.Factory.CreateRibbonCheckBox();
            this.checkBox_PKL = this.Factory.CreateRibbonCheckBox();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.editBox4 = this.Factory.CreateRibbonEditBox();
            this.editBox5 = this.Factory.CreateRibbonEditBox();
            this.checkBox1 = this.Factory.CreateRibbonCheckBox();
            this.button1 = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.button3_An = this.Factory.CreateRibbonButton();
            this.button4_Hien = this.Factory.CreateRibbonButton();
            this.btn_Calc_NXK = this.Factory.CreateRibbonButton();
            this.checkBoxSTT = this.Factory.CreateRibbonCheckBox();
            this.editBoxFromSTT = this.Factory.CreateRibbonEditBox();
            this.editBoxToSTT = this.Factory.CreateRibbonEditBox();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.tinh_PKL = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.checkBox_PKL_ghi = this.Factory.CreateRibbonCheckBox();
            this.group5 = this.Factory.CreateRibbonGroup();
            this.button3 = this.Factory.CreateRibbonButton();
            this.tab_xnk_tisu.SuspendLayout();
            this.group1.SuspendLayout();
            this.group_xnk.SuspendLayout();
            this.group2.SuspendLayout();
            this.group4.SuspendLayout();
            this.group3.SuspendLayout();
            this.group5.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab_xnk_tisu
            // 
            this.tab_xnk_tisu.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab_xnk_tisu.Groups.Add(this.group1);
            this.tab_xnk_tisu.Groups.Add(this.group_xnk);
            this.tab_xnk_tisu.Groups.Add(this.group2);
            this.tab_xnk_tisu.Groups.Add(this.group4);
            this.tab_xnk_tisu.Groups.Add(this.group3);
            this.tab_xnk_tisu.Groups.Add(this.group5);
            this.tab_xnk_tisu.Label = "Xnk Tisu";
            this.tab_xnk_tisu.Name = "tab_xnk_tisu";
            // 
            // group1
            // 
            this.group1.Items.Add(this.button5);
            this.group1.Items.Add(this.editBox_hoadonmau);
            this.group1.Items.Add(this.combobox1);
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
            this.editBox_hoadonmau.Label = " ";
            this.editBox_hoadonmau.Name = "editBox_hoadonmau";
            this.editBox_hoadonmau.ShowLabel = false;
            this.editBox_hoadonmau.Text = null;
            // 
            // combobox1
            // 
            ribbonDropDownItemImpl1.Label = "Tisu";
            ribbonDropDownItemImpl2.Label = "Sunzex";
            this.combobox1.Items.Add(ribbonDropDownItemImpl1);
            this.combobox1.Items.Add(ribbonDropDownItemImpl2);
            this.combobox1.Label = "cty";
            this.combobox1.Name = "combobox1";
            this.combobox1.ShowLabel = false;
            this.combobox1.Text = null;
            this.combobox1.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.combobox1_TextChanged);
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
            this.group_xnk.Items.Add(this.checkBox_openAfter);
            this.group_xnk.Items.Add(this.checkBox_xuatRieng);
            this.group_xnk.Items.Add(this.checkBox_PKL);
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
            // checkBox_openAfter
            // 
            this.checkBox_openAfter.Label = "Mở file được tạo sau khi xuất";
            this.checkBox_openAfter.Name = "checkBox_openAfter";
            this.checkBox_openAfter.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CheckBox3_Click);
            // 
            // checkBox_xuatRieng
            // 
            this.checkBox_xuatRieng.Label = "Xuất file riêng không theo tháng";
            this.checkBox_xuatRieng.Name = "checkBox_xuatRieng";
            this.checkBox_xuatRieng.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBox_xuatRieng_Click);
            // 
            // checkBox_PKL
            // 
            this.checkBox_PKL.Label = "Tự động điền PO từ PKL";
            this.checkBox_PKL.Name = "checkBox_PKL";
            this.checkBox_PKL.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBox_PKL_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.editBox4);
            this.group2.Items.Add(this.editBox5);
            this.group2.Items.Add(this.checkBox1);
            this.group2.Items.Add(this.button1);
            this.group2.Label = "Xuất nhiều TK";
            this.group2.Name = "group2";
            this.group2.Visible = false;
            // 
            // editBox4
            // 
            this.editBox4.Enabled = false;
            this.editBox4.Label = "Mục chứa các tờ khai";
            this.editBox4.Name = "editBox4";
            this.editBox4.Text = null;
            // 
            // editBox5
            // 
            this.editBox5.Enabled = false;
            this.editBox5.Label = "Thư mục Packinglist";
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
            // group4
            // 
            this.group4.Items.Add(this.button3_An);
            this.group4.Items.Add(this.button4_Hien);
            this.group4.Items.Add(this.btn_Calc_NXK);
            this.group4.Items.Add(this.checkBoxSTT);
            this.group4.Items.Add(this.editBoxFromSTT);
            this.group4.Items.Add(this.editBoxToSTT);
            this.group4.Label = "Hỗ trợ tính nhập xuất kho";
            this.group4.Name = "group4";
            // 
            // button3_An
            // 
            this.button3_An.Label = "ẨN CỘT";
            this.button3_An.Name = "button3_An";
            this.button3_An.ScreenTip = "Ẩn các cột không cần thiết trong bảng chi tiết hàng hóa xuất khẩu";
            this.button3_An.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_An);
            // 
            // button4_Hien
            // 
            this.button4_Hien.Label = "HIỆN CỘT";
            this.button4_Hien.Name = "button4_Hien";
            this.button4_Hien.ScreenTip = "Hiện lại các cột đã ẩn trong bảng chi tiết hàng hóa";
            this.button4_Hien.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_Hien);
            // 
            // btn_Calc_NXK
            // 
            this.btn_Calc_NXK.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_Calc_NXK.Image = ((System.Drawing.Image)(resources.GetObject("btn_Calc_NXK.Image")));
            this.btn_Calc_NXK.Label = "Tổng xuất NPL theo từng mã";
            this.btn_Calc_NXK.Name = "btn_Calc_NXK";
            this.btn_Calc_NXK.ShowImage = true;
            this.btn_Calc_NXK.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_Calc_NXK_Click);
            // 
            // checkBoxSTT
            // 
            this.checkBoxSTT.Enabled = false;
            this.checkBoxSTT.Label = "Chọn theo stt";
            this.checkBoxSTT.Name = "checkBoxSTT";
            // 
            // editBoxFromSTT
            // 
            this.editBoxFromSTT.Label = " Từ stt";
            this.editBoxFromSTT.Name = "editBoxFromSTT";
            this.editBoxFromSTT.Text = null;
            // 
            // editBoxToSTT
            // 
            this.editBoxToSTT.Label = " Đến stt";
            this.editBoxToSTT.Name = "editBoxToSTT";
            this.editBoxToSTT.Text = null;
            // 
            // group3
            // 
            this.group3.Items.Add(this.tinh_PKL);
            this.group3.Items.Add(this.button2);
            this.group3.Items.Add(this.checkBox_PKL_ghi);
            this.group3.Label = "Packing list";
            this.group3.Name = "group3";
            // 
            // tinh_PKL
            // 
            this.tinh_PKL.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.tinh_PKL.Image = ((System.Drawing.Image)(resources.GetObject("tinh_PKL.Image")));
            this.tinh_PKL.Label = "Tổng PCS theo vùng được chọn";
            this.tinh_PKL.Name = "tinh_PKL";
            this.tinh_PKL.ShowImage = true;
            this.tinh_PKL.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_tinh_PKL);
            // 
            // button2
            // 
            this.button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
            this.button2.Label = "Tổng toàn bộ";
            this.button2.Name = "button2";
            this.button2.ShowImage = true;
            // 
            // checkBox_PKL_ghi
            // 
            this.checkBox_PKL_ghi.Description = "Ghi kết quả vào file excel thay vì hiện ra thông báo";
            this.checkBox_PKL_ghi.Label = "Ghi vào file";
            this.checkBox_PKL_ghi.Name = "checkBox_PKL_ghi";
            this.checkBox_PKL_ghi.ScreenTip = "Ghi kết quả vào file excel thay vì hiện ra thông báo";
            this.checkBox_PKL_ghi.SuperTip = "Ghi kết quả vào file excel thay vì hiện ra thông báo";
            this.checkBox_PKL_ghi.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBox_PKL_ghi_Click);
            // 
            // group5
            // 
            this.group5.Items.Add(this.button3);
            this.group5.Name = "group5";
            // 
            // button3
            // 
            this.button3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button3.Image = ((System.Drawing.Image)(resources.GetObject("button3.Image")));
            this.button3.Label = "Hướng dẫn và phiên bản";
            this.button3.Name = "button3";
            this.button3.ShowImage = true;
            // 
            // xnk_ribbon
            // 
            this.Name = "xnk_ribbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab_xnk_tisu);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab_xnk_tisu.ResumeLayout(false);
            this.tab_xnk_tisu.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group_xnk.ResumeLayout(false);
            this.group_xnk.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group5.ResumeLayout(false);
            this.group5.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab_xnk_tisu;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_xnk;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_invoice;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_shipping;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox_xuatRieng;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox4;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox5;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button5;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox_hoadonmau;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button7;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox_thumucxuat;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox_openAfter;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox_PKL;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox combobox1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_Calc_NXK;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBoxSTT;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBoxFromSTT;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBoxToSTT;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3_An;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4_Hien;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton tinh_PKL;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox_PKL_ghi;
    }

    partial class ThisRibbonCollection
    {
        internal xnk_ribbon Ribbon1
        {
            get { return this.GetRibbon<xnk_ribbon>(); }
        }
    }
}
