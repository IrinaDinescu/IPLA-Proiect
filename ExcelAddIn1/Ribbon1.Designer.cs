namespace ExcelAddIn1
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.btn_ShowPortofolio = this.Factory.CreateRibbonButton();
            this.btnShowStockEvolution = this.Factory.CreateRibbonButton();
            this.btn_UpdatePortofolio = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.btn_AddBuys = this.Factory.CreateRibbonButton();
            this.btn_SaveBuysDatabase = this.Factory.CreateRibbonButton();
            this.btn_ShowAllBuys = this.Factory.CreateRibbonButton();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btn_AddSales = this.Factory.CreateRibbonButton();
            this.btn_SaveSalesDatabase = this.Factory.CreateRibbonButton();
            this.btn_ShowAllSales = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "Portofolio Manager";
            this.tab1.Name = "tab1";
            // 
            // group3
            // 
            this.group3.Items.Add(this.btn_ShowPortofolio);
            this.group3.Items.Add(this.btnShowStockEvolution);
            this.group3.Items.Add(this.btn_UpdatePortofolio);
            this.group3.Label = "Portofolio";
            this.group3.Name = "group3";
            // 
            // btn_ShowPortofolio
            // 
            this.btn_ShowPortofolio.Label = "Show Portofolio";
            this.btn_ShowPortofolio.Name = "btn_ShowPortofolio";
            this.btn_ShowPortofolio.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_ShowPortofolio_Click);
            // 
            // btnShowStockEvolution
            // 
            this.btnShowStockEvolution.Label = "Show Stock Evolution";
            this.btnShowStockEvolution.Name = "btnShowStockEvolution";
            this.btnShowStockEvolution.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnShowStockEvolution_Click);
            // 
            // btn_UpdatePortofolio
            // 
            this.btn_UpdatePortofolio.Label = "Update Portofolio";
            this.btn_UpdatePortofolio.Name = "btn_UpdatePortofolio";
            this.btn_UpdatePortofolio.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_UpdatePortofolio_Click);
            // 
            // group4
            // 
            this.group4.Items.Add(this.btn_AddBuys);
            this.group4.Items.Add(this.btn_SaveBuysDatabase);
            this.group4.Items.Add(this.btn_ShowAllBuys);
            this.group4.Label = "Buys";
            this.group4.Name = "group4";
            // 
            // btn_AddBuys
            // 
            this.btn_AddBuys.Label = "Add Buys";
            this.btn_AddBuys.Name = "btn_AddBuys";
            this.btn_AddBuys.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_AddBuys_Click);
            // 
            // btn_SaveBuysDatabase
            // 
            this.btn_SaveBuysDatabase.Label = "Save ";
            this.btn_SaveBuysDatabase.Name = "btn_SaveBuysDatabase";
            this.btn_SaveBuysDatabase.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_SaveBuysDatabase_Click);
            // 
            // btn_ShowAllBuys
            // 
            this.btn_ShowAllBuys.Label = "Show All Buys";
            this.btn_ShowAllBuys.Name = "btn_ShowAllBuys";
            this.btn_ShowAllBuys.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_ShowAllBuys_Click);
            // 
            // group1
            // 
            this.group1.Items.Add(this.btn_AddSales);
            this.group1.Items.Add(this.btn_SaveSalesDatabase);
            this.group1.Items.Add(this.btn_ShowAllSales);
            this.group1.Label = "Sales";
            this.group1.Name = "group1";
            // 
            // btn_AddSales
            // 
            this.btn_AddSales.Label = "Add Sales";
            this.btn_AddSales.Name = "btn_AddSales";
            this.btn_AddSales.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_AddSales_Click);
            // 
            // btn_SaveSalesDatabase
            // 
            this.btn_SaveSalesDatabase.Label = "Save";
            this.btn_SaveSalesDatabase.Name = "btn_SaveSalesDatabase";
            this.btn_SaveSalesDatabase.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_SaveSalesDatabase_Click);
            // 
            // btn_ShowAllSales
            // 
            this.btn_ShowAllSales.Label = "Show All Sales";
            this.btn_ShowAllSales.Name = "btn_ShowAllSales";
            this.btn_ShowAllSales.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_ShowAllSales_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_ShowPortofolio;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_AddBuys;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_SaveBuysDatabase;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_ShowAllBuys;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnShowStockEvolution;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_AddSales;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_SaveSalesDatabase;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_ShowAllSales;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_UpdatePortofolio;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
