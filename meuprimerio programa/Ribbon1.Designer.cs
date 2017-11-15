namespace meuprimerio_programa
    {
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
        {
        /// <summary>
        /// Variável de designer necessária.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
            {
            InitializeComponent();
            }

        /// <summary> 
        /// Limpar os recursos que estão sendo usados.
        /// </summary>
        /// <param name="disposing">true se for necessário descartar os recursos gerenciados; caso contrário, false.</param>
        protected override void Dispose(bool disposing)
            {
            if ( disposing && ( components != null ) )
                {
                components.Dispose();
                }
            base.Dispose(disposing);
            }

        #region Código gerado pelo Designer de Componentes

        /// <summary>
        /// Método necessário para suporte ao Designer - não modifique 
        /// o conteúdo deste método com o editor de código.
        /// </summary>
        private void InitializeComponent()
            {
            this.EsctrucMex = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.OlaMundoBtn = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.EsctrucMex.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // EsctrucMex
            // 
            this.EsctrucMex.Groups.Add(this.group1);
            this.EsctrucMex.Groups.Add(this.group2);
            this.EsctrucMex.Label = "Estruc Mex";
            this.EsctrucMex.Name = "EsctrucMex";
            // 
            // group1
            // 
            this.group1.Items.Add(this.OlaMundoBtn);
            this.group1.Items.Add(this.button1);
            this.group1.Label = "Fase 1 - Nós";
            this.group1.Name = "group1";
            // 
            // group2
            // 
            this.group2.Items.Add(this.button2);
            this.group2.Items.Add(this.button3);
            this.group2.Label = "Fase 2 - Conectividades ";
            this.group2.Name = "group2";
            // 
            // OlaMundoBtn
            // 
            this.OlaMundoBtn.Image = global::meuprimerio_programa.Properties.Resources.NósIn;
            this.OlaMundoBtn.Label = "Informar os nós";
            this.OlaMundoBtn.Name = "OlaMundoBtn";
            this.OlaMundoBtn.ShowImage = true;
            this.OlaMundoBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OlaMundoBtn_Click);
            // 
            // button1
            // 
            this.button1.Image = global::meuprimerio_programa.Properties.Resources.NósOut;
            this.button1.Label = "Carregar Nós";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Image = global::meuprimerio_programa.Properties.Resources.BarrasIn;
            this.button2.Label = "Informar Conectividades";
            this.button2.Name = "button2";
            this.button2.ShowImage = true;
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Image = global::meuprimerio_programa.Properties.Resources.BarrasOut;
            this.button3.Label = "Carregar Conectividades";
            this.button3.Name = "button3";
            this.button3.ShowImage = true;
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.EsctrucMex);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.EsctrucMex.ResumeLayout(false);
            this.EsctrucMex.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

            }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab EsctrucMex;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton OlaMundoBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        }

    partial class ThisRibbonCollection
        {
        internal Ribbon1 Ribbon1
            {
            get { return this.GetRibbon<Ribbon1>(); }
            }
        }
    }
