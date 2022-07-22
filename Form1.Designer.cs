namespace WindowsFormsApplication1
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        
        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.sistemaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.processarToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem4 = new System.Windows.Forms.ToolStripMenuItem();
            this.descompactarToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.processarToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.processarBACENToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.processarToolStripMenuItem2 = new System.Windows.Forms.ToolStripMenuItem();
            this.processarCB7_MenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem5 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem9 = new System.Windows.Forms.ToolStripMenuItem();
            this.ProcessarCartaDOCLA = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem12 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem10 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem11 = new System.Windows.Forms.ToolStripMenuItem();
            this.fase01ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.fase02ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem13 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem14 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem8 = new System.Windows.Forms.ToolStripMenuItem();
            this.Consignado = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem15 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem7 = new System.Windows.Forms.ToolStripMenuItem();
            this.boletoA4ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem2 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem6 = new System.Windows.Forms.ToolStripMenuItem();
            this.separaErroToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripMenuItem3 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.sairToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.butProcessarCarne = new System.Windows.Forms.Button();
            this.lblValor = new System.Windows.Forms.Label();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.bgWorkerIndeterminada = new System.ComponentModel.BackgroundWorker();
            this.lblMsg = new System.Windows.Forms.Label();
            this.butProcessaCartas = new System.Windows.Forms.Button();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.sistemaToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Padding = new System.Windows.Forms.Padding(4, 2, 0, 2);
            this.menuStrip1.Size = new System.Drawing.Size(434, 24);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // sistemaToolStripMenuItem
            // 
            this.sistemaToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.processarToolStripMenuItem,
            this.toolStripMenuItem4,
            this.toolStripMenuItem5,
            this.toolStripMenuItem9,
            this.ProcessarCartaDOCLA,
            this.toolStripMenuItem12,
            this.toolStripMenuItem10,
            this.toolStripMenuItem11,
            this.toolStripMenuItem13,
            this.toolStripMenuItem14,
            this.toolStripMenuItem8,
            this.Consignado,
            this.toolStripMenuItem15,
            this.toolStripMenuItem7,
            this.toolStripMenuItem1,
            this.toolStripMenuItem2,
            this.toolStripMenuItem6,
            this.toolStripSeparator1,
            this.toolStripMenuItem3,
            this.toolStripSeparator2,
            this.sairToolStripMenuItem});
            this.sistemaToolStripMenuItem.Name = "sistemaToolStripMenuItem";
            this.sistemaToolStripMenuItem.Size = new System.Drawing.Size(60, 20);
            this.sistemaToolStripMenuItem.Text = "Sistema";
            // 
            // processarToolStripMenuItem
            // 
            this.processarToolStripMenuItem.Name = "processarToolStripMenuItem";
            this.processarToolStripMenuItem.Size = new System.Drawing.Size(229, 22);
            this.processarToolStripMenuItem.Text = "Processar Carnes";
            this.processarToolStripMenuItem.Click += new System.EventHandler(this.processarToolStripMenuItem_Click);
            // 
            // toolStripMenuItem4
            // 
            this.toolStripMenuItem4.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.descompactarToolStripMenuItem,
            this.processarToolStripMenuItem1,
            this.processarBACENToolStripMenuItem,
            this.processarToolStripMenuItem2,
            this.processarCB7_MenuItem});
            this.toolStripMenuItem4.Name = "toolStripMenuItem4";
            this.toolStripMenuItem4.Size = new System.Drawing.Size(229, 22);
            this.toolStripMenuItem4.Text = "Processar Cartas CCB/BACEN";
            // 
            // descompactarToolStripMenuItem
            // 
            this.descompactarToolStripMenuItem.Name = "descompactarToolStripMenuItem";
            this.descompactarToolStripMenuItem.Size = new System.Drawing.Size(165, 22);
            this.descompactarToolStripMenuItem.Text = "Descompactar";
            this.descompactarToolStripMenuItem.Click += new System.EventHandler(this.descompactarToolStripMenuItem_Click);
            // 
            // processarToolStripMenuItem1
            // 
            this.processarToolStripMenuItem1.Name = "processarToolStripMenuItem1";
            this.processarToolStripMenuItem1.Size = new System.Drawing.Size(165, 22);
            this.processarToolStripMenuItem1.Text = "Processar CCB";
            this.processarToolStripMenuItem1.Click += new System.EventHandler(this.toolStripMenuItem4_Click);
            // 
            // processarBACENToolStripMenuItem
            // 
            this.processarBACENToolStripMenuItem.Name = "processarBACENToolStripMenuItem";
            this.processarBACENToolStripMenuItem.Size = new System.Drawing.Size(165, 22);
            this.processarBACENToolStripMenuItem.Text = "Processar BACEN";
            this.processarBACENToolStripMenuItem.Click += new System.EventHandler(this.processarBACENToolStripMenuItem_Click);
            // 
            // processarToolStripMenuItem2
            // 
            this.processarToolStripMenuItem2.Name = "processarToolStripMenuItem2";
            this.processarToolStripMenuItem2.Size = new System.Drawing.Size(165, 22);
            this.processarToolStripMenuItem2.Text = "Processar CCT";
            this.processarToolStripMenuItem2.Click += new System.EventHandler(this.processarToolStripMenuItem2_Click);
            // 
            // processarCB7_MenuItem
            // 
            this.processarCB7_MenuItem.Name = "processarCB7_MenuItem";
            this.processarCB7_MenuItem.Size = new System.Drawing.Size(165, 22);
            this.processarCB7_MenuItem.Text = "Processar CB7";
            this.processarCB7_MenuItem.Click += new System.EventHandler(this.processarCB7_MenuItem_Click);
            // 
            // toolStripMenuItem5
            // 
            this.toolStripMenuItem5.Name = "toolStripMenuItem5";
            this.toolStripMenuItem5.Size = new System.Drawing.Size(229, 22);
            this.toolStripMenuItem5.Text = "Processar CQC";
            this.toolStripMenuItem5.Click += new System.EventHandler(this.toolStripMenuItem5_Click);
            // 
            // toolStripMenuItem9
            // 
            this.toolStripMenuItem9.Name = "toolStripMenuItem9";
            this.toolStripMenuItem9.Size = new System.Drawing.Size(229, 22);
            this.toolStripMenuItem9.Text = "Processar TC";
            this.toolStripMenuItem9.Click += new System.EventHandler(this.toolStripMenuItem9_Click);
            // 
            // ProcessarCartaDOCLA
            // 
            this.ProcessarCartaDOCLA.Name = "ProcessarCartaDOCLA";
            this.ProcessarCartaDOCLA.Size = new System.Drawing.Size(229, 22);
            this.ProcessarCartaDOCLA.Text = "Processar DOCLA";
            this.ProcessarCartaDOCLA.Click += new System.EventHandler(this.ProcessarCartaDOCLA_Click);
            // 
            // toolStripMenuItem12
            // 
            this.toolStripMenuItem12.Name = "toolStripMenuItem12";
            this.toolStripMenuItem12.Size = new System.Drawing.Size(229, 22);
            this.toolStripMenuItem12.Text = "Cartas DUT";
            this.toolStripMenuItem12.Click += new System.EventHandler(this.toolStripMenuItem12_Click_1);
            // 
            // toolStripMenuItem10
            // 
            this.toolStripMenuItem10.Name = "toolStripMenuItem10";
            this.toolStripMenuItem10.Size = new System.Drawing.Size(229, 22);
            this.toolStripMenuItem10.Text = "Carta TCCONSIG";
            this.toolStripMenuItem10.Click += new System.EventHandler(this.toolStripMenuItem10_Click);
            // 
            // toolStripMenuItem11
            // 
            this.toolStripMenuItem11.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fase01ToolStripMenuItem,
            this.fase02ToolStripMenuItem});
            this.toolStripMenuItem11.Name = "toolStripMenuItem11";
            this.toolStripMenuItem11.Size = new System.Drawing.Size(229, 22);
            this.toolStripMenuItem11.Text = "Carta CVV";
            this.toolStripMenuItem11.Click += new System.EventHandler(this.toolStripMenuItem11_Click);
            // 
            // fase01ToolStripMenuItem
            // 
            this.fase01ToolStripMenuItem.Name = "fase01ToolStripMenuItem";
            this.fase01ToolStripMenuItem.Size = new System.Drawing.Size(112, 22);
            this.fase01ToolStripMenuItem.Text = "Fase 01";
            this.fase01ToolStripMenuItem.Click += new System.EventHandler(this.fase01ToolStripMenuItem_Click);
            // 
            // fase02ToolStripMenuItem
            // 
            this.fase02ToolStripMenuItem.Name = "fase02ToolStripMenuItem";
            this.fase02ToolStripMenuItem.Size = new System.Drawing.Size(112, 22);
            this.fase02ToolStripMenuItem.Text = "Fase 02";
            this.fase02ToolStripMenuItem.Click += new System.EventHandler(this.fase02ToolStripMenuItem_Click);
            // 
            // toolStripMenuItem13
            // 
            this.toolStripMenuItem13.Name = "toolStripMenuItem13";
            this.toolStripMenuItem13.Size = new System.Drawing.Size(229, 22);
            this.toolStripMenuItem13.Text = "Carta BNDU";
            this.toolStripMenuItem13.Click += new System.EventHandler(this.toolStripMenuItem13_Click);
            // 
            // toolStripMenuItem14
            // 
            this.toolStripMenuItem14.Name = "toolStripMenuItem14";
            this.toolStripMenuItem14.Size = new System.Drawing.Size(229, 22);
            this.toolStripMenuItem14.Text = "Carta DVVDUPLICIDADE";
            this.toolStripMenuItem14.Click += new System.EventHandler(this.toolStripMenuItem14_Click);
            // 
            // toolStripMenuItem8
            // 
            this.toolStripMenuItem8.Name = "toolStripMenuItem8";
            this.toolStripMenuItem8.Size = new System.Drawing.Size(229, 22);
            this.toolStripMenuItem8.Text = "Processar Detran";
            this.toolStripMenuItem8.Click += new System.EventHandler(this.toolStripMenuItem8_Click);
            // 
            // Consignado
            // 
            this.Consignado.Name = "Consignado";
            this.Consignado.Size = new System.Drawing.Size(229, 22);
            this.Consignado.Text = "Cartas Consignado";
            this.Consignado.Click += new System.EventHandler(this.toolStripMenuItem15_Click);
            // 
            // toolStripMenuItem15
            // 
            this.toolStripMenuItem15.Name = "toolStripMenuItem15";
            this.toolStripMenuItem15.Size = new System.Drawing.Size(229, 22);
            this.toolStripMenuItem15.Text = "Cartas Veiculos";
            this.toolStripMenuItem15.Click += new System.EventHandler(this.toolStripMenuItem15_Click_1);
            // 
            // toolStripMenuItem7
            // 
            this.toolStripMenuItem7.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.boletoA4ToolStripMenuItem});
            this.toolStripMenuItem7.Name = "toolStripMenuItem7";
            this.toolStripMenuItem7.Size = new System.Drawing.Size(229, 22);
            this.toolStripMenuItem7.Text = "Santa Casa";
            this.toolStripMenuItem7.Click += new System.EventHandler(this.toolStripMenuItem7_Click);
            // 
            // boletoA4ToolStripMenuItem
            // 
            this.boletoA4ToolStripMenuItem.Name = "boletoA4ToolStripMenuItem";
            this.boletoA4ToolStripMenuItem.Size = new System.Drawing.Size(125, 22);
            this.boletoA4ToolStripMenuItem.Text = "Boleto A4";
            this.boletoA4ToolStripMenuItem.Click += new System.EventHandler(this.boletoA4ToolStripMenuItem_Click);
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(229, 22);
            this.toolStripMenuItem1.Text = "Gerar Midias";
            this.toolStripMenuItem1.Click += new System.EventHandler(this.toolStripMenuItem1_Click);
            // 
            // toolStripMenuItem2
            // 
            this.toolStripMenuItem2.Name = "toolStripMenuItem2";
            this.toolStripMenuItem2.Size = new System.Drawing.Size(229, 22);
            this.toolStripMenuItem2.Text = "Retirar Contrato";
            this.toolStripMenuItem2.Click += new System.EventHandler(this.toolStripMenuItem2_Click);
            // 
            // toolStripMenuItem6
            // 
            this.toolStripMenuItem6.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.separaErroToolStripMenuItem});
            this.toolStripMenuItem6.Name = "toolStripMenuItem6";
            this.toolStripMenuItem6.Size = new System.Drawing.Size(229, 22);
            this.toolStripMenuItem6.Text = "Remessa";
            // 
            // separaErroToolStripMenuItem
            // 
            this.separaErroToolStripMenuItem.Name = "separaErroToolStripMenuItem";
            this.separaErroToolStripMenuItem.Size = new System.Drawing.Size(133, 22);
            this.separaErroToolStripMenuItem.Text = "Separa Erro";
            this.separaErroToolStripMenuItem.Click += new System.EventHandler(this.separaErroToolStripMenuItem_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(226, 6);
            // 
            // toolStripMenuItem3
            // 
            this.toolStripMenuItem3.Name = "toolStripMenuItem3";
            this.toolStripMenuItem3.Size = new System.Drawing.Size(229, 22);
            this.toolStripMenuItem3.Text = "Gerar Separador";
            this.toolStripMenuItem3.Click += new System.EventHandler(this.toolStripMenuItem3_Click);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(226, 6);
            // 
            // sairToolStripMenuItem
            // 
            this.sairToolStripMenuItem.Name = "sairToolStripMenuItem";
            this.sairToolStripMenuItem.Size = new System.Drawing.Size(229, 22);
            this.sairToolStripMenuItem.Text = "Sair";
            this.sairToolStripMenuItem.Click += new System.EventHandler(this.sairToolStripMenuItem_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(11, 237);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(0, 13);
            this.label1.TabIndex = 1;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(29, 146);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(0, 13);
            this.label2.TabIndex = 2;
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(32, 162);
            this.comboBox1.Margin = new System.Windows.Forms.Padding(2);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(120, 21);
            this.comboBox1.TabIndex = 3;
            this.comboBox1.Visible = false;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(55, 147);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(100, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "Gerar Midia do Lote";
            this.label3.Visible = false;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(32, 196);
            this.button1.Margin = new System.Windows.Forms.Padding(2);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(119, 28);
            this.button1.TabIndex = 5;
            this.button1.Text = "Processar Midia";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Visible = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(29, 52);
            this.label4.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(95, 13);
            this.label4.TabIndex = 7;
            this.label4.Text = "Data de Postagem";
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Location = new System.Drawing.Point(32, 69);
            this.dateTimePicker1.Margin = new System.Windows.Forms.Padding(2);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(212, 20);
            this.dateTimePicker1.TabIndex = 8;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(155, 196);
            this.button2.Margin = new System.Windows.Forms.Padding(2);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(119, 28);
            this.button2.TabIndex = 9;
            this.button2.Text = "Processar Capa";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Visible = false;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(279, 196);
            this.button3.Margin = new System.Windows.Forms.Padding(2);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(119, 28);
            this.button3.TabIndex = 10;
            this.button3.Text = "Gerar Relatório";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Visible = false;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(355, 259);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(58, 13);
            this.label5.TabIndex = 11;
            this.label5.Text = "ver: 2.1.20";
            // 
            // butProcessarCarne
            // 
            this.butProcessarCarne.Location = new System.Drawing.Point(279, 46);
            this.butProcessarCarne.Name = "butProcessarCarne";
            this.butProcessarCarne.Size = new System.Drawing.Size(119, 43);
            this.butProcessarCarne.TabIndex = 12;
            this.butProcessarCarne.Text = "Processar Carnê";
            this.butProcessarCarne.UseVisualStyleBackColor = true;
            this.butProcessarCarne.Click += new System.EventHandler(this.butProcessarCarne_Click);
            // 
            // lblValor
            // 
            this.lblValor.AutoSize = true;
            this.lblValor.Location = new System.Drawing.Point(375, 30);
            this.lblValor.Name = "lblValor";
            this.lblValor.Size = new System.Drawing.Size(23, 13);
            this.lblValor.TabIndex = 13;
            this.lblValor.Text = "%%";
            this.lblValor.Visible = false;
            // 
            // lblMsg
            // 
            this.lblMsg.AutoSize = true;
            this.lblMsg.ForeColor = System.Drawing.Color.Crimson;
            this.lblMsg.Location = new System.Drawing.Point(19, 100);
            this.lblMsg.Name = "lblMsg";
            this.lblMsg.Size = new System.Drawing.Size(13, 13);
            this.lblMsg.TabIndex = 14;
            this.lblMsg.Text = "_";
            // 
            // butProcessaCartas
            // 
            this.butProcessaCartas.Enabled = false;
            this.butProcessaCartas.Location = new System.Drawing.Point(279, 111);
            this.butProcessaCartas.Name = "butProcessaCartas";
            this.butProcessaCartas.Size = new System.Drawing.Size(118, 43);
            this.butProcessaCartas.TabIndex = 15;
            this.butProcessaCartas.Text = "Processar Cartas CCB/BACEN";
            this.butProcessaCartas.UseVisualStyleBackColor = true;
            this.butProcessaCartas.Click += new System.EventHandler(this.butProcessaCartas_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(434, 275);
            this.Controls.Add(this.butProcessaCartas);
            this.Controls.Add(this.lblMsg);
            this.Controls.Add(this.lblValor);
            this.Controls.Add(this.butProcessarCarne);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.dateTimePicker1);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "Form1";
            this.Text = "Processamento Banco Safra";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem sistemaToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem processarToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem sairToolStripMenuItem;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem1;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label4;
        public System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem2;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem3;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem4;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem5;
        private System.Windows.Forms.ToolStripMenuItem descompactarToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem processarToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem6;
        private System.Windows.Forms.ToolStripMenuItem separaErroToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem7;
        private System.Windows.Forms.ToolStripMenuItem boletoA4ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem8;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem9;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem10;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem11;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem12;
        private System.Windows.Forms.ToolStripMenuItem fase01ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem fase02ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem13;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem14;
        private System.Windows.Forms.ToolStripMenuItem processarBACENToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem Consignado;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem15;
        private System.Windows.Forms.ToolStripMenuItem processarToolStripMenuItem2;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ToolStripMenuItem processarCB7_MenuItem;
        private System.Windows.Forms.Button butProcessarCarne;
        private System.Windows.Forms.Label lblValor;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.ComponentModel.BackgroundWorker bgWorkerIndeterminada;
        private System.Windows.Forms.Label lblMsg;
        private System.Windows.Forms.Button butProcessaCartas;
        private System.Windows.Forms.ToolStripMenuItem ProcessarCartaDOCLA;
    }
}

