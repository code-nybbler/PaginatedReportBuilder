﻿namespace PaginatedReportBuilder
{
    partial class PaginatedReportBuilderControl
    {
        /// <summary> 
        /// Variable nécessaire au concepteur.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Nettoyage des ressources utilisées.
        /// </summary>
        /// <param name="disposing">true si les ressources managées doivent être supprimées ; sinon, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Code généré par le Concepteur de composants

        /// <summary> 
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas 
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(PaginatedReportBuilderControl));
            this.toolStripMenu = new System.Windows.Forms.ToolStrip();
            this.tsbClose = new System.Windows.Forms.ToolStripButton();
            this.tssSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.box_entitySelect = new System.Windows.Forms.ComboBox();
            this.lst_forms = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.btn_generate = new System.Windows.Forms.Button();
            this.txt_formxml = new System.Windows.Forms.RichTextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txt_reportxml = new System.Windows.Forms.RichTextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.btn_download = new System.Windows.Forms.Button();
            this.btn_loadEntities = new System.Windows.Forms.ToolStripButton();
            this.toolStripMenu.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // toolStripMenu
            // 
            this.toolStripMenu.ImageScalingSize = new System.Drawing.Size(24, 24);
            this.toolStripMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsbClose,
            this.tssSeparator1,
            this.btn_loadEntities});
            this.toolStripMenu.Location = new System.Drawing.Point(0, 0);
            this.toolStripMenu.Name = "toolStripMenu";
            this.toolStripMenu.Size = new System.Drawing.Size(1314, 25);
            this.toolStripMenu.TabIndex = 4;
            this.toolStripMenu.Text = "toolStrip1";
            // 
            // tsbClose
            // 
            this.tsbClose.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.tsbClose.Name = "tsbClose";
            this.tsbClose.Size = new System.Drawing.Size(40, 22);
            this.tsbClose.Text = "Close";
            this.tsbClose.Click += new System.EventHandler(this.tsbClose_Click);
            // 
            // tssSeparator1
            // 
            this.tssSeparator1.Name = "tssSeparator1";
            this.tssSeparator1.Size = new System.Drawing.Size(6, 25);
            // 
            // box_entitySelect
            // 
            this.box_entitySelect.Enabled = false;
            this.box_entitySelect.FormattingEnabled = true;
            this.box_entitySelect.Location = new System.Drawing.Point(6, 45);
            this.box_entitySelect.Name = "box_entitySelect";
            this.box_entitySelect.Size = new System.Drawing.Size(212, 21);
            this.box_entitySelect.TabIndex = 5;
            this.box_entitySelect.SelectedIndexChanged += new System.EventHandler(this.box_entitySelect_SelectedIndexChanged);
            // 
            // lst_forms
            // 
            this.lst_forms.Enabled = false;
            this.lst_forms.FormattingEnabled = true;
            this.lst_forms.Location = new System.Drawing.Point(6, 85);
            this.lst_forms.Name = "lst_forms";
            this.lst_forms.Size = new System.Drawing.Size(211, 82);
            this.lst_forms.TabIndex = 6;
            this.lst_forms.SelectedIndexChanged += new System.EventHandler(this.lst_forms_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(3, 29);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(80, 13);
            this.label1.TabIndex = 8;
            this.label1.Text = "Entity Selection";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(3, 69);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(77, 13);
            this.label2.TabIndex = 9;
            this.label2.Text = "Form Selection";
            // 
            // btn_generate
            // 
            this.btn_generate.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn_generate.Enabled = false;
            this.btn_generate.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_generate.ForeColor = System.Drawing.SystemColors.Highlight;
            this.btn_generate.Location = new System.Drawing.Point(6, 173);
            this.btn_generate.Name = "btn_generate";
            this.btn_generate.Size = new System.Drawing.Size(212, 23);
            this.btn_generate.TabIndex = 11;
            this.btn_generate.Text = "Generate Report";
            this.btn_generate.UseVisualStyleBackColor = true;
            this.btn_generate.Click += new System.EventHandler(this.btn_generate_Click);
            // 
            // txt_formxml
            // 
            this.txt_formxml.Location = new System.Drawing.Point(224, 45);
            this.txt_formxml.Name = "txt_formxml";
            this.txt_formxml.Size = new System.Drawing.Size(486, 596);
            this.txt_formxml.TabIndex = 12;
            this.txt_formxml.Text = "";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(224, 28);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(50, 13);
            this.label4.TabIndex = 13;
            this.label4.Text = "Form Xml";
            // 
            // txt_reportxml
            // 
            this.txt_reportxml.Location = new System.Drawing.Point(728, 45);
            this.txt_reportxml.Name = "txt_reportxml";
            this.txt_reportxml.Size = new System.Drawing.Size(486, 596);
            this.txt_reportxml.TabIndex = 14;
            this.txt_reportxml.Text = "";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(728, 29);
            this.label5.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(59, 13);
            this.label5.TabIndex = 15;
            this.label5.Text = "Report Xml";
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackgroundImage = global::PaginatedReportBuilder.Properties.Resources.sagemodeicon8080;
            this.pictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.pictureBox1.Enabled = false;
            this.pictureBox1.Location = new System.Drawing.Point(6, 201);
            this.pictureBox1.Margin = new System.Windows.Forms.Padding(2);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(211, 210);
            this.pictureBox1.TabIndex = 16;
            this.pictureBox1.TabStop = false;
            // 
            // btn_download
            // 
            this.btn_download.BackColor = System.Drawing.SystemColors.Control;
            this.btn_download.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn_download.Enabled = false;
            this.btn_download.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_download.ForeColor = System.Drawing.SystemColors.Highlight;
            this.btn_download.Location = new System.Drawing.Point(1094, 646);
            this.btn_download.Margin = new System.Windows.Forms.Padding(2);
            this.btn_download.Name = "btn_download";
            this.btn_download.Size = new System.Drawing.Size(120, 23);
            this.btn_download.TabIndex = 17;
            this.btn_download.Text = "Download Report";
            this.btn_download.UseVisualStyleBackColor = false;
            this.btn_download.Click += new System.EventHandler(this.btn_download_Click);
            // 
            // btn_loadEntities
            // 
            this.btn_loadEntities.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.btn_loadEntities.Image = ((System.Drawing.Image)(resources.GetObject("btn_loadEntities.Image")));
            this.btn_loadEntities.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btn_loadEntities.Name = "btn_loadEntities";
            this.btn_loadEntities.Size = new System.Drawing.Size(78, 22);
            this.btn_loadEntities.Text = "Load Entities";
            this.btn_loadEntities.Click += new System.EventHandler(this.btn_loadEntities_Click);
            // 
            // PaginatedReportBuilderControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.btn_download);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txt_reportxml);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.txt_formxml);
            this.Controls.Add(this.btn_generate);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lst_forms);
            this.Controls.Add(this.box_entitySelect);
            this.Controls.Add(this.toolStripMenu);
            this.Name = "PaginatedReportBuilderControl";
            this.Size = new System.Drawing.Size(1314, 757);
            this.Load += new System.EventHandler(this.PaginatedReportGeneratorControl_Load);
            this.toolStripMenu.ResumeLayout(false);
            this.toolStripMenu.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.ToolStrip toolStripMenu;
        private System.Windows.Forms.ToolStripButton tsbClose;
        private System.Windows.Forms.ToolStripSeparator tssSeparator1;
        private System.Windows.Forms.ComboBox box_entitySelect;
        private System.Windows.Forms.ListBox lst_forms;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btn_generate;
        private System.Windows.Forms.RichTextBox txt_formxml;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.RichTextBox txt_reportxml;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Button btn_download;
        private System.Windows.Forms.ToolStripButton btn_loadEntities;
    }
}
