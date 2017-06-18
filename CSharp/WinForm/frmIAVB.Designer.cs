namespace IAVB
{
    partial class frmIAVB
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
            if (disposing && (m_synth != null)) // CA2213
            {
                m_synth.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Code généré par le Concepteur Windows Form

        /// <summary>
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmIAVB));
            this.cmdSuivant = new System.Windows.Forms.Button();
            this.cmdGo = new System.Windows.Forms.Button();
            this.listAssert = new System.Windows.Forms.ListBox();
            this.listIA = new System.Windows.Forms.ListBox();
            this.textInput = new System.Windows.Forms.TextBox();
            this.listParole = new System.Windows.Forms.ListBox();
            this.SuspendLayout();
            // 
            // cmdSuivant
            // 
            this.cmdSuivant.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.cmdSuivant.BackColor = System.Drawing.SystemColors.Control;
            this.cmdSuivant.Cursor = System.Windows.Forms.Cursors.Default;
            this.cmdSuivant.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmdSuivant.ForeColor = System.Drawing.SystemColors.ControlText;
            this.cmdSuivant.Location = new System.Drawing.Point(464, 238);
            this.cmdSuivant.Name = "cmdSuivant";
            this.cmdSuivant.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.cmdSuivant.Size = new System.Drawing.Size(33, 33);
            this.cmdSuivant.TabIndex = 9;
            this.cmdSuivant.Text = "->";
            this.cmdSuivant.UseVisualStyleBackColor = false;
            this.cmdSuivant.Click += new System.EventHandler(this.cmdSuivant_Click);
            // 
            // cmdGo
            // 
            this.cmdGo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.cmdGo.BackColor = System.Drawing.SystemColors.Control;
            this.cmdGo.Cursor = System.Windows.Forms.Cursors.Default;
            this.cmdGo.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmdGo.ForeColor = System.Drawing.SystemColors.ControlText;
            this.cmdGo.Location = new System.Drawing.Point(416, 238);
            this.cmdGo.Name = "cmdGo";
            this.cmdGo.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.cmdGo.Size = new System.Drawing.Size(33, 33);
            this.cmdGo.TabIndex = 8;
            this.cmdGo.Text = "!";
            this.cmdGo.UseVisualStyleBackColor = false;
            this.cmdGo.Click += new System.EventHandler(this.cmdGo_Click);
            // 
            // listAssert
            // 
            this.listAssert.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.listAssert.BackColor = System.Drawing.SystemColors.Window;
            this.listAssert.Cursor = System.Windows.Forms.Cursors.Default;
            this.listAssert.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.listAssert.ForeColor = System.Drawing.SystemColors.WindowText;
            this.listAssert.ItemHeight = 16;
            this.listAssert.Location = new System.Drawing.Point(8, 294);
            this.listAssert.Name = "listAssert";
            this.listAssert.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.listAssert.Size = new System.Drawing.Size(489, 84);
            this.listAssert.TabIndex = 7;
            this.listAssert.SelectedIndexChanged += new System.EventHandler(this.listAssert_SelectedIndexChanged);
            this.listAssert.DoubleClick += new System.EventHandler(this.listAssert_DoubleClick);
            // 
            // listIA
            // 
            this.listIA.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.listIA.BackColor = System.Drawing.Color.White;
            this.listIA.Cursor = System.Windows.Forms.Cursors.Default;
            this.listIA.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.listIA.ForeColor = System.Drawing.SystemColors.WindowText;
            this.listIA.ItemHeight = 16;
            this.listIA.Location = new System.Drawing.Point(8, 86);
            this.listIA.Name = "listIA";
            this.listIA.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.listIA.Size = new System.Drawing.Size(489, 116);
            this.listIA.TabIndex = 6;
            this.listIA.SelectedIndexChanged += new System.EventHandler(this.listIA_SelectedIndexChanged);
            this.listIA.DoubleClick += new System.EventHandler(this.listIA_DoubleClick);
            // 
            // textInput
            // 
            this.textInput.AcceptsReturn = true;
            this.textInput.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textInput.BackColor = System.Drawing.SystemColors.Window;
            this.textInput.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.textInput.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textInput.ForeColor = System.Drawing.SystemColors.WindowText;
            this.textInput.Location = new System.Drawing.Point(8, 230);
            this.textInput.MaxLength = 0;
            this.textInput.Multiline = true;
            this.textInput.Name = "textInput";
            this.textInput.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.textInput.Size = new System.Drawing.Size(393, 41);
            this.textInput.TabIndex = 5;
            this.textInput.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textInput_KeyPress);
            // 
            // listParole
            // 
            this.listParole.BackColor = System.Drawing.Color.White;
            this.listParole.Cursor = System.Windows.Forms.Cursors.Default;
            this.listParole.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.listParole.ForeColor = System.Drawing.SystemColors.WindowText;
            this.listParole.ItemHeight = 16;
            this.listParole.Location = new System.Drawing.Point(12, 12);
            this.listParole.Name = "listParole";
            this.listParole.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.listParole.Size = new System.Drawing.Size(321, 52);
            this.listParole.TabIndex = 10;
            this.listParole.SelectedIndexChanged += new System.EventHandler(this.listParole_SelectedIndexChanged);
            // 
            // frmIAVB
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(506, 401);
            this.Controls.Add(this.listParole);
            this.Controls.Add(this.cmdSuivant);
            this.Controls.Add(this.cmdGo);
            this.Controls.Add(this.listAssert);
            this.Controls.Add(this.listIA);
            this.Controls.Add(this.textInput);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmIAVB";
            this.Text = "IAVB en C#-WinForm";
            this.Load += new System.EventHandler(this.frmIAVB_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        
        #endregion

        public System.Windows.Forms.Button cmdSuivant;
        public System.Windows.Forms.Button cmdGo;
        public System.Windows.Forms.ListBox listAssert;
        public System.Windows.Forms.ListBox listIA;
        public System.Windows.Forms.TextBox textInput;
        public System.Windows.Forms.ListBox listParole;

    }
}

