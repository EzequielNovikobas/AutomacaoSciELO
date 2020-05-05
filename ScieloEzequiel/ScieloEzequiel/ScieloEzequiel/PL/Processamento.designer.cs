namespace ScieloEzequiel
{
    partial class Processamento
    {
        /// <summary>
        /// Variável de designer necessária.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpar os recursos que estão sendo usados.
        /// </summary>
        /// <param name="disposing">true se for necessário descartar os recursos gerenciados; caso contrário, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código gerado pelo Windows Form Designer

        /// <summary>
        /// Método necessário para suporte ao Designer - não modifique 
        /// o conteúdo deste método com o editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.TxtPesquisa = new System.Windows.Forms.TextBox();
            this.cmdIniciar = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(96, 13);
            this.label1.TabIndex = 7;
            this.label1.Text = "Informe a pesquisa";
            // 
            // TxtPesquisa
            // 
            this.TxtPesquisa.Location = new System.Drawing.Point(12, 25);
            this.TxtPesquisa.Name = "TxtPesquisa";
            this.TxtPesquisa.Size = new System.Drawing.Size(378, 20);
            this.TxtPesquisa.TabIndex = 6;
            // 
            // cmdIniciar
            // 
            this.cmdIniciar.Location = new System.Drawing.Point(293, 134);
            this.cmdIniciar.Name = "cmdIniciar";
            this.cmdIniciar.Size = new System.Drawing.Size(97, 37);
            this.cmdIniciar.TabIndex = 8;
            this.cmdIniciar.Text = "Iniciar";
            this.cmdIniciar.UseVisualStyleBackColor = true;
            this.cmdIniciar.Click += new System.EventHandler(this.cmdIniciar_Click);
            // 
            // Processamento
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(397, 183);
            this.Controls.Add(this.cmdIniciar);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.TxtPesquisa);
            this.Name = "Processamento";
            this.Text = "Processamento";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox TxtPesquisa;
        private System.Windows.Forms.Button cmdIniciar;
    }
}

