namespace iCAR10p
{
    partial class FormCarga
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
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(43, 54);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(242, 44);
            this.button1.TabIndex = 0;
            this.button1.Text = "Archivo Prenda";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.Carga_Archivo_Prendas);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(43, 124);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(242, 47);
            this.button2.TabIndex = 1;
            this.button2.Text = "Generar PDF tag";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.Generar_pdf_tag);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(43, 204);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(242, 43);
            this.button3.TabIndex = 2;
            this.button3.Text = "Leer facturas i-CAR";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.Leer_facturas_iCAR_obtener_patentes_facturadas);
            // 
            // FormCarga
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Name = "FormCarga";
            this.Text = "iCAR10p";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
    }
}

