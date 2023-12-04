namespace SolidworksApiProject.Chapter5
{
    partial class Chapter5Form
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
            this.btn_newapp = new System.Windows.Forms.Button();
            this.btn_getapp = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btn_newapp
            // 
            this.btn_newapp.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            this.btn_newapp.Location = new System.Drawing.Point(6, 5);
            this.btn_newapp.Name = "btn_newapp";
            this.btn_newapp.Size = new System.Drawing.Size(154, 23);
            this.btn_newapp.TabIndex = 0;
            this.btn_newapp.Text = "新建Solidworks应用程序";
            this.btn_newapp.UseVisualStyleBackColor = false;
            this.btn_newapp.Click += new System.EventHandler(this.btn_newapp_Click);
            // 
            // btn_getapp
            // 
            this.btn_getapp.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            this.btn_getapp.Location = new System.Drawing.Point(166, 5);
            this.btn_getapp.Name = "btn_getapp";
            this.btn_getapp.Size = new System.Drawing.Size(154, 23);
            this.btn_getapp.TabIndex = 1;
            this.btn_getapp.Text = "获取Solidworks应用程序";
            this.btn_getapp.UseVisualStyleBackColor = false;
            this.btn_getapp.Click += new System.EventHandler(this.btn_getapp_Click);
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            this.button1.Location = new System.Drawing.Point(6, 34);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(154, 23);
            this.button1.TabIndex = 2;
            this.button1.Text = "5.3.2实例分析";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // Chapter5Form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(328, 60);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.btn_getapp);
            this.Controls.Add(this.btn_newapp);
            this.Name = "Chapter5Form";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "第5章 应用程序对象IsldWorks";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btn_newapp;
        private System.Windows.Forms.Button btn_getapp;
        private System.Windows.Forms.Button button1;
    }
}