namespace FirstProject
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
            this.otLbl = new System.Windows.Forms.Label();
            this.doLbl = new System.Windows.Forms.Label();
            this.OkBtn = new System.Windows.Forms.Button();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.dateTimePicker2 = new System.Windows.Forms.DateTimePicker();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.ExportToExelBtn = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // otLbl
            // 
            this.otLbl.AutoSize = true;
            this.otLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.otLbl.Location = new System.Drawing.Point(48, 66);
            this.otLbl.Name = "otLbl";
            this.otLbl.Size = new System.Drawing.Size(53, 29);
            this.otLbl.TabIndex = 0;
            this.otLbl.Text = "От:";
            // 
            // doLbl
            // 
            this.doLbl.AutoSize = true;
            this.doLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.doLbl.Location = new System.Drawing.Point(392, 64);
            this.doLbl.Name = "doLbl";
            this.doLbl.Size = new System.Drawing.Size(56, 29);
            this.doLbl.TabIndex = 1;
            this.doLbl.Text = "До:";
            // 
            // OkBtn
            // 
            this.OkBtn.BackColor = System.Drawing.Color.White;
            this.OkBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.OkBtn.Location = new System.Drawing.Point(767, 57);
            this.OkBtn.Name = "OkBtn";
            this.OkBtn.Size = new System.Drawing.Size(134, 42);
            this.OkBtn.TabIndex = 4;
            this.OkBtn.Text = "OK";
            this.OkBtn.UseVisualStyleBackColor = false;
            this.OkBtn.Click += new System.EventHandler(this.OkBtn_Click);
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.CalendarFont = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.dateTimePicker1.Location = new System.Drawing.Point(107, 71);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(257, 22);
            this.dateTimePicker1.TabIndex = 6;
            // 
            // dateTimePicker2
            // 
            this.dateTimePicker2.CalendarFont = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.dateTimePicker2.Location = new System.Drawing.Point(454, 69);
            this.dateTimePicker2.Name = "dateTimePicker2";
            this.dateTimePicker2.Size = new System.Drawing.Size(263, 22);
            this.dateTimePicker2.TabIndex = 7;
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(31, 132);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersWidth = 51;
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(1175, 490);
            this.dataGridView1.TabIndex = 8;
            // 
            // ExportToExelBtn
            // 
            this.ExportToExelBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.ExportToExelBtn.Location = new System.Drawing.Point(945, 57);
            this.ExportToExelBtn.Name = "ExportToExelBtn";
            this.ExportToExelBtn.Size = new System.Drawing.Size(252, 42);
            this.ExportToExelBtn.TabIndex = 9;
            this.ExportToExelBtn.Text = "Экспорт в Excel";
            this.ExportToExelBtn.UseVisualStyleBackColor = true;
            this.ExportToExelBtn.Click += new System.EventHandler(this.ExportToExelBtn_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(1239, 650);
            this.Controls.Add(this.ExportToExelBtn);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.dateTimePicker2);
            this.Controls.Add(this.dateTimePicker1);
            this.Controls.Add(this.OkBtn);
            this.Controls.Add(this.doLbl);
            this.Controls.Add(this.otLbl);
            this.Name = "Form1";
            this.Text = "ИЗВЛЕЧЕНИЕ ДАННЫХ ИЗ БД";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label otLbl;
        private System.Windows.Forms.Label doLbl;
        private System.Windows.Forms.Button OkBtn;
        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.DateTimePicker dateTimePicker2;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button ExportToExelBtn;
    }
}

