
namespace W2E
{
    partial class FormW2E
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        /// <summary>
        /// Возвращает элемент текстового поля для лога
        /// </summary>
        /// <returns></returns>
        public System.Windows.Forms.TextBox getLogTextBox()
        {
            return this.logTextBox;
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.iFileLabel = new System.Windows.Forms.Label();
            this.iFileButton = new System.Windows.Forms.Button();
            this.iFileTextBox = new System.Windows.Forms.TextBox();
            this.startButton = new System.Windows.Forms.Button();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.logTextBox = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // iFileLabel
            // 
            this.iFileLabel.AutoSize = true;
            this.iFileLabel.Location = new System.Drawing.Point(12, 34);
            this.iFileLabel.Name = "iFileLabel";
            this.iFileLabel.Size = new System.Drawing.Size(232, 13);
            this.iFileLabel.TabIndex = 0;
            this.iFileLabel.Text = "Выберите файл Word, содержащий таблицы";
            // 
            // iFileButton
            // 
            this.iFileButton.Location = new System.Drawing.Point(397, 47);
            this.iFileButton.Name = "iFileButton";
            this.iFileButton.Size = new System.Drawing.Size(75, 23);
            this.iFileButton.TabIndex = 1;
            this.iFileButton.Text = "Выбрать";
            this.iFileButton.UseVisualStyleBackColor = true;
            this.iFileButton.Click += new System.EventHandler(this.iFileButtonClick);
            // 
            // iFileTextBox
            // 
            this.iFileTextBox.Location = new System.Drawing.Point(15, 50);
            this.iFileTextBox.Name = "iFileTextBox";
            this.iFileTextBox.ReadOnly = true;
            this.iFileTextBox.Size = new System.Drawing.Size(346, 20);
            this.iFileTextBox.TabIndex = 2;
            this.iFileTextBox.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // startButton
            // 
            this.startButton.Location = new System.Drawing.Point(397, 276);
            this.startButton.Name = "startButton";
            this.startButton.Size = new System.Drawing.Size(75, 23);
            this.startButton.TabIndex = 3;
            this.startButton.Text = "Начать";
            this.startButton.UseVisualStyleBackColor = true;
            this.startButton.Click += new System.EventHandler(this.startButtonClick);
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(15, 235);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(457, 23);
            this.progressBar.TabIndex = 4;
            this.progressBar.Tag = "";
            // 
            // logTextBox
            // 
            this.logTextBox.Location = new System.Drawing.Point(15, 89);
            this.logTextBox.Multiline = true;
            this.logTextBox.Name = "logTextBox";
            this.logTextBox.ReadOnly = true;
            this.logTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.logTextBox.Size = new System.Drawing.Size(457, 126);
            this.logTextBox.TabIndex = 5;
            this.logTextBox.TextChanged += new System.EventHandler(this.textBox2_TextChanged);
            // 
            // FormW2E
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(484, 311);
            this.Controls.Add(this.logTextBox);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.startButton);
            this.Controls.Add(this.iFileTextBox);
            this.Controls.Add(this.iFileButton);
            this.Controls.Add(this.iFileLabel);
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(500, 350);
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(500, 350);
            this.Name = "FormW2E";
            this.Text = "Word 2 Excel";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label iFileLabel;
        private System.Windows.Forms.Button iFileButton;
        private System.Windows.Forms.TextBox iFileTextBox;
        private System.Windows.Forms.Button startButton;
        private string filePath;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.TextBox logTextBox;
    }
}

