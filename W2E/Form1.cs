using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace W2E
{
    public partial class FormW2E : Form
    {
        public FormW2E()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Событие, вызываемое кликом кнопки выбора файла. Открывает диалог выбора файла
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void iFileButtonClick(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "doc files (*.doc)|*.doc|docx files (*.docx)|*.docx";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    this.filePath = openFileDialog.FileName;
                    this.iFileTextBox.Text = this.filePath;
                }
            }
        }

        /// <summary>
        /// Событие, вызываемое кликом кнопки начала обработки. 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void startButtonClick(object sender, EventArgs e)
        {
            this.iFileButton.Enabled = false;
            this.progressBar.Style = ProgressBarStyle.Marquee;
            this.progressBar.MarqueeAnimationSpeed = 30;
            this.startButton.Enabled = false;

            // Начинаем перенос содержимого Word в Excel
            translator tr = new translator(this);
            await Task.Run(() => tr.Startup(this.filePath));

            this.iFileButton.Enabled = true;
            this.progressBar.Style = ProgressBarStyle.Blocks;
            this.progressBar.MarqueeAnimationSpeed = 30;
            this.startButton.Enabled = true;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
