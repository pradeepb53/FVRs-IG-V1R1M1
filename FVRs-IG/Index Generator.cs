using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace FVRs_IG
{
    public partial class Form1 : Form
    {
             
        System.IO.StreamWriter excludedWordsFile; // New version V1R1M0 - 08/16/2017
        string excludedWordsFilePath = Application.StartupPath + @"\data\Word-List.txt"; // New version V1R1M0 - 08/16/2017

        public Form1()
        {
            InitializeComponent();
            buildWordList(); // New version V1R1M0 - 08/16/2017
            setButtonStatus();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void setButtonStatus()
        {
            buttonCreate.Hide();
            buttonClear.Hide();
            buttonSaveWords.Enabled = false; // New version V1R1M0 - 08/16/2017

            string serverName = System.Windows.Forms.SystemInformation.ComputerName;
            if ((serverName.Substring(0, 5) == "FVRSC") || (serverName.Substring(0, 3) == "WLT") || (serverName.Substring(0, 7) == "FVRsDev")) // V1R1M1 - 09/18/2017
            {
                
            }
            else
            {
                textBoxSelectFile.Enabled = false;
                buttonSaveWords.Enabled = false; // New version V1R1M0 - 08/16/2017
                textBoxExcludedWords.Enabled = false;  // New version V1R1M0 - 08/16/2017
                MessageBox.Show("Unlicensed Product - Please contact Fraser Valley Reporting Services!");
            }
        }

        private void buildWordList()
        {
           
            try
            {
                this.textBoxExcludedWords.Text = System.IO.File.ReadAllText(excludedWordsFilePath);   // New version V1R1M0 - 08/16/2017
            }
            catch (FileNotFoundException ex)  // New version V1R1M0 - 08/16/2017
            {

                MessageBox.Show("Error - Excluded word list not found! ");  // New version V1R1M0 - 08/16/2017
            } 

        }


        private void textBoxSelectFile_Click(object sender, EventArgs e)
        {
            openFileDialogSelectFile = new OpenFileDialog();
            openFileDialogSelectFile.Filter = "DOC|*.doc|DOCX|*.docx"; // V1R1M1 - 09/18/2017

            if (openFileDialogSelectFile.ShowDialog() == DialogResult.OK)
            {
                textBoxSelectFile.Text = openFileDialogSelectFile.FileName;
                textBoxSelectFile.Enabled = false;
                buttonClear.Enabled = true;
                buttonCreate.BackColor = Color.LimeGreen;
                buttonClear.Show();
                buttonCreate.Show();
                MessageBox.Show("Please exit out of all other programs....... MS Word/MS Excel......etc..");
                buttonSaveWords.Enabled = false; //new version V1R1M0 - 08/16/2017
            }

        }

        private void buttonClear_Click(object sender, EventArgs e)
        {
            textBoxSelectFile.Text = "";
            textBoxSelectFile.Enabled = true;
            buttonCreate.Hide();
            buttonClear.Enabled = false;
            buttonSaveWords.Enabled = true; //new version V1R1M0 - 08/16/2017
        }

        private void buttonCreate_Click(object sender, EventArgs e)
        {

            buttonClear.Enabled = false;
            buttonCreate.Enabled = false;
            progressBarCoreOps.Show();

            this.timerCoreOps.Interval = 10;
            this.timerCoreOps.Start();
  
            this.progressBarCoreOps.Increment(2);

            //------------------ new version V1R1M0 - 08/16/2017 --------------------------------------------------

            var excludedWordsInitialArray = File.ReadAllLines(excludedWordsFilePath);
            string[] excludedWordsFinalList = new string[excludedWordsInitialArray.Length];
            int index = 0;
            foreach (var item in excludedWordsInitialArray)
            {
                if (item.Trim() != "")
                {
                    excludedWordsFinalList[index] = item.Trim(); // Preserve original case - Upper/Lower
                    index++;
                }

            }

            //------------------------------------------------------------------------------------------------------


            this.progressBarCoreOps.Increment(5);

            IndexCore iGenerator = new IndexCore();
            this.progressBarCoreOps.Increment(40);
           
            iGenerator.processTranscript(textBoxSelectFile.Text, excludedWordsFinalList);  //new version V1R1M0 - 08/16/2017
            this.progressBarCoreOps.Increment(95);

            iGenerator.printWordIndex();

            this.timerCoreOps.Stop();
            this.timerCoreOps.Dispose();
            progressBarCoreOps.Hide();

            textBoxSelectFile.Enabled = true;
            buttonClear.Enabled = true;
            buttonCreate.BackColor = Color.LightGray;
            buttonCreate.Enabled = false;

        }

        private void timerCoreOps_Tick(object sender, EventArgs e)
        {
            this.progressBarCoreOps.Increment(1);
        }

        private void buttonSaveWords_Click(object sender, EventArgs e)    // New version V1R1M0 - 08/16/2017
        {
            string excludedWordList = this.textBoxExcludedWords.Text;  // New version V1R1M0 - 08/16/2017
            excludedWordsFile = new System.IO.StreamWriter(excludedWordsFilePath);  // New version V1R1M0 - 08/16/2017
            excludedWordsFile.WriteLine(excludedWordList);  // New version V1R1M0 - 08/16/2017
            excludedWordsFile.Close();  // New version V1R1M0 - 08/16/2017

        }

        private void textBoxExcludedWords_TextChanged(object sender, EventArgs e)
        {
            this.buttonSaveWords.Enabled = true;
        }
    }
}
