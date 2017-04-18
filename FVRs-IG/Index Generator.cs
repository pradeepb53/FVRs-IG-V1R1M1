using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FVRs_IG
{
    public partial class Form1 : Form
    {
        BindingList<String> excludedWordList = new BindingList<string>();
        
        public Form1()
        {
            InitializeComponent();
            setButtonStatus();
            buildWordList();
        }
        
        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void setButtonStatus()
        {
            buttonCreate.Hide();
            buttonClear.Hide();
           // progressBarCoreOps.Hide();
        }

        private void buildWordList()
        {
            excludedWordList.Add("The");
            excludedWordList.Add("They");
            excludedWordList.Add("Them");
            excludedWordList.Add("There");
            excludedWordList.Add("This");
            excludedWordList.Add("That");
            excludedWordList.Add("When");
            excludedWordList.Add("Where");
            excludedWordList.Add("What");

            listBoxWordList.DataSource = excludedWordList;

            
        }

        private void buttonAddWord_Click(object sender, EventArgs e)
        {
            AddNewWord addWords = new AddNewWord();
            addWords.ShowDialog();
            string newWord = addWords.retrieveNewWord();
            if (newWord != "")
            {
                excludedWordList.Add(newWord);
                addWords.Dispose();
            }
        }

        private void textBoxSelectFile_Click(object sender, EventArgs e)
        {
            openFileDialogSelectFile = new OpenFileDialog();
            openFileDialogSelectFile.Filter = "DOC|*.doc";

            if (openFileDialogSelectFile.ShowDialog() == DialogResult.OK)
            {
                textBoxSelectFile.Text = openFileDialogSelectFile.FileName;
                textBoxSelectFile.Enabled = false;
                buttonClear.Enabled = true;
                buttonCreate.BackColor = Color.LimeGreen;
                buttonClear.Show();
                buttonCreate.Show();
            }

        }

        private void buttonClear_Click(object sender, EventArgs e)
        {
            textBoxSelectFile.Text = "";
            textBoxSelectFile.Enabled = true;
            buttonCreate.Hide();
            buttonClear.Enabled = false;

        }

        private void buttonCreate_Click(object sender, EventArgs e)
        {

            buttonAddWord.Enabled = false;
            buttonClear.Enabled = false;
            buttonCreate.Enabled = false;
            progressBarCoreOps.Show();

            // MessageBox.Show("Please Wait.....Index is being generated!");

            this.timerCoreOps.Start();
            this.progressBarCoreOps.Increment(10);

            string[] excludedWords = new string[excludedWordList.Count()];
            int index = 0;
            foreach (String element in excludedWordList)
            {
                excludedWords[index] = element.Trim(); // Preserve original case - Upper/Lower
                this.progressBarCoreOps.Increment(5);
                index++;
            }

            this.progressBarCoreOps.Increment(25);

            IndexCore iGenerator = new IndexCore();
            this.progressBarCoreOps.Increment(25);
            iGenerator.processTranscript(textBoxSelectFile.Text, excludedWords);

            //MessageBox.Show("index page being saved!!");

            this.progressBarCoreOps.Increment(20);

            iGenerator.printWordIndex();

            this.progressBarCoreOps.Increment(5);
            this.timerCoreOps.Stop();
            this.timerCoreOps.Dispose();
            progressBarCoreOps.Hide();

            buttonAddWord.Enabled = true;
            textBoxSelectFile.Enabled = true;
            buttonClear.Enabled = true;
            buttonCreate.BackColor = Color.LightGray;
            buttonCreate.Enabled = false;

        }

        private void timerCoreOps_Tick(object sender, EventArgs e)
        {
            this.progressBarCoreOps.Increment(10);
        }

        private async void buttonProgress_Click(object sender, EventArgs e)
        {
            // progressBarCoreOps.Maximum = 100;
            // progressBarCoreOps.Step = 1;

            //var progress = new Progress<int>(v => { progressBarCoreOps.Value = v; });
            // await Task.Run(() => DoWork(progress));
            var progress = new Progress<int>(ReportProgress);
            await copyFiles(progress);
        }

        private async Task copyFiles(IProgress<int> progress)
        {
            for (int i = 0; i < 11; i++)
            {
                await Task.Run(() => { Thread.Sleep(1000); });
                progress.Report(i);
            }
        }

        private void ReportProgress(int obj)
        {
            progressBarCoreOps.Increment(obj); //
        }

        private void setProgressBar(int obj)
        {
            progressBarCoreOps.Increment(obj);
        }

        private void DoWork(IProgress<int> progress)
        {
            for (int i = 0; i < 100000; i++)
            {

                calculateJ(i);

                if (progress != null)
                {

                    progress.Report(i);
                    progress.Report((i + 1) * 100 / 100000);
                }
            }
        }


        private void calculateJ(int i)
        {
            double pow = Math.Pow(i, i);
        }
    }
}
