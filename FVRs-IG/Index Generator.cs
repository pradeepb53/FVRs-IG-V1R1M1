using System;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
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
            string serverName = System.Windows.Forms.SystemInformation.ComputerName;
            if ((serverName.Substring(0, 5) == "FVRSC") || (serverName.Substring(0, 3) == "WLT"))
            {
                
            }
            else
            {
                textBoxSelectFile.Enabled = false;
                buttonAddWord.Enabled = false;
                listBoxWordList.Enabled = false;
                MessageBox.Show("Unlicensed Product - Please contact Fraser Valley Reporting Services!");
            }
        }

        private void buildWordList()
        {
            excludedWordList.Add("The");
            excludedWordList.Add("the");
            excludedWordList.Add("They");
            excludedWordList.Add("they");
            excludedWordList.Add("Them");
            excludedWordList.Add("them");
            excludedWordList.Add("There");
            excludedWordList.Add("there");
            excludedWordList.Add("This");
            excludedWordList.Add("this");
            excludedWordList.Add("That");
            excludedWordList.Add("that");
            excludedWordList.Add("When");
            excludedWordList.Add("when");
            excludedWordList.Add("Where");
            excludedWordList.Add("where");
            excludedWordList.Add("What");
            excludedWordList.Add("what");

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
                MessageBox.Show("Please exit out of all other programs....... MS Word/MS Excel......etc..");
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

            this.timerCoreOps.Interval = 10;
            this.timerCoreOps.Start();
  
            this.progressBarCoreOps.Increment(2);

            string[] excludedWords = new string[excludedWordList.Count()];
            int index = 0;
            foreach (String element in excludedWordList)
            {
                excludedWords[index] = element.Trim(); // Preserve original case - Upper/Lower
                this.progressBarCoreOps.Increment(1);
                index++;
            }

            this.progressBarCoreOps.Increment(5);

            IndexCore iGenerator = new IndexCore();
            this.progressBarCoreOps.Increment(40);
            iGenerator.processTranscript(textBoxSelectFile.Text, excludedWords);

            this.progressBarCoreOps.Increment(95);

            iGenerator.printWordIndex();

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
            this.progressBarCoreOps.Increment(1);
        }

    }
}
