using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;

namespace FathersApp
{
    public partial class Form1 : Form
    {
        List<Panel> listPanel;
        List<Panel> listPanelIII;
        private int ind;
        private int indIII;
        DataTable table;
        private string ExcelFile;

        public Form1()
        {
            listPanel = new List<Panel>();
            listPanelIII = new List<Panel>();
            table = new DataTable();
            InitializeComponent();
            this.ExcelFile = string.Empty;

            //should be at the end !!!
            InitDataGrid();
        }

        private void BtnNas_Click(object sender, EventArgs e)
        {
            if(ind < listPanel.Count - 1)
            {
                listPanel[++ind].BringToFront();
            }
        }

        private void BtnPop_Click(object sender, EventArgs e)
        {
            if(ind> 0)
            {
                listPanel[--ind].BringToFront();
            }
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            this.listPanel.Add(this.panel1);
            this.listPanel.Add(this.panel2);
            this.listPanel.Add(this.panel3);
            this.listPanel.Add(this.panel4);
            this.listPanel.Add(this.panel5);
            this.listPanel.Add(this.panel6);
            this.listPanel[ind].BringToFront();

            this.listPanelIII.Add(this.panel7);
            this.listPanelIII.Add(this.panel8);
            this.listPanelIII.Add(this.panel9);
            this.listPanelIII[indIII].BringToFront();

            txtFile.Text = "Wybierz bazę danych";
        }

        private void AddNewRow(Student student)
        {
            var datarow = table.NewRow();
            var enumerator = student.ListOfTasks.GetEnumerator();
            var index = 2;
            datarow[0] = student.SchoolId1;
            datarow[1] = student.StudentId1;
            while (enumerator.MoveNext())
            {
                var currentEnumerator = enumerator.Current;
                datarow[index++] = currentEnumerator.Value;
            }
            table.Rows.Add(datarow);

            dataGridView1.DataSource = table;
        }

        private void InitDataGrid()
        {
            var tasks = new Dictionary<string, string>();
            var student = new Student(string.Empty, string.Empty, tasks);
            LoadTasks(student);
            AddNewRow(student, true);
        }

        private void AddNewRow(Student student, bool AddFirstColumn)
        {
            if(AddFirstColumn)
            {
                var enumerator = student.ListOfTasks.GetEnumerator();
                table.Columns.Add("Kod szkoły", typeof(string));
                table.Columns.Add("Kod ucznia", typeof(string));
                while (enumerator.MoveNext())
                {
                    var currentEnumerator = enumerator.Current;
                    table.Columns.Add(currentEnumerator.Key, typeof(string));
                }
                dataGridView1.DataSource = table;
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            var tasks = new Dictionary<string, string>();
            var student = new Student(txtSzkola.Text, txtUczen.Text, tasks);
            LoadTasks(student);
            AddNewRow(student);
            RefreshForm();
            txtFile.Text = "Wprowadziles nowego ucznia: " + txtUczen.Text;
        }

        private void RefreshForm()
        {
            foreach(int i in CheckedLstBox1.CheckedIndices)
            {
                CheckedLstBox1.SetItemCheckState(i, CheckState.Unchecked);
            }

            foreach (int i in checkedListBox1.CheckedIndices)
            {
                checkedListBox1.SetItemCheckState(i, CheckState.Unchecked);
            }

            foreach (int i in checkedListBox2.CheckedIndices)
            {
                checkedListBox2.SetItemCheckState(i, CheckState.Unchecked);
            }

            foreach (int i in lbSex.CheckedIndices)
            {
                lbSex.SetItemCheckState(i, CheckState.Unchecked);
            }

            foreach (int i in lbAccomodation.CheckedIndices)
            {
                lbAccomodation.SetItemCheckState(i, CheckState.Unchecked);
            }

            foreach (int i in lbProfile.CheckedIndices)
            {
                lbProfile.SetItemCheckState(i, CheckState.Unchecked);
            }

            foreach (int i in lbMother.CheckedIndices)
            {
                lbMother.SetItemCheckState(i, CheckState.Unchecked);
            }

            foreach (int i in lbFather.CheckedIndices)
            {
                lbFather.SetItemCheckState(i, CheckState.Unchecked);
            }

            cmb11.ResetText();
            cmb12.ResetText();
            cmb13.ResetText();
            cmb14.ResetText();
            cmb15.ResetText();

            cmbI1.ResetText();
            cmbI2.ResetText();
            cmbI3.ResetText();
            cmbI4.ResetText();
            cmbI5.ResetText();
            cmbI6.ResetText();
            cmbI8.ResetText();
            cmbI7.ResetText();
            cmbI9.ResetText();
            cmbI10.ResetText();
            cmbI11.ResetText();

            cmbII1.ResetText();
            cmbII2.ResetText();
            cmbII3.ResetText();
            cmbII4.ResetText();
            cmbII5.ResetText();
            cmbII6.ResetText();
            cmbII7.ResetText();
            cmbII8.ResetText();
            cmbII9.ResetText();
            cmbII10.ResetText();
            cmbII11.ResetText();
            cmbII12.ResetText();
            cmbII13.ResetText();
            cmbII14.ResetText();

            cmbIII1.ResetText();
            cmbIII2.ResetText();
            cmbIII3.ResetText();
            cmbIII4.ResetText();
            cmbIII5.ResetText();
            cmbIII6.ResetText();
            cmbIII7.ResetText();
            cmbIII8.ResetText();
            cmbIII9.ResetText();
            cmbIII10.ResetText();

            cmb31.ResetText();
            cmb32.ResetText();
            cmb33.ResetText();
            cmb34.ResetText();
            cmb35.ResetText();
            cmb36.ResetText();
            cmb37.ResetText();
            cmb38.ResetText();
            cmb39.ResetText();
            cmb310.ResetText();
            cmb311.ResetText();
            cmb312.ResetText();
            cmb313.ResetText();
            cmb314.ResetText();
            cmb315.ResetText();
            cmb316.ResetText();
            cmb317.ResetText();
            cmb318.ResetText();
            cmb319.ResetText();
            cmb320.ResetText();
        }

        private void LoadTasks(Student student)
        {
            //Część I
            student.ListOfTasks.Add("I.1", CheckedLstBox1.GetItemChecked(0).ToString());
            student.ListOfTasks.Add("I.2", CheckedLstBox1.GetItemChecked(1).ToString());
            student.ListOfTasks.Add("I.3", CheckedLstBox1.GetItemChecked(2).ToString());
            student.ListOfTasks.Add("I.4", CheckedLstBox1.GetItemChecked(3).ToString());
            student.ListOfTasks.Add("I.5", CheckedLstBox1.GetItemChecked(4).ToString());
            student.ListOfTasks.Add("I.6", CheckedLstBox1.GetItemChecked(5).ToString());
            student.ListOfTasks.Add("I.7", CheckedLstBox1.GetItemChecked(6).ToString());
            student.ListOfTasks.Add("I.8", CheckedLstBox1.GetItemChecked(7).ToString());
            student.ListOfTasks.Add("I.9", CheckedLstBox1.GetItemChecked(8).ToString());
            student.ListOfTasks.Add("I.10", CheckedLstBox1.GetItemChecked(9).ToString());

            student.ListOfTasks.Add("I.11", cmb11.Text);
            student.ListOfTasks.Add("I.12", cmb12.Text);
            student.ListOfTasks.Add("I.13", cmb13.Text);
            student.ListOfTasks.Add("I.14", cmb14.Text);

            student.ListOfTasks.Add("I.15", cmb14.Text);

            //Część II.I
            student.ListOfTasks.Add("II.I.1", cmbI1.Text);
            student.ListOfTasks.Add("II.I.2", cmbI2.Text);
            student.ListOfTasks.Add("II.I.3", cmbI3.Text);
            student.ListOfTasks.Add("II.I.4", cmbI4.Text);
            student.ListOfTasks.Add("II.I.5", cmbI5.Text);
            student.ListOfTasks.Add("II.I.6", cmbI6.Text);
            student.ListOfTasks.Add("II.I.7", cmbI7.Text);
            student.ListOfTasks.Add("II.I.8", cmbI8.Text);
            student.ListOfTasks.Add("II.I.9", cmbI9.Text);
            student.ListOfTasks.Add("II.I.10", cmbI10.Text);
            student.ListOfTasks.Add("II.I.11", cmbI11.Text);

            //II.II
            student.ListOfTasks.Add("II.II.1", cmbII1.Text);
            student.ListOfTasks.Add("II.II.2", cmbII2.Text);
            student.ListOfTasks.Add("II.II.3", cmbII3.Text);
            student.ListOfTasks.Add("II.II.4", cmbII4.Text);
            student.ListOfTasks.Add("II.II.5", cmbII5.Text);
            student.ListOfTasks.Add("II.II.6", cmbII6.Text);
            student.ListOfTasks.Add("II.II.7", cmbII7.Text);
            student.ListOfTasks.Add("II.II.8", cmbII8.Text);
            student.ListOfTasks.Add("II.II.9", cmbII9.Text);
            student.ListOfTasks.Add("II.II.10", cmbII10.Text);
            student.ListOfTasks.Add("II.II.11", cmbII11.Text);
            student.ListOfTasks.Add("II.II.12", cmbII12.Text);
            student.ListOfTasks.Add("II.II.13", cmbII13.Text);
            student.ListOfTasks.Add("II.II.14", cmbII14.Text);
            
            //II.III
            student.ListOfTasks.Add("II.III.1", cmbIII1.Text);
            student.ListOfTasks.Add("II.III.2", cmbIII2.Text);
            student.ListOfTasks.Add("II.III.3", cmbIII3.Text);
            student.ListOfTasks.Add("II.III.4", cmbIII4.Text);
            student.ListOfTasks.Add("II.III.5", cmbIII5.Text);
            student.ListOfTasks.Add("II.III.6", cmbIII6.Text);
            student.ListOfTasks.Add("II.III.7", cmbIII7.Text);
            student.ListOfTasks.Add("II.III.8", cmbIII8.Text);
            student.ListOfTasks.Add("II.III.9", cmbIII9.Text);
            student.ListOfTasks.Add("II.III.10", cmbIII10.Text);

            //II.IV
            student.ListOfTasks.Add("II.IV.1", checkedListBox1.SelectedIndex.ToString());

            //II.V
            student.ListOfTasks.Add("II.V.1", checkedListBox2.GetItemChecked(0).ToString());
            student.ListOfTasks.Add("II.V.2", checkedListBox2.GetItemChecked(1).ToString());
            student.ListOfTasks.Add("II.V.3", checkedListBox2.GetItemChecked(2).ToString());
            student.ListOfTasks.Add("II.V.4", checkedListBox2.GetItemChecked(3).ToString());
            student.ListOfTasks.Add("II.V.5", checkedListBox2.GetItemChecked(4).ToString());

            //II.VI
            student.ListOfTasks.Add("II.VI.1", lbSex.SelectedIndex.ToString());
            student.ListOfTasks.Add("II.VI.2", lbAccomodation.SelectedIndex.ToString());
            student.ListOfTasks.Add("II.VI.3", lbFather.SelectedIndex.ToString());
            student.ListOfTasks.Add("II.VI.4", lbMother.SelectedIndex.ToString());
            if (lbProfile.SelectedIndex == 4)
                student.ListOfTasks.Add("II.VI.5", txtInny.Text.ToString());
            else
                student.ListOfTasks.Add("II.VI.5", lbProfile.SelectedIndex.ToString());
            student.ListOfTasks.Add("II.VI.6", txtOcena.Text.ToString());
            student.ListOfTasks.Add("II.VI.7", txtTeachers.Text.ToString());

            ///III
            student.ListOfTasks.Add("III.1", cmb31.Text);
            student.ListOfTasks.Add("III.2", cmb32.Text);
            student.ListOfTasks.Add("III.3", cmb33.Text);
            student.ListOfTasks.Add("III.4", cmb34.Text);
            student.ListOfTasks.Add("III.5", cmb35.Text);
            student.ListOfTasks.Add("III.6", cmb36.Text);
            student.ListOfTasks.Add("III.7", cmb37.Text);
            student.ListOfTasks.Add("III.8", cmb38.Text);
            student.ListOfTasks.Add("III.9", cmb39.Text);
            student.ListOfTasks.Add("III.10", cmb310.Text);
            student.ListOfTasks.Add("III.11", cmb311.Text);
            student.ListOfTasks.Add("III.12", cmb312.Text);
            student.ListOfTasks.Add("III.13", cmb313.Text);
            student.ListOfTasks.Add("III.14", cmb314.Text);
            student.ListOfTasks.Add("III.15", cmb315.Text);
            student.ListOfTasks.Add("III.16", cmb316.Text);
            student.ListOfTasks.Add("III.17", cmb317.Text);
            student.ListOfTasks.Add("III.18", cmb318.Text);
            student.ListOfTasks.Add("III.19", cmb319.Text);
            student.ListOfTasks.Add("III.20", cmb320.Text);
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
           for (int ix = 0; ix < checkedListBox1.Items.Count; ++ix)
           if (!e.Equals(ix))
                    checkedListBox1.SetItemChecked(ix, false);
        }

        private void lbProfile_SelectedIndexChanged(object sender, EventArgs e)
        {
            for (int ix = 0; ix < lbProfile.Items.Count; ++ix)
                if (!e.Equals(ix))
                    lbProfile.SetItemChecked(ix, false);
        }

        private void lbSex_SelectedIndexChanged(object sender, EventArgs e)
        {
            for (int ix = 0; ix < lbSex.Items.Count; ++ix)
                if (!e.Equals(ix))
                    lbSex.SetItemChecked(ix, false);
        }

        private void lbAccomodation_SelectedIndexChanged(object sender, EventArgs e)
        {
            for (int ix = 0; ix < lbAccomodation.Items.Count; ++ix)
                if (!e.Equals(ix))
                    lbAccomodation.SetItemChecked(ix, false);
        }

        private void lbFather_SelectedIndexChanged(object sender, EventArgs e)
        {
            for (int ix = 0; ix < lbFather.Items.Count; ++ix)
                if (!e.Equals(ix))
                    lbFather.SetItemChecked(ix, false);
        }

        private void lbMother_SelectedIndexChanged(object sender, EventArgs e)
        {
            for (int ix = 0; ix < lbMother.Items.Count; ++ix)
                if (!e.Equals(ix))
                    lbMother.SetItemChecked(ix, false);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(indIII>0)
            {
                listPanelIII[--indIII].BringToFront();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if(indIII < listPanelIII.Count - 1)
            {
                listPanelIII[++indIII].BringToFront();
            }

        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            if (this.ExcelFile == string.Empty)
            {
                txtFile.Text = "Wpierw wybierz plik excel";
                this.ExcelFile = DataTransport.DTinit();
                if (this.ExcelFile == string.Empty)
                    txtFile.Text = "Wpierw wybierz plik excel";
                else
                    txtFile.Text = this.ExcelFile;
            }
            else
                txtFile.Text = DataTransport.DTExport(this.ExcelFile, dataGridView1);
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            if (this.ExcelFile == string.Empty)
            {
                txtFile.Text = "Wpierw wybierz plik excel";
                this.ExcelFile = DataTransport.DTinit();
                if (this.ExcelFile == string.Empty)
                    txtFile.Text = "Wpierw wybierz plik excel";
                else
                    txtFile.Text = this.ExcelFile;
            }    
            else
                txtFile.Text = DataTransport.DTAdd(this.ExcelFile, dataGridView1);
        }

        private void btnDataBase_Click(object sender, EventArgs e)
        {
            this.ExcelFile = DataTransport.DTinit();
            if (this.ExcelFile == string.Empty)
                txtFile.Text = "Wpierw wybierz plik excel";
            else
                txtFile.Text = this.ExcelFile;
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            txtFile.Text = "Klikanie w obrazek nie uzupełni tabeli, do roboty";
        }

        private string ReturnDirectory()
        {
            var maindirectory = Directory.GetCurrentDirectory();
            var regex = new Regex("Fathers");
            Match match = regex.Match(maindirectory);
            if (match.Success)
            {
                var indexer = match.Index;
                while (maindirectory[indexer++].Equals('\\') == false && indexer < maindirectory.Length) { };
                return maindirectory.Remove(indexer - 1);
            }
            else
            {
                var rootdirectory = Directory.GetDirectories(@"C:\");
                var enumerator = rootdirectory.GetEnumerator();
                while (enumerator.MoveNext())
                {
                    var cur = enumerator.Current.ToString();
                    match = regex.Match(cur);
                    if (match.Success)
                    {
                        var index = match.Index;
                        while (cur.ToString()[index++].Equals('\\') == false && index < cur.Length) { };
                        return cur.Remove(index - 1);
                    }
                }

                // if while didn't find anything user should help to avoid problems.
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.InitialDirectory = Directory.GetCurrentDirectory();
                openFileDialog.Multiselect = false;
                openFileDialog.Title = "Prosze wybrać folder główny programu";
                DialogResult dialogResult = new DialogResult();
                dialogResult = openFileDialog.ShowDialog();
                switch (dialogResult)
                {
                    case DialogResult.OK:
                        return openFileDialog.FileName;
                    default:
                        break;
                }
                return "Nie znaleziono głównego folderu programu";
            }
        }
    }
}
