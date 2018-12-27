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
using Driver.File.Ini;

namespace SoundCheckTCPClient
{
    public partial class Project : Form
    {
        IniFileClass iniFile = new IniFileClass();

        public Project()
        {
            InitializeComponent();
        }

        string projectName = "";
        string sqcPath = "";
        private void button1_Click(object sender, EventArgs e)
        {
            string soundCheckPath = "";
            if (label1.Text != "")
            {
                string fName = Directory.GetCurrentDirectory().ToString() + "\\Ref.ini";
                FileStream fs = new FileStream(fName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                fs.Close();
                try
                {
                    iniFile.IniReadValue(fName, "基准", "SouncCheckPath");

                    soundCheckPath = Path.GetDirectoryName(iniFile.IniReadValue(fName, "基准", "SouncCheckPath"))+"\\"+ iniFile.IniReadValue(fName, "基准", "SoundCheckConfigFileName"); ;

                    fName = Directory.GetCurrentDirectory().ToString() + "\\测试产品\\" + projectName + "\\" + label1.Text;

                    //fs = new FileStream("D:\\CSoundCheck 15.0-x86\\SoundCheck 15.0.ini", FileMode.OpenOrCreate, FileAccess.ReadWrite);
                    //fs.Close();

                    iniFile.IniWriteValue(soundCheckPath, "Files","RecentFiles.list", fName);
                    iniFile.IniWriteValue(soundCheckPath, "Paths", "SEQ NAME", label1.Text.Substring(0, label1.Text.IndexOf(".")));
                    iniFile.IniWriteValue(soundCheckPath, "Paths", "SEQ FOLDER PATH", Directory.GetCurrentDirectory().ToString() + "\\测试产品\\" + projectName);

                }
                catch { }
                this.Hide();
                Form1 form = new Form1(projectName, fName);
                //form.projectName = label1.Text;
                try
                {
                    form.ShowDialog();
                }
                catch
                {

                }
                this.Close();
            }
        }

        private void Project_Load(object sender, EventArgs e)
        {
            string fName = Directory.GetCurrentDirectory().ToString() + "\\测试产品";
            string[] subfolders = Directory.GetDirectories(fName);
            listBox1.Items.Clear();
            foreach (string s in subfolders)
            {

                listBox1.Items.Add(s.Substring(s.LastIndexOf(@"\") + 1));

               // subItem.DropDownItems.Add(tsmi2);
            }
        }

        private void listBox1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            string fName = Directory.GetCurrentDirectory().ToString() + "\\测试产品\\";
            listBox2.Items.Clear();
            if ((listBox1.SelectedIndex >= 0) && (listBox1.SelectedIndex >= 0))
            {
                DirectoryInfo folder = new DirectoryInfo(fName+listBox1.SelectedItem.ToString());
                foreach (FileInfo file in folder.GetFiles("*.sqc"))
                {
                        listBox2.Items.Add( file.Name);            
                }

                projectName= listBox1.Text;
            }
        }

        private void listBox2_MouseDoubleClick(object sender, MouseEventArgs e)
        {

            if ((listBox2.SelectedIndex >= 0) && (listBox2.SelectedIndex <listBox2.Items.Count))
            {
                label1.Text = listBox2.Text;
            }
        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
