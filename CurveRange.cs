using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SoundCheckTCPClient
{
    public partial class CurveRange : Form
    {
        public  double upper;
        public  double lower;
        public  int judgeType;
        public  double[] XData;
        public  double[] YData;


        public CurveRange()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                lower = double.Parse(textLower.Text);
            }
            catch
            {
                MessageBox.Show("下限值:请输入数值");
                return;
            }

            try
            {
                upper = double.Parse(textUpper.Text);
            }
            catch
            {
                MessageBox.Show("上限值:请输入数值");
                return;
            }

            judgeType = comboBox1.SelectedIndex;

            this.DialogResult = DialogResult.OK;
            this.Close();

            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void CurveRange_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = judgeType;
            textLower.Text = lower.ToString();
            textUpper.Text = upper.ToString();

            chart1.Series[0].Points.DataBindXY(XData, YData);



           
        }
    }
}
