using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Course2
{
    public partial class SIPSelectorForm : Form
    {
        List<WireSIP> _sipWires;

        public WireSIP SelectedSIP { get; private set; }

        public SIPSelectorForm(double current)
        {
            InitializeComponent();

            label1.Text = current.ToString();

            _sipWires = WiresSIPReader.GetWiresSIP();

            comboBox1.Items.AddRange(_sipWires.Select(o => o.Current25.ToString()).ToArray());
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SelectedSIP = _sipWires.Where(o => o.Current25.ToString() == comboBox1.Text).First();
            DialogResult = DialogResult.OK;
        }
    }
}
