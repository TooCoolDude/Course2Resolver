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
    public partial class ASWSelectorForm : Form
    {
        List<WireAS> _aswWires;

        public WireAS SelectedASW { get; private set; }

        public ASWSelectorForm(double current)
        {
            InitializeComponent();

            label1.Text = current.ToString();

            _aswWires = WiresASReader.GetWiresAS();

            comboBox1.Items.AddRange(_aswWires.Select(o => o.TypeAS_Current.ToString()).ToArray());
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SelectedASW = _aswWires.Where(o => o.TypeAS_Current.ToString() == comboBox1.Text).First();
            DialogResult = DialogResult.OK;
        }
    }
}
