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
    public partial class TransformerSelectorForm : Form
    {
        List<Transformer> _transformers;

        public Transformer SelectedTransformer { get; private set; }

        public TransformerSelectorForm(double power)
        {
            InitializeComponent();

            label1.Text = power.ToString();

            _transformers = TransformersReader.GetTransformers();

            comboBox1.Items.AddRange(_transformers.Select(o => $"{o.Power}").ToArray());
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SelectedTransformer = _transformers.Where(o => o.Power == double.Parse(comboBox1.Text)).First();
            DialogResult = DialogResult.OK;
        }
    }
}
