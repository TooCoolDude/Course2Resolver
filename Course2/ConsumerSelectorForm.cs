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
    public partial class ConsumerSelectorForm : Form
    {
        List<ConsumerObject> _consumerObjects;

        public ConsumerObject SelectedConsumerObject { get; private set; }

        public ConsumerSelectorForm()
        {
            InitializeComponent();

            _consumerObjects = ConsumerObjectsReader.GetConsumerObjects();

            comboBox1.Items.AddRange(_consumerObjects.Select(o=>o.ObjName).ToArray());
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SelectedConsumerObject = _consumerObjects.Where(o=>o.ObjName == comboBox1.Text).First();
        }
    }
}
