namespace Course2
{
    public partial class MainForm : Form
    {
        List<Variant> _variants;

        List<ConsumerObject> _consumerObjects;

        List<WireSIP> _wiresSIP;
        
        public MainForm()
        {
            InitializeComponent();

            _variants = VariantsReader.GetVariants();

            _consumerObjects = ConsumerObjectsReader.GetConsumerObjects();

            _wiresSIP = WiresSIPReader.GetWiresSIP();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (int.TryParse(textBox1.Text, out int variantNum) && Enumerable.Range(1,30).Contains(variantNum))
            {
                var variant = _variants.Where(v => v.Num == variantNum).First();
            }

            else
                textBox1.Text = "incorrect input";
        }
    }
}
