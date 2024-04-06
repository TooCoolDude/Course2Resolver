namespace Course2
{
    public partial class MainForm : Form
    {
        List<Variant> _variants;

        List<ConsumerObject> _consumerObjects;

        List<WireSIP> _wiresSIP;

        List<WireAS> _wiresAS;

        List<Transformer> _transformers;
        
        public MainForm()
        {
            InitializeComponent();

            _variants = VariantsReader.GetVariants();

            _consumerObjects = ConsumerObjectsReader.GetConsumerObjects();

            _wiresSIP = WiresSIPReader.GetWiresSIP();

            _wiresAS = WiresASReader.GetWiresAS();

            _transformers = TransformersReader.GetTransformers();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (int.TryParse(textBox1.Text, out int variantNum) && Enumerable.Range(1,30).Contains(variantNum))
            {
                var variant = _variants.Where(v => v.Num == variantNum).First();

                File.Copy("sources\\Template.docx", Directory.GetCurrentDirectory() + "\\Result.docx", true);

                var replacements = Calculator.GetVariablesAndValues(variant);
                DocumentInteractor.WriteChanges(Directory.GetCurrentDirectory() + "\\Result.docx", replacements);
            }

            else
                textBox1.Text = "incorrect input";
        }
    }
}
