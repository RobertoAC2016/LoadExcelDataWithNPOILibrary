namespace LoadExcelDataWithNPOILibrary
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            XLSLoadData xls = new XLSLoadData();
            //ahora revisar como funciona el metodo
            var tbl = xls.getExcelData();
            if (tbl != null)
            {
                dgvDatos.DataSource = tbl;
            }
        }
    }
}