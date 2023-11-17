using System;
using Autodesk.Revit.DB;

namespace Modeling
{
    public partial class LevelHeightForm : System.Windows.Forms.Form
    {
        public double levelHeight;

        public LevelHeightForm(Document doc)
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            levelHeight = Int32.Parse(this.text_Height.Text);
            //MessageBox.Show(levelName);
            Close();
        }

        private void text_Height_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
