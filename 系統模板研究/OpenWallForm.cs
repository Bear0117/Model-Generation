using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace Modeling
{
    public partial class OpenWallForm : System.Windows.Forms.Form
    {
        Document Doc;
        public string levelName;
        public double KerbHeight;
        public double gridsize;

        public OpenWallForm(Document doc)
        {
            InitializeComponent();

            Doc = doc;

            FilteredElementCollector filter_level = new FilteredElementCollector(Doc).OfClass(typeof(Level));
            //MessageBox.Show(filter_level.Count().ToString());
            IList<string> list_level = new List<string>();
            foreach (Level level in filter_level)
            {
                list_level.Add(level.Name);
            }
            cb_level.DataSource = list_level;
        }

        public string LevelName
        {
            get { return this.cb_level.SelectedText; }
        }
        

        private void TextBox1_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void Text_Level_TextChanged(object sender, EventArgs e)
        {

        }

        private void btn_OK_Click(object sender, EventArgs e)
        {
            levelName = this.cb_level.SelectedValue.ToString();
            gridsize = double.Parse(this.GridSizeNumber.Text);
            KerbHeight = Int32.Parse(this.Text_kerb.Text);
            //MessageBox.Show(levelName);
            Close();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void btn_Cancel_Click(object sender, EventArgs e)
        {
            this.DialogResult= DialogResult.Cancel;
            Close();
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void GridSizeNumber_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
