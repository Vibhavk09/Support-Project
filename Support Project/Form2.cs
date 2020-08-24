using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using IronXL;
using System.Linq;

namespace Support_Project
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            WorkBook workbook = WorkBook.Load(@"C:\Users\Vibhav\source\repos\Support Project\Support Project\Data.xlsx");
            WorkSheet sheet = workbook.WorkSheets.First();
            
            //dataGridView1.ColumnCount = 3;
            //dataGridView1.Columns[0].Name = "Sr No.";
            //dataGridView1.Columns[1].Name = "Name of Vendor";
            //dataGridView1.Columns[2].Name = "Email";

            DataTable DTable = new DataTable();
            BindingSource SBind = new BindingSource();
            //ServersTable - DataGridView
            for (int i = 0; i < dataGridView1.ColumnCount; ++i)
            {
                DTable.Columns.Add(new DataColumn(dataGridView1.Columns[i].Name));
            }

            for (int i = 0; i < Apps.Count; ++i)
            {
                DataRow r = DTable.NewRow();
                r.BeginEdit();
                foreach (DataColumn c in DTable.Columns)
                {
                    r[c.ColumnName] = //writing values
            }
                r.EndEdit();
                DTable.Rows.Add(r);
            }
            SBind.DataSource = DTable;
            ServersTable.DataSource = SBind;
        }
    }
}
