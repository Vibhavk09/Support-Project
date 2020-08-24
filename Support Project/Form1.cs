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
    public partial class Form1 : Form
    {
        
        public Form1()
        {
            InitializeComponent();
            
        }


        private void Form1_Load(object sender, EventArgs e)
        {

        }

        

        private void button1_Click(object sender, EventArgs e)
        {
            WorkBook workbook = WorkBook.Load(@"C:\Users\Vibhav\source\repos\Support Project\Support Project\Data.xlsx");
            WorkSheet sheet = workbook.WorkSheets.First();
            //string str = string.Empty;
            //foreach (var cell in sheet["D1:D20"])
            //{
            //    richTextBox1.Text += cell.Text + Environment.NewLine;
            //}


            //--------------------------------------------------------------------------
                    //  Working Data Table
            //

            System.Data.DataTable dataTable = sheet.ToDataTable(true);

            //foreach (DataRow row in dataTable.Rows)
            //{
            //    for (int i = 0; i < dataTable.Columns.Count; i++)
            //    {
            //        richTextBox1.Text += (row[i]) + Environment.NewLine;
            //        //Console.WriteLine((row[i]) + Environment.NewLine);
            //    }
            //}


            //--------------------------------------------------------------------------

            List<Vendordata> vendordata = new List<Vendordata>();

            
            foreach (DataRow row in dataTable.Rows)
            {
                Vendordata a = new Vendordata();
                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    
                    a.cluster = row[1].ToString();
                    a.VendorNo = row[2].ToString();
                    a.VendorName = row[3].ToString();
                    a.ContactPerson = row[4].ToString();
                    a.ContactNumber = row[5].ToString();
                    a.Email = row[6].ToString();;
                    a.AOVendor = row[7].ToString();
                    a.ITContact = row[8].ToString();
                    a.ITContactNUmber = row[5].ToString();
                }
                vendordata.Add(a);
            }

            dataGridView1.DataSource = vendordata;


            //Console.WriteLine("Total Items in list : "+ vendordata.Count());
            //foreach(Vendordata c in vendordata)
            //{
            //    Console.WriteLine("Vendor Name : {0}  COntact Number : {1}", c.VendorName, c.ContactNumber);
            //}
        }



        public class Vendordata
        {
            string _cluster;
            string _vendorNo;
            string _vendorName;
            string _contactPerson;
            string _contactNumber;
            string _email;
            string _addressOfVendor;
            string _ITContact;
            string _ITContactNumber;

            public string cluster
            {
                get;
                set;
            }
            public string VendorNo
            {
                get;
                set;
            }
            public string VendorName
            {
                get;
                set;
            }
            public string ContactPerson
            {
                get;
                set;
            }
            public string ContactNumber
            {
                get;
                set;
            }
            public string Email
            {
                get;
                set;
            }
            public string AOVendor
            {
                get;
                set;
            }
            public string ITContact
            {
                get;
                set;

            }
            public string ITContactNUmber
            {
                get;
                set;

            }

        }

        private void richTextBox1_TextChanged_1(object sender, EventArgs e)
        {

        }
    }
}
