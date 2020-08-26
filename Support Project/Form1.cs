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

        
        List<Vendordata> vendordata = new List<Vendordata>();
        string selectedVendor = string.Empty;
        string selectedProperty = string.Empty;
        List<string> propertyList = new List<string>();
        string[] splitstr;

        private void button1_Click(object sender, EventArgs e)
        {
            //WorkBook workbook = WorkBook.Load(@"C:\Users\Vibhav\source\repos\Support Project\Support Project\Data.xlsx");


            try
            {
                WorkBook workbook = WorkBook.Load(textBox1.Text);
                WorkSheet sheet = workbook.WorkSheets.First();



                //--------------------------------------------------------------------------
                //  Working Data Table
                //

                System.Data.DataTable dataTable = sheet.ToDataTable(true);




                //--------------------------------------------------------------------------


                //                  Original

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
                        a.Email = row[6].ToString(); ;
                        a.AOVendor = row[7].ToString();
                        a.ITContact = row[8].ToString();
                        a.ITContactNUmber = row[5].ToString();
                    }
                    vendordata.Add(a);
                }


                //dataGridView1.DataSource = vendordata;

                //foreach (DataRow row in dataTable.Rows)
                //{
                //    Vendordata a = new Vendordata();

                //    for (int i = 0; i <= 1; i++)
                //    {

                //        a.cluster = row[1].ToString();
                //        a.VendorNo = row[2].ToString();
                //        a.VendorName = row[3].ToString();
                //        a.ContactPerson = row[4].ToString();
                //        a.ContactNumber = row[5].ToString();
                //        a.AOVendor = row[7].ToString();
                //        a.ITContact = row[8].ToString();
                //        a.ITContactNUmber = row[5].ToString();
                //        if (row[6].ToString() != null && (row[6].ToString().Contains(',') || row[6].ToString().Contains(';')))
                //        {
                //            string str = row[6].ToString();
                //            splitstr = str.Split(new char[] { ',', ';' });
                //            foreach (string s in splitstr)
                //            {
                //                Console.WriteLine(s);
                //                MessageBox.Show(s);
                //                a.Email = s;
                //                MessageBox.Show(a.Email);
                //                vendordata.Add(a);
                //            }
                //        }
                //        else
                //        {
                //            a.Email = row[6].ToString();
                //            vendordata.Add(a);
                //        }

                //    }
                //vendordata.Add(a);
                //    splitstr = null;
                //}

                //          SPILTTING TWO EMAILS

                foreach (Vendordata v in vendordata)
                {
                    if (v.Email.Contains(',') || v.Email.Contains(';'))
                    {
                        string str = v.Email;
                        string[] splitstr = str.Split(new char[] { ',', ';' });
                        
                        foreach (string s in splitstr)
                        {

                            Vendordata a = new Vendordata();
                            a.cluster = v.cluster;
                            a.VendorNo = v.VendorNo;
                            a.VendorName = v.VendorName;
                            a.ContactPerson = v.ContactPerson;
                            a.ContactNumber = v.ContactNumber;
                            //MessageBox.Show(s);
                            a.Email = s;
                            //MessageBox.Show(a.Email);
                            a.AOVendor = v.AOVendor;
                            a.ITContact = v.ITContact;
                            a.ITContactNUmber = v.ITContactNUmber;
                            vendordata.Add(a);
                        }
                        
                    }
                }

                dataGridView1.DataSource = vendordata;

                //----------------------------------------------------------------------------------------------

                //      Searching for a specific vendor by his name and retrieving the details of the vendor


                //Vendordata b= vendordata.Find(vd => vd.VendorName == "OCGC Clothings");
                //Console.WriteLine("Vendor Name is : "+b.ContactPerson);


                //-------------------------------------------------------------------------------------------------

                //          Populating Combobox


                foreach (Vendordata l in vendordata)
                {
                    comboBox1.Items.Add(l.VendorName.ToString());
                }
            }
            catch(Exception exception)
            {
                MessageBox.Show( exception.Message);
            }
            finally
            {

            }


            comboBox2.Items.Add("Name of Vendor");
            comboBox2.Items.Add("Cluster");
            comboBox2.Items.Add("Vendor Number");
            comboBox2.Items.Add("Contact Person");
            comboBox2.Items.Add("Contact Number");
            comboBox2.Items.Add("Address of Vendor"); 
            comboBox2.Items.Add("IT Contact"); 
            comboBox2.Items.Add("IT Contact Number");

            //Console.WriteLine( vendordata[1].VendorName);


            //Vendordata finalItem = new Vendordata();
            //if (selectedVendor != null)
            //{
            //    finalItem = vendordata.Find(vd => vd.VendorName == selectedVendor);
            //    Console.WriteLine("Vendor Name :  {0}  Contact Person :  {1}", finalItem.VendorName, finalItem.ContactPerson);
            //}
            //else
            //{
            //    MessageBox.Show("Please select a vendor !!");
            //}

            //object selected = comboBox1.SelectedItem;

            //Console.WriteLine("Selected Item from combobox : " + selected.ToString());

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













        
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

            object selected = comboBox1.SelectedItem;

            selectedVendor = selected.ToString();

            //----------------------------------------------------------------------------------

            //Binding dynamic vendor Details to a Vendordata object

            Vendordata finalItem = new Vendordata();
            List<Vendordata> dynamicList = new List<Vendordata>();
                if (selectedVendor != null)
                {
                    dynamicList.Clear();
                    finalItem = vendordata.Find(vd => vd.VendorName == selectedVendor);
                    Console.WriteLine("Vendor Name :  {0}  Contact Person :  {1}", finalItem.VendorName, finalItem.ContactPerson);
                    dynamicList.Add(finalItem);
                    dataGridView1.DataSource = dynamicList;
                }
                else
                {
                    MessageBox.Show("Please select a vendor !!");
                }
            
                



            //--------------------------------------------------------------------------------------

            //vendordata.Clear();
            //vendordata.Add(finalItem);
            //dataGridView1.DataSource = vendordata;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    textBox1.Text = dialog.FileName;
                }
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Link;
            else
                e.Effect = DragDropEffects.None;
        }

        private void textBox1_DragDrop(object sender, DragEventArgs e)
        {
            




            string[] files = e.Data.GetData(DataFormats.FileDrop) as string[]; // get all files droppeds  
            if (files != null && files.Any())
                textBox1.Text = files.First(); //select the first one  
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            object selected = comboBox2.SelectedItem;
            selectedProperty = selected.ToString();
            Console.WriteLine("Selected Property : " + selectedProperty);
            if(selectedProperty== "Name of Vendor")
            {
                foreach(Vendordata v in vendordata)
                {
                    propertyList.Add(v.VendorName);
                    
                }
            }
            
            else if (selectedProperty == "Cluster")
            {
                propertyList.Clear();
                foreach (Vendordata v in vendordata)
                {
                    propertyList.Add(v.cluster);
                }
            }
            else if (selectedProperty == "Vendor Number")
            {
                propertyList.Clear();
                foreach (Vendordata v in vendordata)
                {
                    propertyList.Add(v.VendorNo);
                }
            }
            else if(selectedProperty == "Contact Person")
            {
                propertyList.Clear();
                foreach (Vendordata v in vendordata)
                {
                    propertyList.Add(v.ContactPerson);
                }
            }
            else if (selectedProperty == "Contact Number")
            {
                propertyList.Clear();
                foreach (Vendordata v in vendordata)
                {
                    propertyList.Add(v.ContactNumber);
                }
            }
            else if (selectedProperty == "Address of Vendor")
            {
                propertyList.Clear();
                foreach (Vendordata v in vendordata)
                {
                    propertyList.Add(v.AOVendor);
                }
            }
            else if (selectedProperty == "IT Contact")
            {
                propertyList.Clear();
                foreach (Vendordata v in vendordata)
                {
                    propertyList.Add(v.ITContact);
                }
            }
            else
            {
                propertyList.Clear();
                foreach (Vendordata v in vendordata)
                {
                    propertyList.Add(v.ITContactNUmber);
                }
            }
           
            foreach(string s in propertyList)
            {
                Console.WriteLine("Property list : " + s);
            }
        }
    }
}
