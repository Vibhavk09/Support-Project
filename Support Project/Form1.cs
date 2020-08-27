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
using System.IO;
using ExcelDataReader;
using System.Reflection;
using System.Runtime.InteropServices;

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
        List<Vendordata> vendordata2 = new List<Vendordata>();
        string selectedVendor = string.Empty;
        string selectedProperty = string.Empty;
        List<string> propertyList = new List<string>();
        //string[] splitstr;
        //string sheetName;
        IExcelDataReader reader;


        //________________________________________DISPLAY BUTTON CONTENT REMOVED______________________________________
        private void button1_Click(object sender, EventArgs e)
        {
            //WorkBook workbook = WorkBook.Load(@"C:\Users\Vibhav\source\repos\Support Project\Support Project\Data.xlsx");


            //try
            //{
            //    WorkBook workbook = WorkBook.Load(textBox1.Text);
            //    WorkSheet sheet = workbook.WorkSheets.First();


            //    FileStream stream = File.Open(textBox1.Text, FileMode.Open, FileAccess.Read);

            //    IExcelDataReader reader;

            //    reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream);

            //    var conf = new ExcelDataSetConfiguration
            //    {
            //        ConfigureDataTable = _ => new ExcelDataTableConfiguration
            //        {
            //            UseHeaderRow = true
            //        }
            //    };

            //    var dataSet = reader.AsDataSet(conf);
            //    var dataTable = dataSet.Tables[0];
            //    // Now you can get data from each sheet by its index or its "name"

            //    //for (int i = 0; i < dataSet.Tables.Count; i++)
            //    //{
            //    //    comboBox3.Items.Add(dataSet.Tables[i].ToString());
            //    //}

            //    //sheetName = comboBox3.SelectedItem.ToString();
            //    //dataTable = dataSet.Tables[sheetName];




            //    //--------------------------------------------------------------------------
            //    //  Working Data Table
            //    //


            //    //--------------------------------------------------------------------------


            //    //                  Original

            //    foreach (DataRow row in dataTable.Rows)
            //    {
            //        Vendordata a = new Vendordata();
            //        for (int i = 0; i < 1; i++)
            //        {
            //            a.cluster = row["Cluster"].ToString();
            //            a.VendorNo = row["Vendor No"].ToString();
            //            a.VendorName = row["Name of Vendor"].ToString();
            //            a.ContactPerson = row["Contact Person"].ToString();
            //            a.ContactNumber = row["Contact Nos"].ToString();
            //            a.Email = row["Email"].ToString(); ;
            //            a.AOVendor = row["Address of Vendor"].ToString();
            //            if (dataTable.Columns.Contains("IT Contact") || dataTable.Columns.Contains("IT Contact ") ||
            //                dataTable.Columns.Contains("IT Contact Number") || dataTable.Columns.Contains("IT Contact Number "))
            //            {
            //                a.ITContact = row["IT Contact "].ToString();
            //                a.ITContactNUmber = row["IT Contact Number"].ToString();
            //            }
            //            else
            //            {
            //                a.ITContact = null;
            //                a.ITContactNUmber = null;
            //            }

            //        }
            //        vendordata.Add(a);
            //    }
            //    //          SPILTTING TWO EMAILS

            //    foreach (Vendordata v in vendordata)
            //    {
            //        if (v.Email.Contains(',') || v.Email.Contains(';') || v.Email.Contains('/'))
            //        {
            //            string str = v.Email;
            //            string[] splitstr = str.Split(new char[] { ',', ';', '/' });

            //            foreach (string s in splitstr)
            //            {

            //                Vendordata a = new Vendordata();
            //                a.cluster = v.cluster;
            //                a.VendorNo = v.VendorNo;
            //                a.VendorName = v.VendorName;
            //                a.ContactPerson = v.ContactPerson;
            //                a.ContactNumber = v.ContactNumber;
            //                a.Email = s;
            //                a.AOVendor = v.AOVendor;
            //                a.ITContact = v.ITContact;
            //                a.ITContactNUmber = v.ITContactNUmber;
            //                vendordata2.Add(a);
            //            }

            //        }
            //        else
            //        {
            //            Vendordata a = new Vendordata();
            //            a.cluster = v.cluster;
            //            a.VendorNo = v.VendorNo;
            //            a.VendorName = v.VendorName;
            //            a.ContactPerson = v.ContactPerson;
            //            a.ContactNumber = v.ContactNumber;
            //            a.Email = v.Email;
            //            a.AOVendor = v.AOVendor;
            //            a.ITContact = v.ITContact;
            //            a.ITContactNUmber = v.ITContactNUmber;
            //            vendordata2.Add(a);
            //        }
            //    }

                //dataGridView1.DataSource = vendordata2;



            //    //-------------------------------------------------------------------------------------------------

            //    //          Populating Combobox


            //    foreach (Vendordata l in vendordata)
            //    {
            //        comboBox1.Items.Add(l.VendorName.ToString());
            //    }
            //}
            //catch (Exception exception)
            //{
            //    MessageBox.Show(exception.Message);
            //}
            //finally
            //{

            //}


            //comboBox2.Items.Add("Name of Vendor");
            //comboBox2.Items.Add("Cluster");
            //comboBox2.Items.Add("Vendor Number");
            //comboBox2.Items.Add("Contact Person");
            //comboBox2.Items.Add("Contact Number");
            //comboBox2.Items.Add("Address of Vendor");
            //comboBox2.Items.Add("IT Contact");
            //comboBox2.Items.Add("IT Contact Number");



            //comboBox2.SelectedIndex = 0;

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
            int? _ITContactNumber;

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

        //___________________________________ SELECTING SPECIFIC VENDOR FOR DETAILS_______________________________________________

        //                                          VENDOR COMBOBOX

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            object selected = comboBox1.SelectedItem;

            

            selectedVendor = selected.ToString();

            //----------------------------------------------------------------------------------

            //Binding dynamic vendor Details to a Vendordata object

            Vendordata finalItem = new Vendordata();
            List<Vendordata> dynamicList = new List<Vendordata>();
            dynamicList.Clear();
            if (selectedVendor != null)
                {
                    
                    finalItem = vendordata2.Find(vd => vd.VendorName == selectedVendor);
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

        //_______________________________________________  FILE SELECTING CODE _______________________________________________

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

        //_________________________________________________    SELECTING SPECIFIC COLUMN DETAILS IN THE TABLE_____________________________

        //                                                              COLUMN DROPDOWN COMBO BOX

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            object selected = comboBox2.SelectedItem;
            selectedProperty = selected.ToString();
            Console.WriteLine("Selected Property : " + selectedProperty);
            if(selectedProperty== "Name of Vendor")
            {
                foreach(Vendordata v in vendordata2)
                {
                    propertyList.Add(v.VendorName);
                    
                }
            }
            
            else if (selectedProperty == "Cluster")
            {
                propertyList.Clear();
                foreach (Vendordata v in vendordata2)
                {
                    propertyList.Add(v.cluster);
                }
            }
            else if (selectedProperty == "Vendor Number")
            {
                propertyList.Clear();
                foreach (Vendordata v in vendordata2)
                {
                    propertyList.Add(v.VendorNo);
                }
            }
            else if(selectedProperty == "Contact Person")
            {
                propertyList.Clear();
                foreach (Vendordata v in vendordata2)
                {
                    propertyList.Add(v.ContactPerson);
                }
            }
            else if (selectedProperty == "Contact Number")
            {
                propertyList.Clear();
                foreach (Vendordata v in vendordata2)
                {
                    propertyList.Add(v.ContactNumber);
                }
            }
            else if (selectedProperty == "Address of Vendor")
            {
                propertyList.Clear();
                foreach (Vendordata v in vendordata2)
                {
                    propertyList.Add(v.AOVendor);
                }
            }
            else if (selectedProperty == "IT Contact")
            {
                propertyList.Clear();
                foreach (Vendordata v in vendordata2)
                {
                    propertyList.Add(v.ITContact);
                }
            }
            else
            {
                propertyList.Clear();
                foreach (Vendordata v in vendordata2)
                {
                    propertyList.Add(v.ITContactNUmber);
                }
            }
           
            foreach(string s in propertyList)
            {
                Console.WriteLine("Property list : " + s);
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        //_____________________________________________ PROCESSING FILE BUTTON CODE _______________________________________________     


        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                WorkBook workbook = WorkBook.Load(textBox1.Text);
                WorkSheet sheet = workbook.WorkSheets.First();
                FileStream stream = File.Open(textBox1.Text, FileMode.Open, FileAccess.Read);
                reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream);
                var conf = new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = true
                    }
                };

                var dataSet = reader.AsDataSet(conf);
                for (int i = 0; i < dataSet.Tables.Count; i++)
                {
                    comboBox3.Items.Add(dataSet.Tables[i].ToString());

                    
                }
                comboBox3.SelectedIndex = 0;
                
                var dataTable= dataSet.Tables[comboBox3.SelectedIndex];
                //                  Original

                foreach (DataRow row in dataTable.Rows)
                {
                    Vendordata a = new Vendordata();
                    for (int i = 0; i < 1; i++)
                    {
                        a.cluster = row["Cluster"].ToString();
                        a.VendorNo = row["Vendor No"].ToString();
                        a.VendorName = row["Name of Vendor"].ToString();
                        a.ContactPerson = row["Contact Person"].ToString();
                        a.ContactNumber = row["Contact Nos"].ToString();
                        a.Email = row["Email"].ToString(); ;
                        a.AOVendor = row["Address of Vendor"].ToString();
                        if (dataTable.Columns.Contains("IT Contact") || dataTable.Columns.Contains("IT Contact ") ||
                            dataTable.Columns.Contains("IT Contact Number") || dataTable.Columns.Contains("IT Contact Number "))
                        {
                            a.ITContact = row["IT Contact "].ToString();
                            a.ITContactNUmber = row["IT Contact Number"].ToString();
                        }
                        else
                        {
                            a.ITContact = null;
                            a.ITContactNUmber = null;
                        }

                    }
                    vendordata.Add(a);
                }
                //          SPILTTING TWO EMAILS

                foreach (Vendordata v in vendordata)
                {
                    if (v.Email.Contains(',') || v.Email.Contains(';') || v.Email.Contains('/'))
                    {
                        string str = v.Email;
                        string[] splitstr = str.Split(new char[] { ',', ';', '/' });

                        foreach (string s in splitstr)
                        {

                            Vendordata a = new Vendordata();
                            a.cluster = v.cluster;
                            a.VendorNo = v.VendorNo;
                            a.VendorName = v.VendorName;
                            a.ContactPerson = v.ContactPerson;
                            a.ContactNumber = v.ContactNumber;
                            a.Email = s;
                            a.AOVendor = v.AOVendor;
                            a.ITContact = v.ITContact;
                            a.ITContactNUmber = v.ITContactNUmber;
                            vendordata2.Add(a);
                        }

                    }
                    else
                    {
                        Vendordata a = new Vendordata();
                        a.cluster = v.cluster;
                        a.VendorNo = v.VendorNo;
                        a.VendorName = v.VendorName;
                        a.ContactPerson = v.ContactPerson;
                        a.ContactNumber = v.ContactNumber;
                        a.Email = v.Email;
                        a.AOVendor = v.AOVendor;
                        a.ITContact = v.ITContact;
                        a.ITContactNUmber = v.ITContactNUmber;
                        vendordata2.Add(a);
                    }
                }

                


                //-------------------------------------------------------------------------------------------------

                //          Populating Combobox


                foreach (Vendordata l in vendordata)
                {
                    comboBox1.Items.Add(l.VendorName.ToString());
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
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



            comboBox2.SelectedIndex = 0;

        }


        //_____________________________   DISPLAYING CONTENT BASED ON DATA SELECTED FROM COMBOBOX CONTAINING SHEETS________________________________


        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                dataGridView1.DataSource = null;
                vendordata.Clear();
                vendordata2.Clear();
                var conf = new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = true
                    }
                };

                var dataSet = reader.AsDataSet(conf);

                object selected = comboBox3.SelectedItem;
                selectedProperty = selected.ToString();

                var dataTable = dataSet.Tables[selectedProperty];

                foreach (DataRow row in dataTable.Rows)
                {
                    Vendordata a = new Vendordata();
                    for (int i = 0; i < 1; i++)
                    {
                        a.cluster = row["Cluster"].ToString();
                        a.VendorNo = row["Vendor No"].ToString();
                        a.VendorName = row["Name of Vendor"].ToString();
                        a.ContactPerson = row["Contact Person"].ToString();
                        a.ContactNumber = row["Contact Nos"].ToString();
                        a.Email = row["Email"].ToString(); ;
                        a.AOVendor = row["Address of Vendor"].ToString();
                        if (dataTable.Columns.Contains("IT Contact") || dataTable.Columns.Contains("IT Contact ") ||
                            dataTable.Columns.Contains("IT Contact Number") || dataTable.Columns.Contains("IT Contact Number "))
                        {
                            a.ITContact = row["IT Contact "].ToString();
                            a.ITContactNUmber = row["IT Contact Number"].ToString();
                        }
                        else
                        {
                            a.ITContact = null;
                            a.ITContactNUmber = null;
                        }

                    }
                    vendordata.Add(a);
                }
                //          SPILTTING TWO EMAILS

                foreach (Vendordata v in vendordata)
                {
                    if (v.Email.Contains(',') || v.Email.Contains(';') || v.Email.Contains('/'))
                    {
                        string str = v.Email;
                        string[] splitstr = str.Split(new char[] { ',', ';', '/' });

                        foreach (string s in splitstr)
                        {

                            Vendordata a = new Vendordata();
                            a.cluster = v.cluster;
                            a.VendorNo = v.VendorNo;
                            a.VendorName = v.VendorName;
                            a.ContactPerson = v.ContactPerson;
                            a.ContactNumber = v.ContactNumber;
                            a.Email = s;
                            a.AOVendor = v.AOVendor;
                            a.ITContact = v.ITContact;
                            a.ITContactNUmber = v.ITContactNUmber;
                            vendordata2.Add(a);
                        }

                    }
                    else
                    {
                        Vendordata a = new Vendordata();
                        a.cluster = v.cluster;
                        a.VendorNo = v.VendorNo;
                        a.VendorName = v.VendorName;
                        a.ContactPerson = v.ContactPerson;
                        a.ContactNumber = v.ContactNumber;
                        a.Email = v.Email;
                        a.AOVendor = v.AOVendor;
                        a.ITContact = v.ITContact;
                        a.ITContactNUmber = v.ITContactNUmber;
                        vendordata2.Add(a);
                    }
                }

                dataGridView1.DataSource = vendordata2;
            }
            catch(Exception exception)
            {
                MessageBox.Show( exception.Message);
            }
            comboBox1.Items.Clear();
            foreach (Vendordata l in vendordata)
            {
                comboBox1.Items.Add(l.VendorName.ToString());
            }
            
        }
    }
}
