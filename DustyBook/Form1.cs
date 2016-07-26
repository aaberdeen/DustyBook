using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Text.RegularExpressions;
using MySql.Data.MySqlClient;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace xml_treeview
{
    public partial class Form1 : Form
    {
        XElement doc;
        IEnumerable<System.Xml.Linq.XElement> result1;
        IEnumerable<System.Xml.Linq.XElement> result2;
        IEnumerable<System.Xml.Linq.XElement> result3;
        IEnumerable<System.Xml.Linq.XElement> result4;
        IEnumerable<System.Xml.Linq.XElement> result5;
        IEnumerable<System.Xml.Linq.XElement> result6;

        Dictionary<string, string> dicA = new Dictionary<string, string>();
        Dictionary<string, string> dicB = new Dictionary<string, string>();
        Dictionary<string, string> dicC = new Dictionary<string, string>();
        Dictionary<string, string> dicD = new Dictionary<string, string>();
        Dictionary<string, string> dicE = new Dictionary<string, string>();

        private MySqlConnection connection;
        MySqlDataAdapter dAdapter;
        DataTable dTable;
        private string server;
        private string database;
        private string uid;
        private string password;
        private string port;
        private string connectionString;
        int treeLevel;
        string User;

        InputBox inputBox;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            Initialize();

            textBox1.Text = @"\\STEWIE\do documents\DustyBookApp\Index.xml";    //Application.StartupPath + "\\Sample.xml";
            // textBox1.SetBounds(64, 8, 256, 20);
            button1.Text = "Populate the TreeView with XML";
            // button1.SetBounds(8, 40, 200, 20);

            string time = DateTime.Now.ToShortTimeString();
            int timeInt = Convert.ToInt16(time.Substring(0, 2));
            User = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            string[] userSplit = User.Split('\\');
            string title;

            if (timeInt < 12)
            {
                title = string.Format("Dusty Book (Beta) - Good Morning {0}", userSplit[1]);
            }
            else
            {
                title = string.Format("Dusty Book (Beta) - Good Afternoon {0}", userSplit[1]);
            }


            this.Text = title;

            inputBox = new InputBox();
            LoadTree();
            //  this.Width = 336;
            //  this.Height = 368;
            // treeView1.SetBounds(8, 72, 312, 264);
        }

        private void Initialize()
        {
            server = "10.1.0.15";
            database = "dd_parts";
            uid = "dev";
            password = null;
            port = "3306";

            connectionString = "SERVER=" + server + ";" + "Port=" + port + ";" + "DATABASE=" + database + ";" + "UID=" + uid + ";" + "PASSWORD=" + password + ";";

            connection = new MySqlConnection(connectionString);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            LoadTree();
        }

        private void LoadTree()
        {
            try
            {

                //  http://stackoverflow.com/questions/12757461/populate-a-combobox-based-on-previous-combobox-selection-in-xml

                doc = XElement.Load(textBox1.Text);

                backgroundWorker1.RunWorkerAsync();

            }
            catch (XmlException xmlEx)
            {
                MessageBox.Show(xmlEx.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }





        //private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    //put this code in the SelectedIndexChanged event of cbProduct
        //    ComboBox a = sender as ComboBox;
        //    string product2Search = a.SelectedItem.ToString();//selected value of cbProduct

        //    string outString;

        //    dicA.TryGetValue(product2Search, out outString);

        //    textBox2.Clear();
        //    textBox2.AppendText(outString);


        //    comboBox2.Items.Clear();
        //    comboBox3.Items.Clear();
        //    comboBox4.Items.Clear();
        //    comboBox5.Items.Clear();

        //    comboBox2.Text = "";
        //    comboBox3.Text = "";
        //    comboBox4.Text = "";
        //    comboBox5.Text = "";

        //   // comboBox2.Items.AddRange(doc.Descendants("items").Where(x => x.Element("a").Value == product2Search).Select(y => y.Element("b").Value).ToArray<string>());

        //    //IEnumerable<XElement> awElements = from el in doc.Descendants("items")
        //    //                                   where el.Element("a").Value == product2Search
        //    //                                    select el;

        //    //foreach (XElement el in awElements)
        //    //    Console.WriteLine(el.Name.ToString());


        //    result2 = from mainitem in doc.Descendants("a")
        //                 where mainitem.Attribute("ID").Value == product2Search
        //                           select mainitem;
        //    dicB.Clear();
        //    foreach (var subitem in result2.First().Descendants("b"))
        //    {
        //        comboBox2.Items.Add(subitem.Attribute("ID").Value);
        //        dicB.Add((subitem.Attribute("ID").Value), (subitem.Attribute("Number").Value));
        //    }


        ////         cbProduct.Items.AddRange(doc.Descendants("items").Select(x => x.Element("a").Value).ToArray<string>());//adds all products
        // //   cbbdoc.Items.AddRange(doc.Descendants("items").Where(x => x.Element("a").Value == product2Search).Select(y => y.Element("b").Value).ToArray<string>());                                                                                                                
        //    //******************************
        //}
        //private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    ComboBox a = sender as ComboBox;
        //    string product2Search = a.SelectedItem.ToString();
        //    string outString;

        //    dicB.TryGetValue(product2Search, out outString);


        //    if (textBox2.Text.Length >= 2)
        //    {
        //       textBox2.Text = textBox2.Text.Remove(1);
        //    }
        //    textBox2.AppendText(outString);            

        //    comboBox3.Items.Clear();
        //    comboBox4.Items.Clear();
        //    comboBox5.Items.Clear();

        //    comboBox3.Text = "";
        //    comboBox4.Text = "";
        //    comboBox5.Text = "";



        //    result3 = from mainitem in result2.Descendants("b")
        //                 where mainitem.Attribute("ID").Value == product2Search
        //                 select mainitem;
        //    dicC.Clear();
        //    foreach (var subitem in result3.First().Descendants("c"))
        //    {
        //        comboBox3.Items.Add(subitem.Attribute("ID").Value);
        //        dicC.Add((subitem.Attribute("ID").Value), (subitem.Attribute("Number").Value));
        //    }

        //}
        //private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    ComboBox a = sender as ComboBox;
        //    string product2Search =   a.SelectedItem.ToString();
        //    string outString;

        //    dicC.TryGetValue(product2Search, out outString);
        //    if (textBox2.Text.Length >= 3)
        //    {
        //        textBox2.Text = textBox2.Text.Remove(2);
        //    }
        //    textBox2.AppendText("_");
        //    textBox2.AppendText(outString);

        //    comboBox4.Items.Clear();
        //    comboBox5.Items.Clear();

        //    comboBox4.Text = "";
        //    comboBox5.Text = "";

        //    result4 = from mainitem in result3.Descendants("c")
        //                 where mainitem.Attribute("ID").Value == product2Search
        //                 select mainitem;
        //    dicD.Clear();
        //    foreach (var subitem in result4.First().Descendants("d"))
        //    {
        //        comboBox4.Items.Add(subitem.Attribute("ID").Value);
        //        dicD.Add((subitem.Attribute("ID").Value), (subitem.Attribute("Number").Value));
        //    }
        //}

        //private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    ComboBox a = sender as ComboBox;
        //    string product2Search = a.SelectedItem.ToString();
        //    string outString;

        //    dicD.TryGetValue(product2Search, out outString);

        //    if (textBox2.Text.Length >= 5)
        //    {
        //        textBox2.Text = textBox2.Text.Remove(4);
        //    }
        //    if (outString != null)
        //    {
        //        textBox2.AppendText(outString);
        //    }

        //    comboBox5.Items.Clear();
        //    comboBox5.Text = "";

        //    result5 = from mainitem in result4.Descendants("d")
        //              where mainitem.Attribute("ID").Value == product2Search
        //              select mainitem;
        //    dicE.Clear();
        //    foreach (var subitem in result5.First().Descendants("e"))
        //    {
        //        comboBox5.Items.Add(subitem.Attribute("ID").Value);
        //        dicE.Add((subitem.Attribute("ID").Value), (subitem.Attribute("Number").Value));
        //    }
        //}

        //private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    ComboBox a = sender as ComboBox;
        //    string product2Search = a.SelectedItem.ToString();
        //    string outString;

        //    dicE.TryGetValue(product2Search, out outString);
        //    if (textBox2.Text.Length >= 6)
        //    {
        //        textBox2.Text = textBox2.Text.Remove(5);
        //    }
        //    textBox2.AppendText(outString);
        //    textBox2.AppendText("_");

        //    result5 = from mainitem in result4.Descendants("e")
        //              where mainitem.Attribute("ID").Value == product2Search
        //              select mainitem;
        //    foreach (var subitem in result5.First().Descendants("f"))
        //    {
        //        comboBox5.Items.Add(subitem.Attribute("ID").Value);

        //    }
        //}

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            // SECTION 1. Create a DOM Document and load the XML data into it.
            XmlDocument dom = new XmlDocument();
            dom.Load(textBox1.Text);

            // SECTION 2. Initialize the TreeView control.
            //Invoke(new MethodInvoker(
            //       delegate {treeView1.Nodes.Clear(); }
            //       ));

            //Invoke(new MethodInvoker(
            //       delegate { treeView1.Nodes.Add(new TreeNode(dom.DocumentElement.Name)); }
            //       ));



            TreeNode tNode = new TreeNode();
            //tNode = treeView1.Nodes[0];
            tNode = new TreeNode(dom.DocumentElement.Name);
            // SECTION 3. Populate the TreeView with the DOM nodes.
            AddNode(dom.DocumentElement, tNode);

            backgroundWorker1.ReportProgress(100, tNode);
            //Invoke(new MethodInvoker(
            //       delegate { treeView1.ExpandAll(); }
            //       ));

        }

        public void SerializeTreeView(TreeView treeView, string fileName)
        {

            XmlWriterSettings xmlWriterSettings = new XmlWriterSettings();
            xmlWriterSettings.NewLineOnAttributes = false;
            xmlWriterSettings.Indent = true;

            XmlWriter textWriter = XmlWriter.Create(fileName, xmlWriterSettings);


            //XmlTextWriter textWriter = new XmlTextWriter(fileName,System.Text.Encoding.ASCII);


            //textWriter.Formatting = Formatting.Indented;
            // writing the xml declaration tag
            textWriter.WriteStartDocument();
            //textWriter.WriteRaw("\r\n");
            // writing the main tag that encloses all node tags
            //textWriter.WriteStartElement("TreeView");

            // save the nodes, recursive method
            SaveNodes(treeView.Nodes, textWriter);

            // textWriter.WriteEndElement();

            textWriter.Close();




        }





        // Xml tag for node, e.g. 'node' in case of <node></node>
        private string XmlNodeTag = "Davis_Derby";

        // Xml attributes for node e.g. <node text="Asia" tag="" 
        // imageindex="1"></node>
        private const string XmlNodeTextAtt = "ID";
        private const string XmlNodeTagAtt = "tag";
        private const string XmlNodeImageIndexAtt = "Number";
        private void SaveNodes(TreeNodeCollection nodesCollection, XmlWriter textWriter)
        {
            for (int i = 0; i < nodesCollection.Count; i++)
            {

                TreeNode node = nodesCollection[i];
                int level = nodesCollection[i].Level;

                switch (level)
                {
                    case 0: { XmlNodeTag = "Davis_Derby"; break; }
                    case 1: { XmlNodeTag = "a"; break; }
                    case 2: { XmlNodeTag = "b"; break; }
                    case 3: { XmlNodeTag = "c"; break; }
                    case 4: { XmlNodeTag = "d"; break; }
                    case 5: { XmlNodeTag = "e"; break; }
                    default:
                        break;
                }

                textWriter.WriteStartElement(XmlNodeTag);



                char[] delimiters = new char[] { '(', ')' };
                string[] number = node.Text.Split(delimiters);
                string num = "";
                if (number.Count() < 2)
                {
                    num = "0";
                }
                else
                {
                    num = number[1];
                }
                textWriter.WriteAttributeString(XmlNodeImageIndexAtt, num);// node.ImageIndex.ToString());

                string[] ID = node.Text.Split('-');
                string id;
                if (ID.Count() < 2)
                {
                    id = "top";
                }
                else
                {
                    id = ID[1];
                }
                textWriter.WriteAttributeString(XmlNodeTextAtt, id.Trim());

                if (node.Tag != null)
                {
                    textWriter.WriteAttributeString(XmlNodeTagAtt,
                                                   node.Tag.ToString());
                    //textWriter.WriteRaw("\r\n");
                }

                //textWriter.WriteRaw("\r\n");
                // add other node properties to serialize here  
                if (node.Nodes.Count > 0)
                {
                    SaveNodes(node.Nodes, textWriter);

                }
                textWriter.WriteEndElement();
                // textWriter.WriteRaw("\r\n");
            }
        }


        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            // treeView1.Nodes[0] = (TreeNode)e.UserState;
            treeView1.Nodes.Clear();
            treeView1.Nodes.Add((TreeNode)e.UserState);
            treeView1.Nodes[0].Expand();
        }

        private void AddNode(XmlNode inXmlNode, TreeNode inTreeNode)
        {
            XmlNode xNode;
            TreeNode tNode;
            XmlNodeList nodeList;
            int i;

            // Loop through the XML nodes until the leaf is reached.
            // Add the nodes to the TreeView during the looping process.
            if (inXmlNode.HasChildNodes)
            {
                nodeList = inXmlNode.ChildNodes;


                for (i = 0; i <= nodeList.Count - 1; i++)
                {
                    xNode = inXmlNode.ChildNodes[i];

                    // inTreeNode.Nodes.Add(new TreeNode(xNode.Name));
                    //   inTreeNode.Nodes.Add(new TreeNode(xNode.SelectSingleNode("ID").Value));
                    inTreeNode.Nodes.Add(new TreeNode(string.Format("({0}) - {1}", xNode.Attributes["Number"].Value, xNode.Attributes["ID"].Value)));


                    // inTreeNode.Text = (inXmlNode.OuterXml).Trim(); 


                    tNode = inTreeNode.Nodes[i];
                    AddNode(xNode, tNode);
                }
            }
            else
            {
                // Here you need to pull the data from the XmlNode based on the
                // type of node, whether attribute values are required, and so forth.


                //  inTreeNode.Text = (inXmlNode.OuterXml).Trim();
                inTreeNode.Text = string.Format("({0}) - {1}", inXmlNode.Attributes["Number"].Value, inXmlNode.Attributes["ID"].Value);

            }
        }



        private List<TreeNode> CurrentNodeMatches = new List<TreeNode>();

        private int LastNodeIndex = 0;

        private string LastSearchText;


        private void button2_Click(object sender, EventArgs e)
        {


            string searchText = this.textBox3.Text;
            if (String.IsNullOrEmpty(searchText))
            {
                return;
            };


            if (LastSearchText != searchText)
            {
                //It's a new Search
                CurrentNodeMatches.Clear();
                LastSearchText = searchText;
                LastNodeIndex = 0;
                SearchNodes(searchText, treeView1.Nodes[0]);
            }

            if (LastNodeIndex >= 0 && CurrentNodeMatches.Count > 0 && LastNodeIndex < CurrentNodeMatches.Count)
            {
                TreeNode selectedNode = CurrentNodeMatches[LastNodeIndex];
                LastNodeIndex++;
                this.treeView1.SelectedNode = selectedNode;
                this.treeView1.SelectedNode.Expand();
                this.treeView1.Select();

            }

            if (LastNodeIndex >= 0 && CurrentNodeMatches.Count > 0 && LastNodeIndex >= CurrentNodeMatches.Count)
            {
                LastNodeIndex = 0;
            }

        }

        private void SearchNodes(string SearchText, TreeNode StartNode)
        {
            TreeNode node = null;
            while (StartNode != null)
            {
                if (StartNode.Text.ToLower().Contains(SearchText.ToLower()))
                {
                    CurrentNodeMatches.Add(StartNode);
                };
                if (StartNode.Nodes.Count != 0)
                {
                    SearchNodes(SearchText, StartNode.Nodes[0]);//Recursive Search 
                };
                StartNode = StartNode.NextNode;
            };

        }

        private void treeView1_AfterSelect(object sender, System.Windows.Forms.TreeViewEventArgs e)
        {
            string temp = e.Node.FullPath;

            char[] delimiters = new char[] { '\\' };
            string[] parts = temp.Split(delimiters,
                             StringSplitOptions.RemoveEmptyEntries);
            int i = 0;
            textBox2.Clear();
            foreach (string part in parts)
            {

                //textBox2.AppendText(part.Split(new char[] { '(', ')' })[1]);

                int start = part.IndexOf("(");
                int end = part.IndexOf(")");


                if (i == 3)
                {
                    textBox2.AppendText("_");
                }

                if (start != -1)
                {

                    textBox2.AppendText(part.Substring(start + 1, end - start - 1));
                }
                i++;
            }
        }

        private void comboBox1_Click(object sender, EventArgs e)
        {
            getTables();
        }
        public void getTables()
        {
            MySqlCommand command = connection.CreateCommand();
            command.CommandText = "SHOW TABLES;";
            MySqlDataReader Reader;
            connection.Open();
            Reader = command.ExecuteReader();

            List<String> Tablenames = new List<String>();
            while (Reader.Read())
            {
                Tablenames.Add(Reader.GetString(0));

            }

            comboBox1.DataSource = Tablenames;
            connection.Close();

        }

        private void button3_Click(object sender, EventArgs e)
        {
            string form = string.Format("{0}", comboBox1.SelectedValue);
            LoadTable(form);
        }

        private void LoadTable(string table)
        {
            //create the database query
            string query = string.Format("SELECT * FROM {0}", table);

            //create an OleDbDataAdapter to execute the query
            dAdapter = new MySqlDataAdapter(query, connectionString);

            //create a command builder
            MySqlCommandBuilder cBuilder = new MySqlCommandBuilder(dAdapter);

            //create a DataTable to hold the query results
            dTable = new DataTable();

            //---------------------------

            try
            {
                //fill the DataTable
                dAdapter.Fill(dTable);

                //the DataGridView
                //DataGridView dgView = new DataGridView();

                //BindingSource to sync DataTable and DataGridView
                BindingSource bSource = new BindingSource();

                //set the BindingSource DataSource
                bSource.DataSource = dTable;

                //set the DataGridView DataSource
                dataGridView1.DataSource = bSource;
                textBox5.Text = table;
                dataGridView1.Columns["CodeA"].Width = 50;
                dataGridView1.Columns["CodeA"].DefaultCellStyle.Format = "00";
                dataGridView1.Columns["CodeB"].Width = 50;
                dataGridView1.Columns["CodeB"].DefaultCellStyle.Format = "000";
                dataGridView1.Columns["CodeNumber"].Width = 75;
                dataGridView1.Columns["CodeNumber"].DefaultCellStyle.Format = "000";

                dataGridView1.Columns["Description"].Width = 240;
                try
                {
                    dataGridView1.Columns["Value"].Width = 50;
                    dataGridView1.Columns["Mount"].Width = 50;
                    dataGridView1.Columns["Package"].Width = 55;
                    dataGridView1.Columns["RoHS"].Width = 40;
                    dataGridView1.Columns["Elements"].Width = 50;
                    dataGridView1.Columns["Tol"].Width = 50;
                    dataGridView1.Columns["Power"].Width = 50;
                    dataGridView1.Columns["Voltage"].Width = 50;
                }
                catch { }
            }
            catch (Exception e)
            {
                // MessageBox.Show(e.Message);
            }

            comboBox2.Items.Clear();
            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                comboBox2.Items.Add(column.HeaderText);
            }
        }



        private void InsertColumn(string name, string table)
        {

            ////ALTER TABLE `dd_parts`.`47_256` ADD COLUMN `User` VARCHAR(45) NULL  AFTER `Description` ;

            ////create the database query
            //string query = string.Format("ALTER TABLE '{0}' ADD COLUMN '{1}' VARCHAR(45) NULL  AFTER `Description`", table, name);

            ////create an OleDbDataAdapter to execute the query
            //dAdapter = new MySqlDataAdapter(query, connectionString);


            MySqlCommand command = connection.CreateCommand();

            string createCommnad = string.Format(@"ALTER TABLE `dd_parts`.`{0}` ADD COLUMN `{1}` VARCHAR(45) NULL  AFTER `{2}`", table, name, myColumnName);

            command.CommandText = createCommnad;
            MySqlDataReader Reader;
            connection.Open();
            try
            {
                Reader = command.ExecuteReader();
            }
            catch (Exception f)
            {
                MessageBox.Show(f.Message);
            }
            connection.Close();


        }

        private void DeleteColumn(string name, string table)
        {

            ////ALTER TABLE `dd_parts`.`47_256` ADD COLUMN `User` VARCHAR(45) NULL  AFTER `Description` ;

            ////create the database query
            //string query = string.Format("ALTER TABLE '{0}' ADD COLUMN '{1}' VARCHAR(45) NULL  AFTER `Description`", table, name);

            ////create an OleDbDataAdapter to execute the query
            //dAdapter = new MySqlDataAdapter(query, connectionString);


            MySqlCommand command = connection.CreateCommand();
            //ALTER TABLE `dd_parts`.`10_170` DROP COLUMN `hello2`


            string createCommnad = string.Format(@"ALTER TABLE `dd_parts`.`{0}` DROP COLUMN `{1}` ", table, name);

            command.CommandText = createCommnad;
            MySqlDataReader Reader;
            connection.Open();
            try
            {
                Reader = command.ExecuteReader();
            }
            catch (Exception f)
            {
                MessageBox.Show(f.Message);
            }
            connection.Close();




        }


        private void LoadTables(string table)
        {



            TreeNodeCollection nodes = mySelectedNode.Nodes;

            string[] nodeList = new string[9];
            int i = 0;
            foreach (TreeNode n in nodes)
            {


                char[] delimiters = new char[] { '(', ')' };
                string[] number = n.Text.Split(delimiters);

                if (number.Count() < 2)
                {
                    nodeList[i] = "0";
                }
                else
                {
                    nodeList[i] = number[1];
                }
                i++;

            }

            string query = "";
            foreach (string s in nodeList)
            {
                if (s != null)
                {
                    query = query + string.Format("SELECT * FROM `{0}{1}` union all ", table, s);
                }
            }

            query = query.Remove((query.Count() - 10), 10);  // remove 'union all at the end


            //create an OleDbDataAdapter to execute the query
            dAdapter = new MySqlDataAdapter(query, connectionString);

            //create a command builder
            MySqlCommandBuilder cBuilder = new MySqlCommandBuilder(dAdapter);

            //create a DataTable to hold the query results
            dTable = new DataTable();

            //---------------------------

            try
            {
                //fill the DataTable
                dAdapter.Fill(dTable);

                //the DataGridView
                //DataGridView dgView = new DataGridView();

                //BindingSource to sync DataTable and DataGridView
                BindingSource bSource = new BindingSource();

                //set the BindingSource DataSource
                bSource.DataSource = dTable;

                //set the DataGridView DataSource
                dataGridView1.DataSource = bSource;
                textBox5.Text = table;
                dataGridView1.Columns["CodeA"].Width = 50;
                dataGridView1.Columns["CodeB"].Width = 50;
                dataGridView1.Columns["CodeNumber"].Width = 75;
                dataGridView1.Columns["Description"].Width = 240;
                try
                {
                    dataGridView1.Columns["Value"].Width = 50;
                    dataGridView1.Columns["Mount"].Width = 50;
                    dataGridView1.Columns["Package"].Width = 55;
                    dataGridView1.Columns["RoHS"].Width = 40;
                    dataGridView1.Columns["Elements"].Width = 50;
                    dataGridView1.Columns["Tol"].Width = 50;
                    dataGridView1.Columns["Power"].Width = 50;
                    dataGridView1.Columns["Voltage"].Width = 50;
                }
                catch { }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                dAdapter.Update(dTable);
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
        }

        private void treeView1_DoubleClick(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            int length = textBox2.TextLength;


            if (length == 1)
            {
                if (Char.IsLetter(textBox2.Text[0]))
                {
                    //string form = string.Format(textBox2.Text);

                    LoadTable(textBox2.Text);
                }
            }

            if (length == 5)
            {

                if (Char.IsLetter(textBox2.Text[0]))
                {
                    //string form = string.Format(textBox2.Text);

                    LoadTable(textBox2.Text.Substring(0, 3));
                }
                else
                {

                    string form = string.Format("{0}", textBox2.Text);
                    LoadTables(form);
                }


            }
            if (length == 6)
            {
                if (Char.IsLetter(textBox2.Text[0]))
                {
                    //string form = string.Format(textBox2.Text);

                    LoadTable(textBox2.Text.Substring(0, 4));
                }
                else
                {



                    string form = string.Format("{0}", textBox2.Text);
                    LoadTable(form);

                    //string pdfFile = string.Format(@"\\STEWIE\do documents\DustyBookApp\{0}.pdf", textBox2.Text);

                    //try
                    //{
                    //    System.Diagnostics.Process.Start(pdfFile);
                    //}
                    //catch (Exception f)
                    //{
                    //    MessageBox.Show(f.Message);
                    //}
                }

            }
            this.Cursor = Cursors.Default;
        }

        private void createNewDatabaseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (textBox2.TextLength == 6)
            {
                MySqlCommand command = connection.CreateCommand();

                string createCommnad = string.Format(@"CREATE  TABLE `dd_parts`.`{0}` 
                                                      (`CodeA` INT(3) NOT NULL,
                                                       `CodeB` INT(4) NOT NULL,
                                                       `CodeNumber` INT NOT NULL ,
                                                       `Description` VARCHAR(45) NULL ,

                                                       PRIMARY KEY (`CodeNumber`, `CodeB`, `CodeA`),
                                                       UNIQUE INDEX `CodeNumber_UNIQUE` (`CodeNumber` ASC) );"
                                                      , textBox2.Text);

                command.CommandText = createCommnad;
                MySqlDataReader Reader;
                connection.Open();
                try
                {
                    Reader = command.ExecuteReader();
                }
                catch (Exception f)
                {
                    MessageBox.Show(f.Message);
                }
                connection.Close();
            }
        }

        private TreeNode m_OldSelectNode;
        private void treeView1_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            // Show menu only if the right mouse button is clicked.
            if (e.Button == MouseButtons.Right)
            {

                // Point where the mouse is clicked.
                Point p = new Point(e.X, e.Y);

                // Get the node that the user has clicked.
                TreeNode node = treeView1.GetNodeAt(p);
                if (node != null)
                {

                    // Select the node the user has clicked.
                    // The node appears selected until the menu is displayed on the screen.
                    m_OldSelectNode = treeView1.SelectedNode;
                    treeView1.SelectedNode = node;

                    // Find the appropriate ContextMenu depending on the selected node.
                    switch (node.Level)
                    {

                        case 5:
                            contextMenuStrip1.Show(treeView1, p);
                            break;

                        default: contextMenuStrip2.Show(treeView1, p);
                            break;
                    }

                    // Highlight the selected node.
                    treeView1.SelectedNode = m_OldSelectNode;
                    m_OldSelectNode = null;
                }
            }

        }

        private void pDFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            int length = textBox2.TextLength;

            if (length == 6)
            {
                //string form = string.Format("{0}", textBox2.Text);
                //LoadTable(form);

                // string pdfFile = string.Format(@"\\STEWIE\do documents\DustyBookApp\PDFs\{0}.pdf", textBox2.Text);

                string strNamedDestination = textBox2.Text; // Must be defined in PDF file.
                string book = textBox2.Text.Remove(2);
                string strFilePath = string.Format(@"\\STEWIE\do documents\DustyBookApp\PDFs\{0}.pdf", book);
                string strParams = " /n /A \"pagemode=bookmarks&nameddest=" + strNamedDestination + "\" \"" + strFilePath + "\"";

                try
                {
                    // System.Diagnostics.Process.Start(pdfFile);
                    System.Diagnostics.Process.Start("AcroRd32.exe", strParams);

                }
                catch (Exception f)
                {
                    MessageBox.Show(f.Message);
                }

            }


            //if (length == 5)
            //{
            //    //string form = string.Format("{0}", textBox2.Text);
            //    //LoadTable(form);

            //  //  string pdfFile = string.Format(@"\\STEWIE\do documents\DustyBookApp\PDFs\47_all.pdf", textBox2.Text);

            //    try
            //    {
            //       // System.Diagnostics.Process.Start(pdfFile, @"/A nameddest=47_216");  //   "/A \"page=2=OpenActions\");

            //        //AcroRd32.exe /n /A "pagemode=bookmarks@nameddest=47_006" "47_all.pdf"

            //        string strNamedDestination = "47_176"; // Must be defined in PDF file.
            //        string strFilePath = @"\\STEWIE\do documents\DustyBookApp\PDFs\47_all.pdf";
            //        string strParams = " /n /A \"pagemode=bookmarks&nameddest=" + strNamedDestination + "\" \"" + strFilePath + "\"";
            //        System.Diagnostics.Process.Start("AcroRd32.exe", strParams);


            //    }
            //    catch (Exception f)
            //    {
            //        MessageBox.Show(f.Message);
            //    }

            //}





            this.Cursor = Cursors.Default;
        }



        private void textBox2_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
        {



        }

        private void textBox2_PreviewKeyDown(object sender, System.Windows.Forms.PreviewKeyDownEventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            if (textBox4.Text == "")
            {
                treeView1.CollapseAll();
                //treeView1.TopNode.Expand();
                treeView1.Nodes[0].Expand();
                treeLevel = 0;

            }
            else
            {
                foreach (TreeNode _parentNode in treeView1.Nodes)
                {
                    treeLevel = 0;
                    expandTreeNode(_parentNode, treeLevel);
                }
            }
        }

        private void expandTreeNode(TreeNode _parentNode, int level)
        {
            foreach (TreeNode _childNode in _parentNode.Nodes)
            {
                if (_childNode.Text.StartsWith(string.Format("({0})", textBox4.Text.Substring(level, 1))))      //(this.fieldFilterTxtBx.Text))
                {


                    _childNode.Expand();
                    treeLevel++;
                    if (treeLevel == 2)
                    {
                        treeLevel++;
                    }
                    if (textBox4.Text.Count() > treeLevel)
                    {


                        expandTreeNode(_childNode, treeLevel);
                    }
                    // this.fieldsTree.Nodes.Add((TreeNode)_childNode.Clone());
                }
            }
        }

        private void addDataBaseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (textBox2.TextLength == 6)
            {
                MySqlCommand command = connection.CreateCommand();

                string createCommnad = string.Format(@"CREATE  TABLE `dd_parts`.`{0}` 
                                                      (`CodeA` INT(3) NOT NULL,
                                                       `CodeB` INT(4) NOT NULL,
                                                       `CodeNumber` INT NOT NULL ,
                                                       `Description` VARCHAR(400) NULL ,
                                                       `User` VARCHAR(45) NULL ,

                                                       PRIMARY KEY (`CodeNumber`, `CodeB`, `CodeA`),
                                                       UNIQUE INDEX `CodeNumber_UNIQUE` (`CodeNumber` ASC) );"
                                                      , textBox2.Text);

                command.CommandText = createCommnad;
                MySqlDataReader Reader;
                connection.Open();
                try
                {
                    Reader = command.ExecuteReader();
                }
                catch (Exception f)
                {
                    MessageBox.Show(f.Message);
                }
                connection.Close();
            }
        }

        private void dataGridView1_UserAddedRow(object sender, System.Windows.Forms.DataGridViewRowEventArgs e)
        {
            //try
            //{
            //    e.Row.Cells["User"].Value = User;

            //}
            //catch
            //{
            //}
        }

        private void dataGridView1_CellValueChanged(object sender, System.Windows.Forms.DataGridViewCellEventArgs e)
        {
            try
            {
                int changedRow = e.RowIndex;

                dataGridView1.Rows[changedRow].Cells["User"].Value = User;



            }
            catch
            {
            }
        }

        private void addBranchToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            int length = textBox2.TextLength;

            if (length < 6)
            {


                //string form = string.Format("{0}", textBox2.Text);
                //LoadTable(form);

                string pdfFile = string.Format(@"\\STEWIE\do documents\DustyBookApp\PDFs\{0}.pdf", textBox2.Text);

                try
                {
                    System.Diagnostics.Process.Start(pdfFile);
                }
                catch (Exception f)
                {
                    MessageBox.Show(f.Message);
                }

            }
            this.Cursor = Cursors.Default;
        }

        private void button5_Click(object sender, EventArgs e)
        {

            int fileCount = -1;
            do
            {
                fileCount++;
            }
            while (System.IO.File.Exists(textBox1.Text + (fileCount > 0 ? "(" + fileCount.ToString() + ")" : "")));

            //Not create a file
            // System.IO.File.Create(@"C:\temp.xml" + (fileCount > 0 ? "(" + (fileCount).ToString() + ")" : ""));

            System.IO.File.Copy(textBox1.Text, (textBox1.Text + (fileCount > 0 ? "(" + (fileCount).ToString() + ")" : "")), true);


            SerializeTreeView(treeView1, textBox1.Text);
            button5.Enabled = false;
        }



        private TreeNode mySelectedNode;
        private void treeView1_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            mySelectedNode = treeView1.GetNodeAt(e.X, e.Y);
        }

        private void editToolStripMenuItem_Click(object sender, EventArgs e)
        {

            if (mySelectedNode != null && mySelectedNode.Parent != null)
            {
                button5.Enabled = true;
                treeView1.SelectedNode = mySelectedNode;
                treeView1.LabelEdit = true;
                if (!mySelectedNode.IsEditing)
                {
                    mySelectedNode.BeginEdit();
                }
            }
            else
            {
                MessageBox.Show("No tree node selected or selected node is a root node.\n" +
                   "Editing of root nodes is not allowed.", "Invalid selection");
            }
        }


        private void treeView1_AfterLabelEdit(object sender,
         System.Windows.Forms.NodeLabelEditEventArgs e)
        {
            button5.Enabled = true;
            if (e.Label != null)
            {

                if (e.Label.Length > 0)
                {
                    if (e.Label.IndexOfAny(new char[] { '@', '.', ',', '!' }) == -1)
                    {
                        // Stop editing without canceling the label change.
                        e.Node.EndEdit(false);
                    }
                    else
                    {
                        /* Cancel the label edit action, inform the user, and 
                           place the node in edit mode again. */
                        e.CancelEdit = true;
                        MessageBox.Show("Invalid tree node label.\n" +
                           "The invalid characters are: '@','.', ',', '!'",
                           "Node Label Edit");
                        e.Node.BeginEdit();
                    }
                }
                else
                {
                    /* Cancel the label edit action, inform the user, and 
                       place the node in edit mode again. */
                    e.CancelEdit = true;
                    MessageBox.Show("Invalid tree node label.\nThe label cannot be blank",
                       "Node Label Edit");
                    e.Node.BeginEdit();
                }
            }
        }

        private void addToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (mySelectedNode != null && mySelectedNode.Parent != null)
            {
                button5.Enabled = true;
                treeView1.SelectedNode = mySelectedNode;

                TreeNode treeNode = new TreeNode("(0) - name");
                treeView1.SelectedNode.Nodes.Add(treeNode);
            }
            else
            {
                MessageBox.Show("No tree node selected or selected node is a root node.\n" +
                   "Editing of root nodes is not allowed.", "Invalid selection");
            }
        }

        private void moveUpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            button5.Enabled = true;
            Extensions.MoveUp(mySelectedNode);
        }

        private void moveDownToolStripMenuItem_Click(object sender, EventArgs e)
        {
            button5.Enabled = true;
            Extensions.MoveDown(mySelectedNode);
        }

        private void moveUpToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            button5.Enabled = true;
            Extensions.MoveUp(mySelectedNode);
        }

        private void moveDownToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            button5.Enabled = true;
            Extensions.MoveDown(mySelectedNode);
        }


        private void dataGridView1_ColumnHeaderMouseClick(object sender, System.Windows.Forms.DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {

                myColumnName = dataGridView1.Columns[e.ColumnIndex].HeaderText;


                contextMenuStrip3.Show(dataGridView1, myGridViewClick);



            }
        }

        private Point myGridViewClick;
        private string myColumnName;


        private void dataGridView1_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            myGridViewClick = new Point(e.X, e.Y);
        }

        private void insertToolStripMenuItem_Click(object sender, EventArgs e)
        {



            inputBox.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;

            var result = inputBox.ShowDialog();
            if (result == DialogResult.OK)
            {

            }

            string form = string.Format("{0}", textBox2.Text);

            InsertColumn(inputBox.inputValue, form);

        }

        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Delete column. Are you sure?",
       "Important Note",
       MessageBoxButtons.YesNo,
       MessageBoxIcon.Exclamation,
       MessageBoxDefaultButton.Button1);


            if (dialogResult == DialogResult.Yes)
            {
                string form = string.Format("{0}", textBox2.Text);
                DeleteColumn(myColumnName, form);
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }


        }

        private void dataGridView1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.C)
            {
                DataObject d = dataGridView1.GetClipboardContent();
                Clipboard.SetDataObject(d);
                e.Handled = true;
            }
            else if (e.Control && e.KeyCode == Keys.V)
            {
                string s = Clipboard.GetText();
                string[] lines = s.Split('\n');
                int noLines = lines.Count();
                //for (int i = 0; i < noLines; i++)
                //{
                //    dTable.Rows.Add();
                //}

                int row = dataGridView1.CurrentCell.RowIndex;
                int col = dataGridView1.CurrentCell.ColumnIndex;
                //int row = 0;
                //int col = 0;


                //foreach (string line in lines)
                //{
                //dTable.Rows.Add();
                //}



                foreach (string line in lines)
                {
                    if (row < dataGridView1.RowCount && line.Length > 0)
                    {



                        string[] cells = line.Split('\t');
                        for (int i = 0; i < cells.GetLength(0); ++i)
                        {
                            if (col + i < this.dataGridView1.ColumnCount)
                            {
                                try
                                {

                                    dataGridView1[col + i, row].Value = Convert.ChangeType(cells[i], dataGridView1[col + i, row].ValueType);
                                }
                                catch (Exception error)
                                {
                                    MessageBox.Show(error.Message);
                                    break;
                                }
                            }
                            else
                            {
                                break;
                            }
                        }
                        row++;
                    }
                    else
                    {
                        break;
                    }
                }

            }
        }

        private void dataGridView1_CellMouseUp(object sender, System.Windows.Forms.DataGridViewCellMouseEventArgs e)
        {
            //if (headerClick == false)
            //{
            // Show menu only if the right mouse button is clicked.
            if (e.Button == MouseButtons.Right)
            {

                // Point where the mouse is clicked.
                Point p = new Point(e.X, e.Y);

             //   contextMenuStrip4.Show(dataGridView1, p);




                //// Get the node that the user has clicked.
                //TreeNode node = treeView1.GetNodeAt(p);
                //if (node != null)
                //{

                //    // Select the node the user has clicked.
                //    // The node appears selected until the menu is displayed on the screen.
                //    m_OldSelectNode = treeView1.SelectedNode;
                //    treeView1.SelectedNode = node;

                //    // Find the appropriate ContextMenu depending on the selected node.
                //    switch (node.Level)
                //    {

                //        case 5:
                //            contextMenuStrip1.Show(treeView1, p);
                //            break;

                //        default: contextMenuStrip2.Show(treeView1, p);
                //            break;
                //    }

                //    // Highlight the selected node.
                //    treeView1.SelectedNode = m_OldSelectNode;
                //    m_OldSelectNode = null;
                //}
            }

        }

        private void getApprovalToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //            System.Net.Mail.MailMessage message = new System.Net.Mail.MailMessage();
            //            message.To.Add("andyaberdeen@davisderby.com");
            //            message.Subject = "Approval Needed";
            //            message.From = new System.Net.Mail.MailAddress("DustyBook App");
//            message.Body = @"<html>
//                         <body>
//                         <Table>
//                         <tr>
//                          <td> 
//                        <input type=""radio"" name=""sex"" value=""male"" /> Male<br />
//                        <input type=""radio"" name=""sex"" value=""female"" /> Female
//                          </td>
//                         </tr>
//                         </table>
//                         </body>
//                        </html>";



            try
            {
                Outlook.Application oApp = new Outlook.Application();

                // These 3 lines solved the problem
                Outlook.NameSpace ns = oApp.GetNamespace("MAPI");
                Outlook.MAPIFolder f = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                System.Threading.Thread.Sleep(5000); // test

                Outlook.MailItem oMsg = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
                oMsg.HTMLBody = "Hello, here is your message!";
                oMsg.Subject = "Approval Needed !";
                oMsg.VotingOptions = "Accept; Reject;";
                Outlook.Recipients oRecips = (Outlook.Recipients)oMsg.Recipients;
                Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add("andyaberdeen@davisderby.com");
                oRecip.Resolve();
                oMsg.Send();
                oRecip = null;
                oRecips = null;
                oMsg = null;
                oApp = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }







            //    System.Net.Mail.SmtpClient smtp = new System.Net.Mail.SmtpClient("yoursmtphost");
            //    smtp.Send(message);
            //}
            //}

        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            DataView DV = new DataView(dTable);
            DV.RowFilter = string.Format("{0} LIKE '%{1}%'", comboBox2.Text,textBox6.Text );

            dataGridView1.DataSource = DV;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }

    public static class Extensions
    {
        public static void MoveUp(this TreeNode node)
        {
            TreeNode parent = node.Parent;
            TreeView view = node.TreeView;
            if (parent != null)
            {
                int index = parent.Nodes.IndexOf(node);
                if (index > 0)
                {
                    parent.Nodes.RemoveAt(index);
                    parent.Nodes.Insert(index - 1, node);
                }
            }
            else if (node.TreeView.Nodes.Contains(node)) //root node
            {
                int index = view.Nodes.IndexOf(node);
                if (index > 0)
                {
                    view.Nodes.RemoveAt(index);
                    view.Nodes.Insert(index - 1, node);
                }
            }
        }

        public static void MoveDown(this TreeNode node)
        {
            TreeNode parent = node.Parent;
            TreeView view = node.TreeView;
            if (parent != null)
            {
                int index = parent.Nodes.IndexOf(node);
                if (index < parent.Nodes.Count - 1)
                {
                    parent.Nodes.RemoveAt(index);
                    parent.Nodes.Insert(index + 1, node);
                }
            }
            else if (view != null && view.Nodes.Contains(node)) //root node
            {
                int index = view.Nodes.IndexOf(node);
                if (index < view.Nodes.Count - 1)
                {
                    view.Nodes.RemoveAt(index);
                    view.Nodes.Insert(index + 1, node);
                }
            }
        }
    }
}
