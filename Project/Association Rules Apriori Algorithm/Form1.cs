using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Association_Rules_Algorithm
{
    public partial class Form1 : Form
    {
        Dictionary<string, int> HashTableOFSetOf1Items = new Dictionary<string, int>();
        List<ItemSet> DataBase = new List<ItemSet>();
        ItemSet Itemset = new ItemSet();

        List<Item> ItemsWithCount = new List<Item>();
        Item ItemWithCount = new Item();

        Dictionary<string, int> HashTableOFSetOf2Items = new Dictionary<string, int>();
        List<SetOfItems> SetOf2ItemsWithCount = new List<SetOfItems>();
        SetOfItems SOIsWithCount = new SetOfItems();

        List<SetOfItems> SetOf3ItemsWithCount = new List<SetOfItems>();

        List<SetOfItems> AssociationRulesSets = new List<SetOfItems>();

        int SupportThreshold = 0;
        int MinimumConfidenceThreshold = 0;

        bool ExcelChecker = false;

        public Form1()
        {
            InitializeComponent();
        }
        private void ExcelButton_Click(object sender, EventArgs e)
        {
            ExcelChecker = true;
            string filePath = string.Empty;
            string fileExt = string.Empty;
            OpenFileDialog file = new OpenFileDialog(); //open dialog to choose file  
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK) //if there is a file choosen by the user  
            {
                filePath = file.FileName; //get the path of the file  
                fileExt = Path.GetExtension(filePath); //get the file extension  
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    try
                    {
                        DataTable dtExcel = new DataTable();
                        dtExcel = ReadExcel(filePath, fileExt); //read excel file
                        dataGridView1.DataSource = null;
                        dataGridView1.Rows.Clear();
                        dataGridView1.Columns.Clear();
                        dataGridView1.Refresh();
                        dataGridView1.DefaultCellStyle.Font = new Font("Microsoft Sans Serif", 10);
                        dataGridView1.Visible = true;
                        dataGridView1.DataSource = dtExcel;
                        dataGridView1.Columns[0].HeaderCell.Style.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);
                        dataGridView1.Columns[1].HeaderCell.Style.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);
                        dataGridView1.EnableHeadersVisualStyles = false;
                        dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Aqua;
                        ExcelChecker = true;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                        ExcelChecker = false;
                    }
                }
                else
                {
                    MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error  
                    ExcelChecker = false;
                }
            }
        }
        public DataTable ReadExcel(string fileName, string fileExt)
        {
            string conn = string.Empty;
            DataTable dtexcel = new DataTable();
            if (fileExt.CompareTo(".xls") == 0)
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=NO';"; //for above excel 2007  
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [Sheet1$]", con); //here we read data from sheet1  
                    oleAdpt.Fill(dtexcel); //fill excel data into dataTable  
                }
                catch
                {
                    MessageBox.Show("Error.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error 
                }
            }
            return dtexcel;
        }
        public void CollectDataFromDataGridView()
        {
            DataBase = new List<ItemSet>();
            for (int r = 0; r < dataGridView1.Rows.Count - 1; r++)
            {
                string ID = dataGridView1[0, r].Value.ToString();
                List<string> SplitItemsText = dataGridView1[1, r].Value.ToString().Split(new[] { ' ', ',' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                Itemset = new ItemSet(ID, SplitItemsText);
                DataBase.Add(Itemset);
            }
        }
        public void ARA()
        {
            //#1
            CountOfEachItem();
            //#2
            SupportThreshold = SupportThreshold * DataBase.Count / 100;
            //MessageBox.Show("SupportThreshold = " + SupportThreshold);
            RemoveItemWithCountLessThanSupportThreshold();
            //#3
            Form_2_ItemSet_And_FindTheirOccurrences();
            //#4
            Remove_2_ItemSetWithCountLessThanSupportThreshold();
            //#5
            Form_3_ItemSet_From_2_ItemSet_And_FindTheirOccurrences();
            //#6 Generate Association Rules For the frequent
            GenerateAssociationRules();
        }
        public void CreateData()
        {
            Itemset = new ItemSet("1", new List<string> { "I1", "I2", "I3" });
            DataBase.Add(Itemset);
            Itemset = new ItemSet("2", new List<string> { "I2", "I3", "I4" });
            DataBase.Add(Itemset);
            Itemset = new ItemSet("3", new List<string> { "I4", "I5" });
            DataBase.Add(Itemset);
            Itemset = new ItemSet("4", new List<string> { "I1", "I2", "I4" });
            DataBase.Add(Itemset);
            Itemset = new ItemSet("5", new List<string> { "I1", "I2", "I3", "I5" });
            DataBase.Add(Itemset);
            Itemset = new ItemSet("6", new List<string> { "I1", "I2", "I3", "I4" });
            DataBase.Add(Itemset);
        }
        public void CountOfEachItem()
        {
            HashTableOFSetOf1Items = new Dictionary<string, int>();
            for (int i = 0; i < DataBase.Count; i++)
            {
                for (int j = 0; j < DataBase[i].Items.Count; j++)
                {
                    if (!HashTableOFSetOf1Items.ContainsKey(DataBase[i].Items[j]))
                    {
                        HashTableOFSetOf1Items.Add(DataBase[i].Items[j], 1);
                    }
                    else
                    {
                        HashTableOFSetOf1Items[DataBase[i].Items[j]] += 1;
                    }
                }
            }

            ItemsWithCount = new List<Item>();
            ICollection key = HashTableOFSetOf1Items.Keys;

            foreach (string k in key)
            {
                ItemWithCount = new Item(k, Convert.ToInt32(HashTableOFSetOf1Items[k]));
                ItemsWithCount.Add(ItemWithCount);
                //MessageBox.Show(ItemsWithCount[ItemsWithCount.Count - 1].I + " " + ItemsWithCount[ItemsWithCount.Count - 1].Count);
            }
        }
        public void RemoveItemWithCountLessThanSupportThreshold()
        {
            for (int i = 0; i < ItemsWithCount.Count; i++)
            {
                if (ItemsWithCount[i].Count < SupportThreshold)
                {
                    ItemsWithCount.RemoveAt(i);
                    i--;
                }
            }
        }
        public void Form_2_ItemSet_And_FindTheirOccurrences()
        {
            SetOf2ItemsWithCount = new List<SetOfItems>();
            string[] Data = new string[2];
            List<string> Items = new List<string>();
            for (int i = 0; i < ItemsWithCount.Count; i++)
            {
                Items.Add(ItemsWithCount[i].I);
            }
            Combination(Items, Items.Count, 2, 0, Data, 0, ref SetOf2ItemsWithCount);
            for (int i = 0; i < SetOf2ItemsWithCount.Count ; i++)
            {
                for (int k = 0; k < DataBase.Count; k++)
                {
                    bool Check = true;
                    for (int j = 0; j < SetOf2ItemsWithCount[i].Items.Count; j++)
                    {
                        if (!DataBase[k].Items.Contains(SetOf2ItemsWithCount[i].Items[j]))
                        {
                            Check = false;
                            break;
                        }
                    }
                    if (Check)
                    {
                        SetOf2ItemsWithCount[i].Count += 1;
                    }
                }
            }
        }
        public void Remove_2_ItemSetWithCountLessThanSupportThreshold()
        {
            //And save them in HashTable
            HashTableOFSetOf2Items = new Dictionary<string, int>();
            for (int i = 0; i < SetOf2ItemsWithCount.Count; i++)
            {
                if (SetOf2ItemsWithCount[i].Count < SupportThreshold)
                {
                    SetOf2ItemsWithCount.RemoveAt(i);
                    i--;
                }
                else
                    HashTableOFSetOf2Items.Add(SetOf2ItemsWithCount[i].Items[0] + ", " + SetOf2ItemsWithCount[i].Items[1] + ", ", SetOf2ItemsWithCount[i].Count);
            }
        }
        public void Form_3_ItemSet_From_2_ItemSet_And_FindTheirOccurrences()
        {
            Hashtable H = new Hashtable();
            List<string> Items = new List<string>();
            for (int i = 0; i < SetOf2ItemsWithCount.Count; i++)
            {
                for (int j = 0; j < SetOf2ItemsWithCount[i].Items.Count; j++)
                {
                    if (!H.ContainsKey(SetOf2ItemsWithCount[i].Items[j]))
                    {
                        H.Add(SetOf2ItemsWithCount[i].Items[j], SetOf2ItemsWithCount[i].Items[j]);
                        Items.Add(SetOf2ItemsWithCount[i].Items[j]);
                    }
                }
            }

            SetOf3ItemsWithCount = new List<SetOfItems>();
            string[] Data = new string[3];
            Combination(Items, Items.Count, 3, 0, Data, 0, ref SetOf3ItemsWithCount);

            for (int i = 0; i < SetOf3ItemsWithCount.Count; i++)
            {
                bool Check = true;
                for (int j = 0; j < SetOf3ItemsWithCount[i].Items.Count - 1; j++)
                {
                    for (int k = j + 1; k < SetOf3ItemsWithCount[i].Items.Count; k++)
                    {
                        for (int l = 0; l < SetOf2ItemsWithCount.Count; l++)
                        {
                            if (SetOf2ItemsWithCount[l].Items.All(new List<string> { SetOf3ItemsWithCount[i].Items[j], SetOf3ItemsWithCount[i].Items[k] }.Contains))
                            {
                                Check = true;
                                break;
                            }
                            else
                            {
                                Check = false;
                            }
                        }
                        if (!Check)
                            break;
                    }
                    if (!Check)
                        break;
                }
                if (!Check)
                {
                    SetOf3ItemsWithCount.RemoveAt(i);
                    i--;
                }
            }
        }
        public void GenerateAssociationRules()
        {
            //Frequant of SetOf3Items
            for (int i = 0; i < SetOf3ItemsWithCount.Count; i++)
            {
                for (int j = 0; j < DataBase.Count; j++)
                {
                    bool Check = true;
                    for (int k = 0; k < SetOf3ItemsWithCount[i].Items.Count; k++)
                    {
                        if (!DataBase[j].Items.Contains(SetOf3ItemsWithCount[i].Items[k]))
                        {
                            Check = false;
                            break;
                        }
                    }
                    if (Check)
                        SetOf3ItemsWithCount[i].Count++;
                }
            }
            string CollectAll = "Generate Association Rules" + "\r\n";
            for (int i = 0; i < SetOf3ItemsWithCount.Count; i++)
            {
                AssociationRulesSets = new List<SetOfItems>();
                for (int j = 1; j < 3; j++)
                {
                    string[] Data = new string[j];
                    All_Combination_To_r(SetOf3ItemsWithCount[i].Items, SetOf3ItemsWithCount[i].Items.Count, j, 0, Data, 0);
                }
                CollectAll += i + "] {";
                for (int j = 0; j < SetOf3ItemsWithCount[i].Items.Count; j++)
                {
                    CollectAll += SetOf3ItemsWithCount[i].Items[j] + ", ";
                }
                CollectAll += "} ---> Count = " + SetOf3ItemsWithCount[i].Count + "\r\n";
                for (int z = 0; z < AssociationRulesSets.Count; z++)
                {
                    if (AssociationRulesSets[z].Items.Count == 2)
                    {
                        string Collect = "";
                        for (int y = 0; y < AssociationRulesSets[z].Items.Count; y++)
                        {
                            Collect += AssociationRulesSets[z].Items[y] + ", ";
                        }
                        AssociationRulesSets[z].Count = SetOf3ItemsWithCount[i].Count * 100 / HashTableOFSetOf2Items[Collect];
                        if (AssociationRulesSets[z].Count >= MinimumConfidenceThreshold)
                            CollectAll +=  "{" + Collect + "} ---> Count = " + HashTableOFSetOf2Items[Collect] + " ---> Confidence = " + SetOf3ItemsWithCount[i].Count.ToString() + " * 100 / " + HashTableOFSetOf2Items[Collect].ToString() + " = " + AssociationRulesSets[z].Count + "%" + "\r\n";
                    }
                    else if (AssociationRulesSets[z].Items.Count == 1)
                    {
                        AssociationRulesSets[z].Count = SetOf3ItemsWithCount[i].Count * 100 / HashTableOFSetOf1Items[AssociationRulesSets[z].Items[0]];
                        if (AssociationRulesSets[z].Count >=MinimumConfidenceThreshold)
                            CollectAll += "{" + AssociationRulesSets[z].Items[0] + "} ---> Count = " + HashTableOFSetOf1Items[AssociationRulesSets[z].Items[0]] + " ---> Confidence = " + SetOf3ItemsWithCount[i].Count.ToString() + " * 100 / " + HashTableOFSetOf1Items[AssociationRulesSets[z].Items[0]].ToString() + " = " + AssociationRulesSets[z].Count + "%" + "\r\n";
                    }
                    else
                        MessageBox.Show("Error!!!");
                }
                CollectAll += "\r\n";
            }
            ARRrichTextBox.Text = CollectAll;
        }

        /* arr[] ---> Input Array 
        data[] ---> Temporary array to store 
        current combination start & end ---> 
        Staring and Ending indexes in arr[] 
        index ---> Current index in data[] 
        r ---> Size of a combination to be 
        printed */
        public void Combination(List<string> Items, int n, int r, int index, string[] Data, int i, ref List<SetOfItems> SetItems)
    {

            // Current combination is ready to 
            // be printed, print it 
            if (index == r)
            {
                SOIsWithCount = new SetOfItems();
                for (int j = 0; j < r; j++)
                    SOIsWithCount.Add(Data[j]);
                SetItems.Add(SOIsWithCount);

                return;
            }

            // When no more elements are there 
            // to put in data[] 
            if (i >= n)
                return;

            // current is included, put next 
            // at next location 
            Data[index] = Items[i];
            Combination(Items, n, r, index + 1, Data, i + 1,ref SetItems);

            // current is excluded, replace 
            // it with next (Note that i+1  
            // is passed, but index is not 
            // changed) 
            Combination(Items, n, r, index, Data, i + 1,ref SetItems);
        }
        public void All_Combination_To_r(List<string> Items, int n, int r, int index, string[] Data, int i)
        {

            // Current combination is ready to 
            // be printed, print it 
            if (index == r)
            {
                SOIsWithCount = new SetOfItems();
                for (int j = 0; j < index; j++)
                    SOIsWithCount.Add(Data[j]);
                AssociationRulesSets.Add(SOIsWithCount);

                return;
            }

            // When no more elements are there 
            // to put in data[] 
            if (i >= n)
                return;

            // current is included, put next 
            // at next location 
            Data[index] = Items[i];
            All_Combination_To_r(Items, n, r, index + 1, Data, i + 1);

            // current is excluded, replace 
            // it with next (Note that i+1  
            // is passed, but index is not 
            // changed) 
            All_Combination_To_r(Items, n, r, index, Data, i + 1);
        }

        private void GenerateButton_Click(object sender, EventArgs e)
        {
            Hashtable HashTableOFSetOf1Items = new Hashtable();
            DataBase = new List<ItemSet>();

            ItemsWithCount = new List<Item>();

            Hashtable HashTableOFSetOf2Items = new Hashtable();
            List<SetOfItems> SetOf2ItemsWithCount = new List<SetOfItems>();

            SetOf3ItemsWithCount = new List<SetOfItems>();

            AssociationRulesSets = new List<SetOfItems>();

            SupportThreshold = 0;
            MinimumConfidenceThreshold = 0;

            bool Check = true;

            Check = UserEnter_ST_And_MCT();
            if (Check)
            {
                if (ExcelChecker)
                {
                    CollectDataFromDataGridView();
                }
                else
                {
                    MessageBox.Show("Error In Excel", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (DataBase.Count > 0)
                    ARA();
            }

        }
        public bool UserEnter_ST_And_MCT()
        {
            SupportThreshold = 0;
            if (!int.TryParse(STTextBox.Text, out SupportThreshold))
            {
                MessageBox.Show("Insert Integer Number Of 'SupportThershold:'");
                return false;
            }
            else
            {
                if (SupportThreshold < 0 || SupportThreshold > 100)
                {
                    MessageBox.Show("'Support Threshold:' Range Must Be From 0 To 100");
                    return false;
                }
            }
            MinimumConfidenceThreshold = 0;
            if (!int.TryParse(MCTTextBox.Text, out MinimumConfidenceThreshold))
            {
                MessageBox.Show("Insert Integer Number Of 'MinimumConfidenceThreshold:'");
                return false;
            }
            else
            {
                if (MinimumConfidenceThreshold < 0 || MinimumConfidenceThreshold > 100)
                {
                    MessageBox.Show("'Minimum Confidence Threshold:' Range Must Be From 0 To 100");
                    return false;
                }
            }
            return true;
        }
    }
}
