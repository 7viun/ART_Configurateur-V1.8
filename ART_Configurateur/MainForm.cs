using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace ART_Configurateur
{
    public partial class ConfigurateurForm : Form
    {
        public List<string> listImageName;
        public List<string> listImageAdresse;
        public string thisAssemblyPath;

        string pathImage;
        string pathXML;
        XmlDocument xmlDoc;

        public ConfigurateurForm()
        {
            InitializeComponent();
            //listImageName = new List<string>();
            //set the height of dataGridView for image
            //设置自动换行  
            this.dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            //设置自动调整高度  
            //this.dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            this.dataGridView1.RowTemplate.Height = 65;

            //for adding all of plugins of Artelia in the DataGrisView
            #region loding plugins
            //string dddd = "2018";//version de plugin
            //pathImage = @"C:\Users\danni.liu\Desktop\MON\TESTimages\";
            thisAssemblyPath = Assembly.GetExecutingAssembly().Location;
            pathImage = Path.GetDirectoryName(thisAssemblyPath);

            
            string pngKey = ".png";
            string nameKey = "Revit_ART_";
            //string xmlKey = ".xml";
            //string configuKey = "Revit_ART_Configurateur_";

            xmlDoc = new XmlDocument();

            pathXML = @"C:\Users\" + System.Environment.UserName + @"\AppData\Roaming\Autodesk\Revit\Addins" + @"\Revit_ART_Configurateur.xml";

            xmlDoc.Load(pathXML);

            //DirectoryInfo dirXML = new DirectoryInfo(@"C:\Users\" + System.Environment.UserName + @"\AppData\Roaming\Autodesk\Revit\Addins");
            //FileInfo[] _fileList = dirXML.GetFiles(); // file list of XML configurateur

            DirectoryInfo dirIcon = new DirectoryInfo(pathImage);
            FileInfo[] fileList = dirIcon.GetFiles(); // file list of image icone DLL

            /*for (int i = 0; i < _fileList.Length; i++)
            {
                if (_fileList[i].Name.Contains(xmlKey) && _fileList[i].Name.Contains(configuKey))
                {
                    XmlDocument newDoc = new XmlDocument();
                    newDoc.Load(_fileList[i].DirectoryName + "/" + _fileList[i].Name);
                    XmlNode root1 = xmlDoc.DocumentElement;
                    XmlNode root2 = newDoc.DocumentElement;

                    //给根节点Books创建子节点
                    XmlElement version = xmlDoc.CreateElement(_fileList[i].Name);
                    //将book添加到根节点
                    root1.AppendChild(version);

                    foreach (XmlElement xnItem in root2)
                    {
                        XmlNode root = xmlDoc.ImportNode(xnItem, true);
                        root1.AppendChild(root);
                    }
                    
                    File.Delete(_fileList[i].DirectoryName + "/" + _fileList[i].Name);//添加成功后把多余的xml文件删除
                    xmlDoc.Save(pathXML);

                }
            }*/

            for (int i = 0; i < fileList.Length; i++)
            {
                if (fileList[i].Name.Contains(pngKey) && fileList[i].Name.Contains(nameKey))
                {

                    int index = this.dataGridView1.Rows.Add();
                    this.dataGridView1.Rows[index].Cells[0].Value = Image.FromFile(fileList[i].FullName);
                    //Name like: Revit_ART_****
                    //去后缀带前缀的插件名
                    string pluginName = fileList[i].Name.Substring(0, fileList[i].Name.LastIndexOf("."));
                    this.dataGridView1.Rows[index].Cells[1].Value = pluginName;
                    //Name like : ****
                    //不带后缀前缀的插件名
                    //string result = Regex.Match(_fileList[i].Name, "(?<=Revit_ART_).*?(?=.png)").Value;
                    //this.dataGridView1.Rows[index].Cells[1].Value = result;

                    //for keeping the information of setting
                    #region 保持上一次设置
                    XmlElement root = xmlDoc.DocumentElement;
                    XmlNode node = root.SelectSingleNode(pluginName);

                    string description = node.Attributes["Description"].Value;
                    this.dataGridView1.Rows[index].Cells[2].Value = description;

                    if (node.Attributes["Mode"].Value == "yes")
                    {
                        dataGridView1.Rows[index].Cells[3].Value = true;
                    }
                    #endregion
                }
            }
            #endregion

        }

        //Modify the file according to the selected conditions in the list
        private void button1_Click(object sender, EventArgs e)
        {
            XmlElement root = xmlDoc.DocumentElement;

            //change the mode to yes 
            foreach (string s in GetIsVisible())
            {

                XmlNode node = root.SelectSingleNode(s);
                node.Attributes["Mode"].Value = "yes";
            }

            //change the mode to no
            foreach (string s in GetNotVisible())
            {

                XmlNode node = root.SelectSingleNode(s);
                node.Attributes["Mode"].Value = "no";
            }
            xmlDoc.Save(pathXML);
            
            MessageBox.Show("La configuration a été enregistrée. Veuillez redémarrer Revit pour qu'elle devienne effective.");
        }

        //Get items in the list of DataGrisView which are selected
        private IEnumerable<string> GetIsVisible()
        {
            int count = dataGridView1.Rows.Count;
            for (int i = 0; i < count; i++)
            {
                DataGridViewRow row = dataGridView1.Rows[i];
                if (((Boolean)((DataGridViewCheckBoxCell)row.Cells[3]).FormattedValue == true))
                {
                    // 返回Name列
                    yield return row.Cells[1].FormattedValue.ToString();
                }
            }
        }

        //Get items in the list of DataGrisView which are not selected
        private IEnumerable<string> GetNotVisible()
        {
            int count = dataGridView1.Rows.Count;
            for (int i = 0; i < count; i++)
            {
                DataGridViewRow row = dataGridView1.Rows[i];
                if (((Boolean)((DataGridViewCheckBoxCell)row.Cells[3]).FormattedValue == false))
                {
                    // 返回Name列
                    yield return row.Cells[1].FormattedValue.ToString();
                }
            }
        }

    }
}

