using System;
using System.Reflection;
using System.Windows.Media.Imaging;
using System.Windows.Media;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows;
using System.Drawing;
using System.Collections.Generic;
using System.IO;
using System.Xml;


namespace ART_Configurateur
{
    //for the securite, if the userdomain isn't "ARTELIAGROUP", the command is not availability
    public class CommandEnabler : IExternalCommandAvailability
    {
        #region IExternalCommandAvailability Members

        public bool IsCommandAvailable(UIApplication uiApp, CategorySet catSet)
        {
            string userDomainName = System.Environment.UserDomainName;

            if (userDomainName == "ARTELIAGROUP")
                return true;
            else
                return false;
        }

        #endregion
    }


    [Transaction(TransactionMode.Manual)]
    public class Application : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication a)
        {
            #region pour mise à jour 
            string pathXML = @"C:\Users\" + System.Environment.UserName + @"\AppData\Roaming\Autodesk\Revit\Addins";
            //juge whether there is the original fiche XML, if not, copy the fiche
            if (!File.Exists(pathXML + @"\Revit_ART_Configurateur.xml"))
            {
                File.Copy(pathXML + @"\Revit_ART_Configurateur_V1.xml", pathXML + @"\Revit_ART_Configurateur.xml");
                File.Delete(pathXML + @"\Revit_ART_Configurateur_V1.xml");
            }
            else
            {
                File.Delete(pathXML + @"\Revit_ART_Configurateur_V1.xml");
            }
            //load XML
            XmlDocument orgineXML = new XmlDocument();
            orgineXML.Load(pathXML + @"\Revit_ART_Configurateur.xml");

            DirectoryInfo dirXML = new DirectoryInfo(pathXML);
            FileInfo[] _fileList = dirXML.GetFiles(); // file list of XML configurateur

            string xmlKey = ".xml";
            string configuKey = "Revit_ART_Configurateur_";
            //juge whether there are other fiche XML
            for (int i = 0; i < _fileList.Length; i++)
            {
                if (_fileList[i].Name.Contains(xmlKey) && _fileList[i].Name.Contains(configuKey))
                {
                    //if there are, juge whether is added in the fiche orginal XML
                    XmlNodeList elemList = orgineXML.GetElementsByTagName(_fileList[i].Name);
                    if ( elemList.Count == 0)
                    {
                        XmlDocument newDoc = new XmlDocument();
                        newDoc.Load(_fileList[i].DirectoryName + "/" + _fileList[i].Name);
                        XmlNode root1 = orgineXML.DocumentElement;
                        XmlNode root2 = newDoc.DocumentElement;

                        //给根节点创建子节点
                        XmlElement version = orgineXML.CreateElement(_fileList[i].Name);
                        //将添加到根节点
                        root1.AppendChild(version);

                        foreach (XmlElement xnItem in root2)
                        {
                            XmlNode root = orgineXML.ImportNode(xnItem, true);
                            root1.AppendChild(root);
                        }

                        File.Delete(_fileList[i].DirectoryName + "/" + _fileList[i].Name);//添加成功后把多余的xml文件删除
                        orgineXML.Save(pathXML + @"\Revit_ART_Configurateur.xml");
                    }

                    else
                    {
                        File.Delete(_fileList[i].DirectoryName + "/" + _fileList[i].Name);
                    }
                }
            }
            #endregion

            // Method to add Tab and Panel 
            RibbonPanel panel = ribbonPanel(a, "ARTELIA");
            // Reflection to look for this assembly path 
            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;
            // Add button to panel 
            PushButton button = panel.AddItem(new PushButtonData("ButtonConfigurateur", "Configurateur\ndu ruban", thisAssemblyPath, "ART_Configurateur.ConfigurateurPlugin")) as PushButton;

            //Add securite
            button.AvailabilityClassName = "ART_Configurateur.CommandEnabler";

            button.ToolTip = "Configuration de la visibilité des plugins dans le ruban Artelia";
            // Reflection of path to image 
            string imagePath = Path.GetDirectoryName(thisAssemblyPath) + @"\ART_Configurateur.png";
            Uri uriImage = new Uri(imagePath, UriKind.RelativeOrAbsolute);
            BitmapImage largeimage = new BitmapImage(uriImage);
            button.LargeImage = largeimage;

            //help
            ContextualHelp contextHelp = new ContextualHelp(ContextualHelpType.Url, "https://connect.arteliagroup.com/community/user-group-bim/revit-user-group");
            button.SetContextualHelp(contextHelp);

            return Result.Succeeded;
        }

        public RibbonPanel ribbonPanel(UIControlledApplication a, string panelName)
        {
            // Tab name 
            string tab = "Artelia";

            // Empty ribbon panel 
            RibbonPanel ribbonPanel = null;
            // Try to create ribbon tab. 
            try
            {
                //a.CreateRibbonPanel("My Test Tools");
                a.CreateRibbonTab(tab);
            }
            catch { }
            // Try to create ribbon panel. 
            try
            {
                RibbonPanel panel = a.CreateRibbonPanel(tab, panelName);
            }
            catch { }
            // Search existing tab for your panel. 
            List<RibbonPanel> panels = a.GetRibbonPanels(tab);
            foreach (RibbonPanel p in panels)
            {
                if (p.Name == panelName)
                {
                    ribbonPanel = p;
                }
            }
            //return panel 
            return ribbonPanel;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

    }
}
