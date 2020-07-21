using System;
using System.Collections;

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Architecture;


namespace ART_Configurateur
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Journaling(Autodesk.Revit.Attributes.JournalingMode.NoCommandData)]
    public class ConfigurateurPlugin : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //即使没有活动projet也可用
            //Autodesk.Revit.ApplicationServices.Application revitApp = commandData.Application.ActiveUIDocument.Application.Application;
            //Autodesk.Revit.DB.Document revirDoc = commandData.Application.ActiveUIDocument.Document;
            try
            {
                //prepare data

                ConfigurateurForm displayForm = new ConfigurateurForm();
                displayForm.Show();

                return Autodesk.Revit.UI.Result.Succeeded;
            }
            catch (Exception e)
            {
                message = e.Message;
                return Autodesk.Revit.UI.Result.Failed;
            }
        }
    }
}