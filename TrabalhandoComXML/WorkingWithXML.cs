using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = System.Windows.Forms.Application;

namespace TrabalhandoComXML
{
    public class WorkingWithXML
    {
        public SAPbouiCOM.Application oApplication;
        public SAPbouiCOM.Form oForm;
        public void SetaApplication()
        {
            SAPbouiCOM.SboGuiApi sboGuiApi = null;
            string sConnectionString = null;
            sboGuiApi = new SAPbouiCOM.SboGuiApi();
            sConnectionString = System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));
            sboGuiApi.Connect(sConnectionString);
            oApplication = sboGuiApi.GetApplication(-1);
        }       
        private void LoadFromXML(ref string pFileName)
        {
            System.Xml.XmlDocument oXmlDoc = null;
            oXmlDoc = new System.Xml.XmlDocument();
            string sPath = null;
            sPath = System.IO.Directory.GetParent(Application.StartupPath).ToString();
            sPath = System.IO.Directory.GetParent(sPath).ToString();

         //C: \Users\SAPB1\source\repos\SAPB1\TrabalhandoComXML\bin
            oXmlDoc.Load(sPath + "\\" + pFileName);
            string sXML = oXmlDoc.InnerXml.ToString();

            oApplication.LoadBatchActions(ref  sXML);
        }

        private void SaveAsXML(ref SAPbouiCOM.Form pForm)
        {
            System.Xml.XmlDocument oXmlDoc = null;
            oXmlDoc = new System.Xml.XmlDocument();
            string sXmlString = null;
            sXmlString = pForm.GetAsXML();


            oXmlDoc.LoadXml(sXmlString);

            string sPath = null;
            sPath = System.IO.Directory.GetParent(Application.StartupPath).ToString();

            oXmlDoc.Save(sPath + @"\FormSimples_1.xml");
        //    oApplication.SetStatusBarMessage("Dir: " + sPath, SAPbouiCOM.BoMessageTime.bmt_Short, false);
        
        }
        public  WorkingWithXML() {            

            SetaApplication();
            string sXMLFormName = "FormSimples.xml";
            LoadFromXML(ref sXMLFormName);

            oForm = oApplication.Forms.Item("MeuFormSimples");
            oForm.Visible = true;

            SaveAsXML(ref oForm);

        }
    }
}
