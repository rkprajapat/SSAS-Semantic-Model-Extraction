using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Windows.Forms;

using System.Text;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;


namespace SSAS2012_MyBI
{
    public partial class WebForm3 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

            //load url arguments in variables
            string INSTANCE_NAME = @Request.QueryString["instance"];
            string CATALOG_NAME = Request.QueryString["catalog"];
            string CUBE_NAME = Request.QueryString["cube"];
            
            //Create Excel file name
            string ConnectionName = INSTANCE_NAME.Replace("\\","_") + "_" + CATALOG_NAME + "_" + CUBE_NAME;
            Response.Write(ConnectionName);

       
            //Create Workbook
            string filename = Server.MapPath(@"tmp/" + ConnectionName + ".xlsx");
            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filename, SpreadsheetDocumentType.Workbook);

            // Add a WorkbookPart to the document.
            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "mySheet" };
            sheets.Append(sheet);

            //Add a connectionPart to the workbookpart
            ConnectionsPart connectionsPart1 = workbookpart.AddNewPart<ConnectionsPart>();
            Connections connections1 = new Connections();            
            Connection connection1 = new Connection() { Id = (UInt32Value)1U, KeepAlive = true, Name = ConnectionName, Type = (UInt32Value)5U, RefreshedVersion = 5, Background = true };            
            DatabaseProperties databaseProperties1 = new DatabaseProperties() { Connection = "Provider=MSOLAP.4;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=" + CATALOG_NAME + ";Data Source=" + @INSTANCE_NAME + ";MDX Compatibility=1;Safety Options=2;MDX Missing Member Mode=Error", Command = CUBE_NAME, CommandType = (UInt32Value)1U };
            OlapProperties olapProperties1 = new OlapProperties() { SendLocale = true, RowDrillCount = (UInt32Value)1000U };
            connection1.Append(databaseProperties1);
            connection1.Append(olapProperties1);
            connections1.Append(connection1);
            connectionsPart1.Connections = connections1;

            //Add a PivottableCache part
            PivotTableCacheDefinitionPart pivotTableCacheDefinitionPart1 = workbookpart.AddNewPart<PivotTableCacheDefinitionPart>();
            PivotCacheDefinition pivotCacheDefinition1 = new PivotCacheDefinition() { SaveData = false, BackgroundQuery = true, SupportSubquery = true, SupportAdvancedDrill = true };
            pivotTableCacheDefinitionPart1.PivotCacheDefinition = pivotCacheDefinition1;

            workbookpart.Workbook.Save();
            // Close the document.
            spreadsheetDocument.Close();

            Response.Clear();
            Response.AddHeader("content-disposition", "attachment; filename=" + @"D:\MyBI\SSAS\SSAS2012_MyBI\tmp\test.xlsx");
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.WriteFile(Server.MapPath(@"tmp/" + ConnectionName + ".xlsx"));
            Response.End();
        }
    }
}
