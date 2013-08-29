using System;
using System.Net;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.IO;
using System.Drawing;
using System.Text;

using Microsoft.AnalysisServices.AdomdClient;
using Microsoft.AnalysisServices;
using Microsoft.Office.Interop.Excel;


namespace SSAS2012_MyBI
{
    public partial class WebForm1 : System.Web.UI.Page
    {        

        protected void Page_Load(object sender, EventArgs e)
        {
            
        }

        protected void ButtonSubmit_Click(object sender, EventArgs e)
        {

            string InstanceName = TextBoxSource.Text;
                        
            //List Catalogs(DBs)
            string strMDX = "select * from $system.DBSCHEMA_CATALOGS";
            string strConnString = "Datasource=" + InstanceName; 

            System.Data.DataTable dtDB = GetDataTable(strConnString, strMDX); //returns data table with DB details
            
            //Enumerate cubes in each database
            ArrayList DBList = ReturnValueFromDataTable(dtDB, 0); //Get DB Name (Catalog_Name) from data table
            dtDB.Dispose();

            System.Data.DataTable dtCubes = new System.Data.DataTable();
            System.Data.DataTable dtPartial = new System.Data.DataTable();

            foreach (string DBName in DBList)
            {
                strMDX = "SELECT * FROM $system.MDSchema_Cubes WHERE CUBE_SOURCE=1";
                strConnString = "Datasource=" + TextBoxSource.Text + ";Catalog=" + DBName; 
                dtPartial = GetDataTable(strConnString, strMDX); //get all Cubes for the DB DBName
                DataColumn CubeLink = dtPartial.Columns.Add("Link"); // Add a column "Link" in the data table
                CubeLink.SetOrdinal(0); //make new column as first column in table
           
                dtCubes.Merge(dtPartial);
                dtPartial.Dispose();
            }

            GridViewResults.AutoGenerateColumns = true;
            GridViewResults.DataSource = dtCubes; //returns data table with DB details
            GridViewResults.DataBind();
            dtCubes.Dispose();   
        }



        //Create Link Button in grid view to open excel
        protected void gvResults_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            if (e.Row.DataItem != null)
            {

                string INSTANCE_NAME = TextBoxSource.Text;
                string CATALOG_NAME = e.Row.Cells[1].Text;
                string CUBE_NAME = e.Row.Cells[3].Text;
                
                //below code is not used as a hyperlink is used for ODC file, we can also use a imagebutton to do run a client function
                var FormLink = new System.Web.UI.WebControls.ImageButton();
                FormLink.ImageUrl = "img/xl.png";
                FormLink.Width = 25;
                FormLink.Height = 25;

                //pass control to Javascript function ExcelLoad with cube parameters
                FormLink.OnClientClick = "ExcelLoad(\"" + INSTANCE_NAME.Replace("\\","\\\\") + "\",\"" + CATALOG_NAME + "\",\"" + CUBE_NAME + "\"); return false;";                  
                e.Row.Cells[0].Controls.Add(FormLink);
            }
        }



        //Write and ODC file (office document connector), not being used right now, but can be used to create odc files
        protected String WriteODC(string catalog, string cube) {
            string filePath = AppDomain.CurrentDomain.BaseDirectory + "tmp/";
            string ODCFile = Path.ChangeExtension(Path.GetRandomFileName(),".odc");
            
            StreamWriter w;
            w = File.CreateText(filePath + ODCFile);
            string instance = TextBoxSource.Text;
            w.WriteLine("<meta name=SourceType content=OLEDB><xml id=docprops><o:DocumentProperties xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns=\"http://www.w3.org/TR/REC-html40\">");
            w.WriteLine("<o:Name>" + instance.Replace("\\","_") + "_" + catalog + "_" + cube + "</o:Name>");
            w.WriteLine("</o:DocumentProperties></xml>");
            w.WriteLine("<xml id=msodc><odc:OfficeDataConnection  xmlns:odc=\"urn:schemas-microsoft-com:office:odc\"  xmlns=\"http://www.w3.org/TR/REC-html40\"><odc:Connection odc:Type=\"OLEDB\">");
            w.WriteLine("<odc:ConnectionString>Provider=MSOLAP.4;Integrated Security=SSPI;Persist Security Info=True;Data Source=" + instance + ";Initial Catalog=" + catalog + "</odc:ConnectionString>");
            w.WriteLine("<odc:CommandType>Cube</odc:CommandType><odc:CommandText>" + cube + "</odc:CommandText></odc:Connection></odc:OfficeDataConnection></xml>");
            w.Flush();
            w.Close();

            var EscapedPath = "tmp/" + ODCFile;
            return  @EscapedPath; 
        }


        protected void LoadODC(string filepath) {
            Response.Write(filepath);
            //Intialize Excel
            Application ObjExcel = new Application();
            WorkbookConnection objWBConnection;
            ObjExcel.Visible = true;
            objWBConnection = ObjExcel.ActiveWorkbook.Connections.AddFromFile2(@filepath);
            objWBConnection.Refresh();
            ObjExcel = null; 
        
        }

     
        //Loop through a data table and return value for a particular column in array
        protected ArrayList ReturnValueFromDataTable(System.Data.DataTable dtOlap, int ColumnID)
        {
            ArrayList ValueArray = new ArrayList();
            
            foreach (DataRow row in dtOlap.Rows) // Loop over the rows.
            {
                ValueArray.Add(row.ItemArray[ColumnID].ToString());                
            }
            return ValueArray;
        }


        //Below subroutine will return a data table with results. It takes the DB connection string and query as input
        protected System.Data.DataTable GetDataTable(string ConnString, string query)
        {
            System.Data.DataTable dtOlap = new System.Data.DataTable();
            //Connect to Analysis Server
            AdomdConnection conn = new AdomdConnection(ConnString);
            
            System.Diagnostics.Debug.WriteLine(conn.ConnectionString);

            try
            {                
                conn.Open();

                //Create adomd command using connection and MDX query
                AdomdCommand cmd = new AdomdCommand(query, conn);
                AdomdDataAdapter adr = new AdomdDataAdapter(cmd);
                adr.Fill(dtOlap);                
            }
            catch (InvalidCastException e) 
            {
                Response.Write("Access denied on " + ConnString);
            }
            finally
            {
                //Close DB Connection
                conn.Close();                
            }
            return dtOlap;
        }
       
       
    }
}

