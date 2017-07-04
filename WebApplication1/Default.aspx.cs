using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using System.IO;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Configuration;

namespace WebApplication1
{
    public partial class _Default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                lblMessage.Text = string.Empty;
                hlSurnames.NavigateUrl = string.Empty;
                hlAddresses.NavigateUrl = string.Empty;
            }
        }

        protected void btnUpload_Click(object sender, EventArgs e)
        {
            if (!fuNewFile.HasFile)
            {
                lblMessage.Text = "No File to Process";
                return;
            }
            else if ((System.IO.Path.GetExtension(fuNewFile.FileName) != ".csv")
                && (System.IO.Path.GetExtension(fuNewFile.FileName) != ".xlsx")
                && (System.IO.Path.GetExtension(fuNewFile.FileName) != ".xls"))
            {
                lblMessage.Text = "Incorrect File Type";
                return;
            }
            else
            {
                try
                {
                    string NewFileName = Guid.NewGuid().ToString();
                    String path = Server.MapPath("~/Files/");
                    fuNewFile.PostedFile.SaveAs(path + NewFileName + System.IO.Path.GetExtension(fuNewFile.FileName));
                    String FullPath = string.Format("{0}{1}{2}", path, NewFileName, System.IO.Path.GetExtension(fuNewFile.FileName));
                    ExcelImport(FullPath, NewFileName);
                }
                catch (Exception err)
                {
                    lblMessage.Text = err.Message;
                }
            }
        }

        protected void ExcelImport(string fullFilePath, string NewFileName)
        {
            String path = Server.MapPath("~/Files/");
            DataTable dt = new DataTable();
            ToData.ExcelToDataTable etdt = new ToData.ExcelToDataTable(fullFilePath);
            etdt.ReadDocument(150, ref dt);
            etdt.Dispose();

            var surnameCount = dt.AsEnumerable().GroupBy(row => row.Field<string>("LastName")).Select(grp => new { LastName = grp.Key, NameCount = grp.Select(row => row.Field<string>("FirstName")).Distinct().Count() });
            
            DataTable dtSurnameResults = new DataTable();
            dtSurnameResults.Columns.Add("LastName", typeof(string));
            dtSurnameResults.Columns.Add("NameCount", typeof(string));

            foreach (var r in surnameCount)
            {
                DataRow drSurnames = dtSurnameResults.NewRow();
                drSurnames["LastName"] = r.LastName;
                drSurnames["NameCount"] = r.NameCount;
                dtSurnameResults.Rows.Add(drSurnames);
            }

            DataView dvSurnames = dt.DefaultView;
            dvSurnames.Sort = "NameCount DESC, LastName ASC";
            DataTable dtSurnames = dvSurnames.ToTable();
            GridView gvSurnames = new GridView();
            gvSurnames.DataSource = dtSurnames;
            string sSurnames = string.Empty;
            int iSurnames = 0, iCount = 0;

            for (int a = 0; a < gvSurnames.HeaderRow.Cells.Count; a++)
            {
                if (gvSurnames.HeaderRow.Cells[a].Text == "LastName") { iSurnames = a; }
                if (gvSurnames.HeaderRow.Cells[a].Text == "NameCount") { iCount = a; }
            }

            for (int e = 0; e < gvSurnames.Rows.Count; e++)
            {
                sSurnames = string.Format("{0}{1}, {2}<br>", sSurnames, gvSurnames.Rows[e].Cells[iSurnames].ToString(), gvSurnames.Rows[e].Cells[iCount].ToString());
            }

            System.IO.File.WriteAllText(path + NewFileName + ".txt", sSurnames);
            hlSurnames.NavigateUrl = "~/Files/" + NewFileName + "_Surnames.txt";
            hlSurnames.Visible = true;

            var addressCols = from ac in dt.AsEnumerable() select ac.Field<string>("Address");

            DataTable dtAddressResults = new DataTable();
            dtSurnameResults.Columns.Add("Address", typeof(string));
            dtSurnameResults.Columns.Add("AddressSortBy", typeof(string));

            foreach (var t in addressCols)
            {
                DataRow drAddresses = dtAddressResults.NewRow();
                drAddresses["Address"] = t;
                drAddresses["AddressSortBy"] = t.Substring(t.IndexOf(" ") + 1);
                dtAddressResults.Rows.Add(drAddresses);
            }

            DataView dvAddresses = dtAddressResults.DefaultView;
            dvAddresses.Sort = "AddressSortBy ASC";
            DataTable dtAddresses = dvAddresses.ToTable();
            GridView gvAddresses = new GridView();
            gvAddresses.DataSource = dtAddresses;
            string Adressess = string.Empty;
            int Address = 0;

            for (int o = 0; o < gvAddresses.HeaderRow.Cells.Count; o++)
            {
                if (gvAddresses.HeaderRow.Cells[o].Text == "Address")
                {
                    Address = o;
                    o = gvAddresses.HeaderRow.Cells.Count;
                }
            }

            for (int i = 0; i < gvAddresses.Rows.Count; i++)
            {
                Adressess = string.Format("{0}{1}<br>", Adressess, gvAddresses.Rows[i].Cells[Address].ToString());
            }

            System.IO.File.WriteAllText(path + NewFileName + ".txt", Adressess);
            hlAddresses.NavigateUrl = "~/Files/" + NewFileName + "_Addresses.txt";
            hlAddresses.Visible = true;
        }
    }
}