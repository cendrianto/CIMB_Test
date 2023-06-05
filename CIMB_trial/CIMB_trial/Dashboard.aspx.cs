using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.SqlClient;
using System.Configuration;
using System.Data.OleDb;
using System.Text;

namespace CIMB_trial
{
    public partial class Dashboard : System.Web.UI.Page
    {
        private static string connectionString = ConfigurationManager.ConnectionStrings["ConnectionStr"].ToString();
        protected void Page_Load(object sender, EventArgs e)
        {
            if (Session["username"] == null)
            {
                Response.Redirect("./Login.aspx");
            }
        }

        protected void BtnUpload_Click(object sender, EventArgs e)
        {
            try
            {
                if (Fileupload.HasFile)
                {
                    Fileupload.SaveAs("D:\\testing\\" + Fileupload.FileName);
                }
                else
                {
                    lbMsg.Text = "Please Choose Excel File first";
                }
            }
            catch
            {

            }
            
        }

        protected void btnLogout_Click(object sender, EventArgs e)
        {
            Session["username"] = null;
            Response.Redirect("./Login.aspx");
        }

        protected void BtnImportExcel_Click (object sender, EventArgs e)
        {
            try
            {
                string DirUpload = "";
                if (Fileupload.HasFile)
                {
                    DirUpload = "D:\\testing\\" + Fileupload.FileName;
                    Fileupload.SaveAs("D:\\testing\\" + Fileupload.FileName);
                    ImportDataFromExcel(DirUpload);
                }
                else
                {
                    lbMsg.Text = "Please Choose Excel File first";
                }
            }
            catch (Exception ex)
            {

            }           
            
        }

        public void ImportDataFromExcel(string excelFilePath)
        {
            string tblSQL_penjualan = "tbl_penjualan";
            string tblSQL_penjualan_details = "tbl_penjualan_details";
            string tblSQL_barang = "tbl_barang";
            string tblSQL_ticket = "tbl_ticket";
            string tblSQL_ticket_process = "tbl_ticket_process";

            string QueryExcel_Penjualan = "select id_trx,no_invoice,total_berat,ongkos_kirim,total_harga,total_harga_beli,kode_user,alamat_penerima,tgl_kirim,id_ekspedisi,jenis_pengiriman,tgl_trx from [penjualan$]";
            string QueryExcel_Penjualan_detail = "select id_trx_detail,id_trx,no_invoice,id_produk,jml_barang,berat,harga_satuan,diskon,harga from [penjualan_detail$]";
            string QueryExcel_barang = "select id_produk,nama,id_kategori,berat,harga_beli,stok,harga_jual from [barang$]";
            string QueryExcel_ticket = "select ticket_code,ticket_date,customer_id,subject,id_product,issue from [ticket$]";
            string QueryExcel_ticket_process = "select ticket_code,status,user_id,update_date from [ticket_process$]";

            StringBuilder sb = new StringBuilder();

            if (ImportFile(QueryExcel_Penjualan_detail, tblSQL_penjualan_details, excelFilePath) == true)
            {
                string Msg = "File Penjualan Details imported successfully.";
                sb.Append(Msg);
            }

            if (ImportFile(QueryExcel_Penjualan, tblSQL_penjualan, excelFilePath) == true)
            {
                string Msg = "File Penjualan imported successfully.";
                sb.Append(Msg);
            }

            if (ImportFile(QueryExcel_barang, tblSQL_barang, excelFilePath) == true)
            {
                string Msg = "File Barang imported successfully.";
                sb.Append(Msg);
            }

            if (ImportFile(QueryExcel_ticket, tblSQL_ticket, excelFilePath) == true)
            {
                string Msg = "File ticket imported successfully.";
                sb.Append(Msg);
            }

            if (ImportFile(QueryExcel_ticket_process, tblSQL_ticket_process, excelFilePath) == true)
            {
                string Msg = "File ticket process imported successfully.";
                sb.Append(Msg);
            }

            sb = sb.Replace(".", ";" + "<br/>");
            lbMsg.Text = sb.ToString();
            lbMsg.ForeColor = System.Drawing.Color.Green;

          
        }

        protected Boolean ImportFile(string QueryExcel, string TblSQL, string Filepath)
        {
            try
            {
                string excelconnectionstring = @"provider=Microsoft.ACE.OLEDB.12.0;data source=" + Filepath +
                ";extended properties=" + "\"Excel 12.0;HDR=YES;\"";

                OleDbConnection oledbconn = new OleDbConnection(excelconnectionstring);
                OleDbCommand oledbcmd = new OleDbCommand(QueryExcel, oledbconn);
                oledbconn.Open();
                OleDbDataReader dr = oledbcmd.ExecuteReader();
                SqlBulkCopy bulkcopy = new SqlBulkCopy(connectionString);
                bulkcopy.DestinationTableName = TblSQL;
                while (dr.Read())
                {
                    bulkcopy.WriteToServer(dr);
                }
                dr.Close();
                oledbconn.Close();
                

                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }
    }
}