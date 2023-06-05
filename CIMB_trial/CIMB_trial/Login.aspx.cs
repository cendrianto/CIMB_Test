using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.SqlClient;
using System.Configuration;
using System.Data;

namespace CIMB_trial
{
    public partial class Login : System.Web.UI.Page
    {
        private static string connectionString = ConfigurationManager.ConnectionStrings["ConnectionStr"].ToString();
        DataTable dt = new DataTable();

        protected void Page_Load(object sender, EventArgs e)
        {
                        
        }

        void clearvalue()
        {
            lbMsg.Text = "";
            tbPwd.Text = "";
            tbUN.Text = "";
        }

        public Boolean validation()
        {
            if (tbUN.Text == "" || tbUN.Text == string.Empty)
            {
                lbMsg.Text = "Username can't be empty";
                return false;
            }

            if (tbPwd.Text == "" || tbPwd.Text == string.Empty)
            {
                lbMsg.Text = "Password can't be empty";
                return false;
            }
            else
            {
                return true;
            }            
        }

        protected void btnLogin_Click(object sender, EventArgs e)
        {
            try
            {
                if (validation() == true)
                {
                    string Encrypt = security.MD5hash(tbPwd.Text);

                    using (SqlConnection con = new SqlConnection(connectionString))
                    {
                        con.Open();
                        string query = "select * from tbl_user where username=@User and password=@password";
                        SqlCommand sqlCmd = new SqlCommand(query, con);

                        sqlCmd.Parameters.AddWithValue("@User", tbUN.Text);
                        sqlCmd.Parameters.AddWithValue("@password", Encrypt);
                        sqlCmd.ExecuteNonQuery();

                        SqlDataAdapter da = new SqlDataAdapter(sqlCmd);

                        try
                        {
                            da.Fill(dt);
                        }
                        catch (Exception ex)
                        {

                        }
                        con.Close();
                    }

                    if (dt.Rows.Count > 0)
                    {
                        Session["username"] = tbUN.Text.ToString();
                        Response.Redirect("./Dashboard.aspx");
                    }
                    else
                    {
                        lbMsg.Text = "Please check again Username & Password!";
                    }
                }
            }
            catch
            {

            }
            
        }

        protected void btnRegister_Click(object sender, EventArgs e)
        {
            Response.Redirect("./Register.aspx");
        }
    }
}