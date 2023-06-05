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
    public partial class Register : System.Web.UI.Page
    {
        private static string connectionString = ConfigurationManager.ConnectionStrings["ConnectionStr"].ToString();
        DataTable dt = new DataTable();
        DataTable dtCheckName = new DataTable();

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

            if (tbPwd2.Text == "" || tbPwd2.Text == string.Empty)
            {
                lbMsg.Text = "Please Confirm the Password";
                return false;
            }

            if (tbPwd.Text != tbPwd2.Text)
            {
                lbMsg.Text = "Password must be same";
                return false;
            }

            if (checkusername() == false)
            {
                lbMsg.Text = "Username already exist";
                return false;
            }

            else
            {
                return true;
            }
        }

        public Boolean checkusername()
        {
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();
                string query = "select * from tbl_user where username=@User";
                SqlCommand sqlCmd = new SqlCommand(query, con);

                sqlCmd.Parameters.AddWithValue("@User", tbUN.Text);
                sqlCmd.ExecuteNonQuery();

                SqlDataAdapter da = new SqlDataAdapter(sqlCmd);

                try
                {
                    da.Fill(dtCheckName);
                }
                catch (Exception ex)
                {

                }
                con.Close();
            }

            if(dtCheckName.Rows.Count > 0)
            {
                return false;
            }
            else
            {
                return true;
            }
            
        }

        protected void BtnConfirm_Click(object sender, EventArgs e)
        {
            try
            {
                if (validation() == true)
                {
                    string Encrypt = security.MD5hash(tbPwd.Text);

                    using (SqlConnection con = new SqlConnection(connectionString))
                    {
                        con.Open();
                        string query = "insert into tbl_user (username,Password) values (@User,@password)";
                        SqlCommand sqlCmd = new SqlCommand(query, con);

                        sqlCmd.Parameters.AddWithValue("@User", tbUN.Text);
                        sqlCmd.Parameters.AddWithValue("@password", Encrypt);
                        sqlCmd.ExecuteNonQuery();

                        con.Close();
                    }

                    lbMsg.Text = "Thank you, your account has been created";
                    lbMsg.ForeColor = System.Drawing.Color.Green;
                }
            }
            catch
            {

            }
            
        }

        protected void BtnBackLogin_Click(object sender, EventArgs e)
        {
            Response.Redirect("./Login.aspx");
        }

        
    }
}
