using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Etat_Balance
{
    public partial class Frm_Details: Form
    {
        string Compte = "";
        public Frm_Details()
        {
            InitializeComponent();
        }
        public Frm_Details(string value)
        {
            InitializeComponent();
            Compte = value;
            DataTable dtGetBalance = new DataTable();
            dtGetBalance = GetBalanceByValue(value);
            foreach (DataRow row in dtGetBalance.Rows)
            {
                txt_compte.Text = Compte;
                txt_libelle.Text = row["LIBELLE"].ToString();
                txt_ran.Text = (Convert.ToDouble(row["DEBIT/RAN"].ToString()) - Convert.ToDouble(row["CREDIT/RAN"].ToString())).ToString();
                txt_janvier.Text = (Convert.ToDouble(row["DEBIT/JANV"].ToString()) - Convert.ToDouble(row["CREDIT/JANV"].ToString())).ToString();
                txt_fevrier.Text = (Convert.ToDouble(row["DEBIT/FEV"].ToString()) - Convert.ToDouble(row["CREDIT/FEV"].ToString())).ToString();

                txt_mars.Text = (Convert.ToDouble(row["DEBIT/MARS"].ToString()) - Convert.ToDouble(row["DEBIT/MARS2"].ToString())).ToString();
                txt_avril.Text = (Convert.ToDouble(row["DEBIT/AVR"].ToString()) - Convert.ToDouble(row["CREDIT/AVR"].ToString())).ToString();

                txt_mai.Text = (Convert.ToDouble(row["DEBIT/MAI"].ToString()) - Convert.ToDouble(row["CREDIT/MAI"].ToString())).ToString();
                txt_total.Text = row["SOLDE"].ToString();

            }
        }
        static DataTable GetBalanceByValue(string value)
        {
            DataTable dt = new DataTable();
            using (SqlConnection cnx = new SqlConnection(Connection.ConnectionString_EtatsComptables1))
            {

                SqlDataAdapter DA = new SqlDataAdapter($@"select * from [dbo].[BALANCE] where COMPTE like '{value}' ", cnx);
                DA.Fill(dt);
                return dt;
            }
        }
    }
}
