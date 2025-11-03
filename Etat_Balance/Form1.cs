using EtatsComptables;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Z.Dapper.Plus;
using ExcelDataReader;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using DocumentFormat.OpenXml.Drawing;

namespace Etat_Balance
{
    public partial class Form1 : Form
    {
        DataTableCollection tables;
        List<Grand_Livre> ListGrand_Livre = new List<Grand_Livre>();

        public Form1()
        {
            InitializeComponent();

        }
        static DataTable GetBalanceByVqlue(string value)
        {
            DataTable dt = new DataTable();
            using (SqlConnection cnx = new SqlConnection(Connection.ConnectionString_EtatsComptables1))
            {

                SqlDataAdapter DA = new SqlDataAdapter($@"
                                                select Compte,left(Compte,1) as 'C',left(Compte,2) as 'R',
                                                left(Compte,3) as 'B', LIBELLE,CONVERT(decimal(18,2),
                                                sum([DEBIT/RAN]-[CREDIT/RAN])) AS 'SOLDE RAN',
                                                CONVERT(decimal(18,2),sum([DEBIT/JANV]-[CREDIT/JANV])) as 'AU 31/01/2025',
                                                CONVERT(decimal(18,2),sum([DEBIT/FEV]-[CREDIT/FEV])) as 'AU 28/02/2025',
                                                CONVERT(decimal(18,2),sum([DEBIT/MARS]-[DEBIT/MARS2])) as 'AU 31/03/2025',
                                                CONVERT(decimal(18,2),sum([DEBIT/AVR]-[CREDIT/AVR])) as 'AU 30/04/2025',
                                                CONVERT(decimal(18,2),sum([DEBIT/MAI]-[CREDIT/MAI]))as 'AU 31/05/2025',
                                                CONVERT(decimal(18,2),
                                                sum([DEBIT/RAN]-[CREDIT/RAN])+sum([DEBIT/JANV]-[CREDIT/JANV])+
                                                sum([DEBIT/FEV]-[CREDIT/FEV])+
                                                sum([DEBIT/MARS]-[DEBIT/MARS2])+
                                                sum([DEBIT/AVR]-[CREDIT/AVR]) +sum([DEBIT/MAI]-[CREDIT/MAI])) as 'SOLDE FINAL'
                                                from [dbo].[BALANCE] 
                                                group by Compte ,LIBELLE
                                                having COMPTE like '%{value}%'
                                                order by Compte ", cnx);
                DA.Fill(dt);
                return dt;
            }
        }
        private void txt_chercheBalance_TextChanged(object sender, EventArgs e)
        {
            dgv_balance.DataSource = GetBalanceByVqlue(txt_chercheBalance.Text);

        }
        private void btn_parcourir_Click_2(object sender, EventArgs e)
        {
            OpenFileDialoge(txtPath.Text, tables, cboSheet);
        }
        void OpenFileDialoge(string txtname, DataTableCollection table, ComboBox cmb)
        {
            try
            {
                using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx" })
                {
                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        txtPath.Text = ofd.FileName;
                        using (var stream = File.Open(ofd.FileName, FileMode.Open, FileAccess.Read))
                        {
                            using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                            {
                                DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                                {
                                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                                    {
                                        UseHeaderRow = true
                                    }
                                });

                                tables = result.Tables;
                                cmb.Items.Clear();
                                foreach (DataTable item in tables)
                                    cmb.Items.Add(item.TableName);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void cboSheet_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            try
            {

                btn_importer.Enabled = true;
                DataTable dt = tables[cboSheet.SelectedItem.ToString()];
                if (dt != null)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (dt.Rows[i]["Compte"].ToString() != "")
                        {
                            Grand_Livre grand_Livre = new Grand_Livre();
                            grand_Livre.Compte = dt.Rows[i]["Compte"].ToString();
                            grand_Livre.NomCompte = dt.Rows[i]["Nom Compte"].ToString();
                            grand_Livre.Date = dt.Rows[i]["Date"].ToString();
                            grand_Livre.Type_Numéro_de_pièce = dt.Rows[i]["Type / Numéro de pièce"].ToString();
                            grand_Livre.Libellé = dt.Rows[i]["Libellé"].ToString();
                            grand_Livre.Libellé_ligne = dt.Rows[i]["Libellé ligne"].ToString();
                            grand_Livre.Date_Echeance = dt.Rows[i]["Date Echeance"].ToString();
                            grand_Livre.Tiers = dt.Rows[i]["Tiers"].ToString();
                            grand_Livre.Site = dt.Rows[i]["Site"].ToString();
                            grand_Livre.Journal = dt.Rows[i]["Journal"].ToString();
                            grand_Livre.Lettre = dt.Rows[i]["Lettre"].ToString();
                            grand_Livre.Etat = dt.Rows[i]["Etat"].ToString();
                            grand_Livre.Débit = dt.Rows[i]["Débit"].ToString();
                            grand_Livre.Crédit = dt.Rows[i]["Crédit"].ToString();
                            grand_Livre.Solde_débiteur = dt.Rows[i]["Solde débiteur"].ToString();
                            grand_Livre.Solde_créditeur = dt.Rows[i]["Solde créditeur"].ToString();
                            ListGrand_Livre.Add(grand_Livre);
                        }
                    }
                    dgv_grandLivre.DataSource = ListGrand_Livre;
                    dgv_grandLivre.Columns[0].Visible = false;
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void btn_importer_Click_1(object sender, EventArgs e)
        {
            using (SqlConnection connection = new SqlConnection(Connection.ConnectionString_EtatsComptables1))
            {
                connection.Open();

                SqlCommand com1 = new SqlCommand($@"delete  from [dbo].[BALANCE] ", connection);
                com1.ExecuteNonQuery();

                SqlCommand com2 = new SqlCommand($@" delete  from [dbo].[Grand_Livre]", connection);
                com2.ExecuteNonQuery();

                SqlCommand com3 = new SqlCommand($@" delete  from [dbo].[GrandLivreGrouping] ", connection);
                com3.ExecuteNonQuery();
            }

            btn_importer.Enabled = false;
            btn_parcourir.Enabled = false;
            dgv_grandLivre.FirstDisplayedScrollingRowIndex = dgv_grandLivre.RowCount - 1;
            var listSuiviConsommation = dgv_grandLivre.DataSource as List<Grand_Livre>;
            using (IDbConnection db = new SqlConnection(Connection.ConnectionString_EtatsComptables1))
            {
                db.BulkInsert(listSuiviConsommation);
            }

            try
            {

                using (SqlConnection connection = new SqlConnection(Connection.ConnectionString_EtatsComptables1))
                {
                    connection.Open();




                    SqlCommand com = new SqlCommand($@"insert into GrandLivreGrouping  select TRIM(Compte) Compte,TRIM(NomCompte) NomCompte,Journal,CONVERT(varchar(10), Date, 103) Date,
                                sum(convert(float,replace( Replace(Translate(REPLACE(REPLACE(isnull(Débit,0), CHAR(13), ''),
                                CHAR(10), ''), ' -\','???'),'?',''),',','.'))) as Débit,
                                sum(convert(float, replace( Replace(Translate(REPLACE(REPLACE(isnull(Crédit,0), CHAR(13), ''), CHAR(10), ''),
                                ' -\','???'),'?',''),',','.'))) as Crédit
                                from [dbo].[Grand_Livre]
                                group by Compte,NomCompte,Date,Journal
                                order by TRIM(Compte) ", connection);
                    com.ExecuteNonQuery();

                    SqlCommand cmd2 = new SqlCommand("prc_EtatBalance2", connection);
                    cmd2.CommandTimeout = 0;
                    cmd2.ExecuteNonQuery();
                    btn_importer.Enabled = true;
                    btn_parcourir.Enabled = true;


                    DataTable dtBalance = new DataTable();
                    dtBalance = GetALLBalance();
                    dgv_balance.DataSource = dtBalance;
                    dgv_balance.Columns[6].HeaderCell.Value = "JANVIER\nAU 31/01/" + DateTime.Now.Year;
                    dgv_balance.Columns[7].HeaderCell.Value = "FEVRIER\nAU 28/02/" + DateTime.Now.Year;
                    dgv_balance.Columns[8].HeaderCell.Value = "MARS\nAU 31/03/" + DateTime.Now.Year;
                    dgv_balance.Columns[9].HeaderCell.Value = "AVRIL\nAU 30/04/" + DateTime.Now.Year;
                    dgv_balance.Columns[10].HeaderCell.Value = "MAI\nAU 31/05/" + DateTime.Now.Year;
                    dgv_balance.Columns[11].HeaderCell.Value = "01/01/" + DateTime.Now.Year + " au 31/05/" + DateTime.Now.Year + "\nSOLDE FINAL";
                    tabControl1.SelectedTab = tabPage2;

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("" + ex.Message);
            }
        }
        private void dgv_balance_CellDoubleClick_1(object sender, DataGridViewCellEventArgs e)
        {
            string compte = "";
            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = dgv_balance.Rows[e.RowIndex];
                compte = row.Cells[1].Value.ToString();
                Frm_Details frm = new Frm_Details(compte);
                frm.ShowDialog();

            }
        }
        static DataTable GetALLBalance()
        {
            DataTable dt = new DataTable();
            using (SqlConnection cnx = new SqlConnection(Connection.ConnectionString_EtatsComptables1))
            {

                SqlDataAdapter DA = new SqlDataAdapter($@"
                                            select Compte,left(Compte,1) as 'C',left(Compte,2) as 'R',
                                            left(Compte,3) as 'B', LIBELLE,CONVERT(decimal(18,2),
                                            sum([DEBIT/RAN]-[CREDIT/RAN])) AS 'SOLDE RAN',
                                            CONVERT(decimal(18,2),sum([DEBIT/JANV]-[CREDIT/JANV])) as 'AU 31/01/2025',
                                            CONVERT(decimal(18,2),sum([DEBIT/FEV]-[CREDIT/FEV])) as 'AU 28/02/2025',
                                            CONVERT(decimal(18,2),sum([DEBIT/MARS]-[DEBIT/MARS2])) as 'AU 31/03/2025',
                                            CONVERT(decimal(18,2),sum([DEBIT/AVR]-[CREDIT/AVR]) )as 'AU 30/04/2025',
                                            CONVERT(decimal(18,2),sum([DEBIT/MAI]-[CREDIT/MAI]))as 'AU 31/05/2025',
                                            CONVERT(decimal(18,2),
                                            sum([DEBIT/RAN]-[CREDIT/RAN])+sum([DEBIT/JANV]-[CREDIT/JANV])+
                                            sum([DEBIT/FEV]-[CREDIT/FEV])+sum([DEBIT/MARS]-[DEBIT/MARS2])+
                                            sum([DEBIT/AVR]-[CREDIT/AVR]) +sum([DEBIT/MAI]-[CREDIT/MAI])) as 'SOLDE FINAL' 
                                            from [dbo].[BALANCE] 
                                            group by Compte ,LIBELLE
                                            order by Compte ", cnx);
                DA.Fill(dt);
                return dt;
            }
        }

        private void txt_chercheBalance_TextChanged_1(object sender, EventArgs e)
        {
            dgv_balance.DataSource = GetBalanceByVqlue(txt_chercheBalance.Text);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            #region
            //double totalHT = 0;
            //double totalTTC = 0;
            //double totalTVA = 0;
            //double totallIGNE = 0;
            //for (int i = 0; i < dgv_impression.Rows.Count - 1; i++)
            //{
            //    totalHT += Convert.ToDouble(dgv_impression.Rows[i].Cells[12].Value);
            //    totalTTC += Convert.ToDouble(dgv_impression.Rows[i].Cells[14].Value);
            //    totalTVA += Convert.ToDouble(dgv_impression.Rows[i].Cells[15].Value);
            //    totallIGNE += Convert.ToDouble(dgv_impression.Rows[i].Cells[17].Value);
            //}
            #endregion
            using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel Workbook|*.xlsx" })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        // Create a new Excel workbook
                        System.Data.DataTable dt = DataGridView_To_Datatable(dgv_balance);
                        dt.exportToExcel(sfd.FileName);

                        using (XLWorkbook workbook = new XLWorkbook(sfd.FileName))
                        {
                            var hoja = workbook.Worksheets.Worksheet(1);


                            // Set the height of the second row to 50

                            #region
                            //string valuecelltotalHT1 = "Q" + (dgv_impression.Rows.Count + 4).ToString();
                            //string valuecelltotalHT2 = "R" + (dgv_impression.Rows.Count + 4).ToString();
                            //hoja.Cell(valuecelltotalHT1).Value = " SUM MONTANT HT";
                            //hoja.Cell(valuecelltotalHT2).Value = totalHT;
                            hoja.Row(1).InsertRowsAbove(1);
                            //hoja.Row(2).Cells("A1").Value = "test";
                            var row2 = hoja.Row(1);
                            //row2.Style.Fill.BackgroundColor = XLColor.DarkOrange;
                            row2.Height = 30;
                            //row2.Cell("A1").Value = "test";
                            //hoja.Cell(2, 1).Value = "Initial Value";
                            hoja.Cells("A1").Value = "";
                            hoja.Cells("B1").Value = "";
                            hoja.Cells("C1").Value = "";
                            hoja.Cells("D1").Value = "";
                            hoja.Cells("E1").Value = "";
                            hoja.Cells("F1").Value = "RAN " + DateTime.Now.Year;
                            hoja.Cells("G1").Value = "JANVIER";
                            hoja.Cells("H1").Value = "FEVRIER";
                            hoja.Cells("I1").Value = "MARS";
                            hoja.Cells("J1").Value = "AVRIL";
                            hoja.Cells("K1").Value = "MAI";
                            hoja.Cells("L1").Value = "01/01/"+DateTime.Now.Year + " au 31/05/"+DateTime.Now.Year;

                            hoja.Cell("F1").Style.Fill.SetBackgroundColor(XLColor.Yellow);
                            hoja.Cells("G1").Style.Fill.SetBackgroundColor(XLColor.Yellow);
                            hoja.Cells("H1").Style.Fill.SetBackgroundColor(XLColor.Yellow);
                            hoja.Cells("I1").Style.Fill.SetBackgroundColor(XLColor.Yellow);
                            hoja.Cells("J1").Style.Fill.SetBackgroundColor(XLColor.Yellow);
                            hoja.Cells("K1").Style.Fill.SetBackgroundColor(XLColor.Yellow);
                            hoja.Cells("L1").Style.Fill.SetBackgroundColor(XLColor.Yellow);

                            var col1 = hoja.Column(5);
                            col1.Width = 40;
                            for (int i = 6; i < 13; i++)
                            {
                                var col = hoja.Column(i);
                                col.Width = 25;
                            }


                            hoja.Cell(2, 1).Value = "COMPTE";
                            hoja.Cell(2, 2).Value = "R";
                            hoja.Cell(2, 3).Value = "P";
                            hoja.Cell(2, 4).Value = "LIBELLE";
                            hoja.Cell(2, 5).Value = "SOLDE RAN";
                            hoja.Cell(2, 6).Value = "AU 31/01/" + DateTime.Now.Year;
                            hoja.Cell(2, 7).Value = "AU 28/02/" + DateTime.Now.Year;
                            hoja.Cell(2, 8).Value = "AU 31/03/" + DateTime.Now.Year;
                            hoja.Cell(2, 9).Value = "AU 30/04/" + DateTime.Now.Year;

                            hoja.Cell(2, 8).Value = "AU 31/05/" + DateTime.Now.Year;
                            hoja.Cell(2, 9).Value = "SOLDE FINAL";
                            hoja.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                            hoja.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                            double value;
                            foreach (IXLRow row in hoja.Rows())
                            {
                                foreach (IXLCell cell in row.Cells())
                                {
                                    bool successfullyParsed = double.TryParse(cell.Value.ToString(), out value);
                                    if (successfullyParsed)
                                    {
                                        if(double.Parse(cell.Value.ToString())<0)
                                        {
                                            cell.Style.Font.FontColor = XLColor.Red;
                                        }
                                    }
                                }
                            }
                            //hoja.Cell(valuecelltotalHT2).Style
                            //.Border.SetTopBorder(XLBorderStyleValues.Medium)
                            //.Border.SetRightBorder(XLBorderStyleValues.Medium)
                            //.Border.SetBottomBorder(XLBorderStyleValues.Medium)
                            //.Border.SetLeftBorder(XLBorderStyleValues.Medium);
                            //hoja.Cell(valuecelltotalHT1).Style
                            //.Border.SetTopBorder(XLBorderStyleValues.Medium)
                            //.Border.SetRightBorder(XLBorderStyleValues.Medium)
                            //.Border.SetBottomBorder(XLBorderStyleValues.Medium)
                            //.Border.SetLeftBorder(XLBorderStyleValues.Medium);
                            //hoja.Cell(2, 2).Style.Alignment =Alignment. "R";

                            //string valuecelltotalTTC1 = "Q" + (dgv_impression.Rows.Count + 5).ToString();
                            //string valuecelltotalTTC2 = "R" + (dgv_impression.Rows.Count + 5).ToString();
                            //hoja.Cell(valuecelltotalTTC1).Value = " SUM MONTANT TTC";
                            //hoja.Cell(valuecelltotalTTC2).Value = totalTTC;

                            //hoja.Cell(valuecelltotalTTC2).Style.Font.FontSize = 10;
                            //hoja.Cell(valuecelltotalTTC2).Style.Font.FontName = "Arial";

                            //hoja.Cell(valuecelltotalTTC2).Style
                            //.Border.SetTopBorder(XLBorderStyleValues.Medium)
                            //.Border.SetRightBorder(XLBorderStyleValues.Medium)
                            //.Border.SetBottomBorder(XLBorderStyleValues.Medium)
                            //.Border.SetLeftBorder(XLBorderStyleValues.Medium);
                            //hoja.Cell(valuecelltotalTTC1).Style
                            //.Border.SetTopBorder(XLBorderStyleValues.Medium)
                            //.Border.SetRightBorder(XLBorderStyleValues.Medium)
                            //.Border.SetBottomBorder(XLBorderStyleValues.Medium)
                            //.Border.SetLeftBorder(XLBorderStyleValues.Medium);

                            //string valuecelltotalTVA1 = "Q" + (dgv_impression.Rows.Count + 6).ToString();
                            //string valuecelltotalTVA2 = "R" + (dgv_impression.Rows.Count + 6).ToString();
                            //hoja.Cell(valuecelltotalTVA1).Value = " SUM MONTANT TVA";
                            //hoja.Cell(valuecelltotalTVA2).Value = totalTVA;

                            //hoja.Cell(valuecelltotalTVA2).Style.Font.FontSize = 10;
                            //hoja.Cell(valuecelltotalTVA2).Style.Font.FontName = "Arial";
                            //hoja.Cell(valuecelltotalTVA2).Style
                            //.Border.SetTopBorder(XLBorderStyleValues.Medium)
                            //.Border.SetRightBorder(XLBorderStyleValues.Medium)
                            //.Border.SetBottomBorder(XLBorderStyleValues.Medium)
                            //.Border.SetLeftBorder(XLBorderStyleValues.Medium);
                            //hoja.Cell(valuecelltotalTVA1).Style
                            //.Border.SetTopBorder(XLBorderStyleValues.Medium)
                            //.Border.SetRightBorder(XLBorderStyleValues.Medium)
                            //.Border.SetBottomBorder(XLBorderStyleValues.Medium)
                            //.Border.SetLeftBorder(XLBorderStyleValues.Medium);

                            //string valuecelltotallIGNE1 = "Q" + (dgv_impression.Rows.Count + 7).ToString();
                            //string valuecelltotallIGNE2 = "R" + (dgv_impression.Rows.Count + 7).ToString();
                            //hoja.Cell(valuecelltotallIGNE1).Value = " SUM LIGNE";
                            //hoja.Cell(valuecelltotallIGNE2).Value = totallIGNE;

                            //hoja.Cell(valuecelltotallIGNE2).Style.Font.FontSize = 10;
                            //hoja.Cell(valuecelltotallIGNE2).Style.Font.FontName = "Arial";
                            //hoja.Cell(valuecelltotallIGNE2).Style
                            //.Border.SetTopBorder(XLBorderStyleValues.Medium)
                            //.Border.SetRightBorder(XLBorderStyleValues.Medium)
                            //.Border.SetBottomBorder(XLBorderStyleValues.Medium)
                            //.Border.SetLeftBorder(XLBorderStyleValues.Medium);
                            //hoja.Cell(valuecelltotallIGNE1).Style
                            //.Border.SetTopBorder(XLBorderStyleValues.Medium)
                            //.Border.SetRightBorder(XLBorderStyleValues.Medium)
                            //.Border.SetBottomBorder(XLBorderStyleValues.Medium)
                            //.Border.SetLeftBorder(XLBorderStyleValues.Medium);

                            //hoja.Cell(valuecelltotalHT1).Style.Fill.SetBackgroundColor(XLColor.Yellow);
                            //hoja.Cell(valuecelltotalTTC1).Style.Fill.SetBackgroundColor(XLColor.Yellow);
                            //hoja.Cell(valuecelltotalTVA1).Style.Fill.SetBackgroundColor(XLColor.Yellow);
                            //hoja.Cell(valuecelltotallIGNE1).Style.Fill.SetBackgroundColor(XLColor.Yellow);

                            //hoja.Cell(valuecelltotalHT2).Style.Fill.SetBackgroundColor(XLColor.Pink);
                            //hoja.Cell(valuecelltotalTTC2).Style.Fill.SetBackgroundColor(XLColor.Pink);
                            //hoja.Cell(valuecelltotalTVA2).Style.Fill.SetBackgroundColor(XLColor.Pink);
                            //hoja.Cell(valuecelltotallIGNE2).Style.Fill.SetBackgroundColor(XLColor.Pink);

                            //hoja.Columns("Q:Q").Width = 20;
                            //for (int i = 1; i < dgv_impression.Rows.Count + 1; i++)
                            //{
                            //    hoja.Rows($"{i.ToString()}:{i.ToString()}").Height = 17;
                            //}
                            #endregion

                            MessageBox.Show("Data is exported!");
                            workbook.SaveAs(sfd.FileName);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }
        public static System.Data.DataTable DataGridView_To_Datatable(DataGridView dg)
        {
            System.Data.DataTable ExportDataTable = new System.Data.DataTable();
            foreach (DataGridViewColumn col in dg.Columns)
            {
                ExportDataTable.Columns.Add(col.Name);
            }
            foreach (DataGridViewRow row in dg.Rows)
            {
                DataRow dRow = ExportDataTable.NewRow();
                foreach (DataGridViewCell cell in row.Cells)
                {
                    dRow[cell.ColumnIndex] = cell.Value;
                }
                ExportDataTable.Rows.Add(dRow);
            }
            return ExportDataTable;
        }


    }
}
