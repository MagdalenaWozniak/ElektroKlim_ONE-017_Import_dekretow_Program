using cdn_api;
using ExcelDataReader;
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

namespace ImportDekretow
{
    public partial class Form1 : Form
    {
        private static int sesja = 0;
        private static int XLWersjaAPI = 20220;
        private static int XLTrybWsadowy = 1;
        private static int XLWinieta = -1;
        private static int XLUtworzWlasnaSesje = 1;
        private static string XLSQLServer = "172.51.1.8";
        private static string XLSQLDatabase = "ElektroKlim_Pro";//"ERPXL_PoloMarket"
        private static string XLUser = "ADMIN"; //
        private static string XLPass = "NE!w0rk12";//"Dupa2003!";
        private static int dokID = 0;
        string connectionString = @"Data Source = EKSRV04\COMARCH; Initial Catalog = ERPXL_E_P; User ID = ComarchAPI; Password=hT@S+NQZBT";
        string date1 = "1800-12-28";

        //@"Data Source = SQLSERVER; Initial Catalog = DomarXL; User ID=AdminS;Password=6,eF#W451,t6,D@";

        public Form1()
        {
            //this.connectionString = GetConnectionString();
            InitializeComponent();
        }


        void Button1_Click(object sender, EventArgs e)
        {
            try
            {
                #region excel5
                #region wybranie pliku
                Cursor.Current = Cursors.WaitCursor;
                if (tB1.Text == "nie wybrano pliku")
                {
                    MessageBox.Show("Nie wybrano pliku arkusza.", "Komunikat",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                string filePath = tB1.Text;
                DataTable dataTable;
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    IExcelDataReader reader;
                    reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream);

                    var conf = new ExcelDataSetConfiguration
                    {
                        ConfigureDataTable = _ => new ExcelDataTableConfiguration
                        {
                            UseHeaderRow = false//true
                        }
                    };

                    var dataSet = reader.AsDataSet(conf);
                    dataTable = dataSet.Tables[0];
                }
                #endregion

                #region Przypisywanie danych z datasetu do listy przyszłych pozycji.

                SprawdzamCzyArkuszNieJestPusty(dataTable);
                dodajDekrety(dataTable);

                #endregion

                #endregion
            }
            catch(Exception en)
            {
                MessageBox.Show("Błąd pobierania danych z excel: " + en.Message, "Komunikat",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }


        }
        #region Dodawanie dekretów
        void dodajDekrety(DataTable dataTable)
        {
            bool jest = true;

            for (int k = 1; k < dataTable.Rows.Count; k++)
            {

                string konto = Convert.ToString(dataTable.Rows[k][2]).Trim();
                if (konto.Length == 3)
                    konto += "-01";
                else if (konto.Length == 4)
                    konto += "01";

                string jestQuery = "if exists (SELECT [KKS_Konto] FROM [CDN].[Konta] WHERE KKS_Konto = '"+konto+"') SELECT 'TAK' else SELECT 'NIE'";
                string istnieje = ConnectToSQL(connectionString, jestQuery);

                if(istnieje=="NIE")
                {
                    jest = false;
                    MessageBox.Show("Brak konta: " + konto, "Komunikat",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    break;
                    
                }
            }

            if (jest)
            {

                XLLoginInfo_20220 nowyUzytkownik = new XLLoginInfo_20220();
                nowyUzytkownik.Wersja = XLWersjaAPI;
                nowyUzytkownik.Winieta = XLWinieta;
                nowyUzytkownik.TrybWsadowy = XLTrybWsadowy;
                nowyUzytkownik.UtworzWlasnaSesje = XLUtworzWlasnaSesje;
                nowyUzytkownik.SerwerKlucza = XLSQLServer;
                nowyUzytkownik.Baza = XLSQLDatabase;
                nowyUzytkownik.OpeIdent = XLUser;
                nowyUzytkownik.OpeHaslo = XLPass;
                nowyUzytkownik.PortHTTPSerweraKlucza = 1;
                nowyUzytkownik.RodzajSerweraKlucza = 0;
                nowyUzytkownik.SerwerKlucza = @"EKSRV04\COMARCH::5000216950";
                nowyUzytkownik.ProgramID = "ONE017 - Dodawanie Dekretów";

                cdn_api.cdn_api.XLLogin(nowyUzytkownik, ref sesja);

                int i = 1;


                string akronim = Convert.ToString(dataTable.Rows[i][3]).ToUpper();
                int count = 0;
                int count2 = 0;

                string nazwa = znajdzNazwe(tB1.Text);


                string date2 = "";
                if (Convert.ToString(dataTable.Rows[i][1]).IndexOf(".") == 1)
                    date2 = Convert.ToString(dataTable.Rows[i][1]).Substring(5, 4) + "-" + Convert.ToString(dataTable.Rows[i][1]).Substring(2, 2) + "-0" + Convert.ToString(dataTable.Rows[i][1]).Substring(0, 1);
                else
                    date2 = Convert.ToString(dataTable.Rows[i][1]).Substring(6, 4) + "-" + Convert.ToString(dataTable.Rows[i][1]).Substring(3, 2) + "-" + Convert.ToString(dataTable.Rows[i][1]).Substring(0, 2);
                XLDekretZbiorczyInfo_20220 zbiorczy = new XLDekretZbiorczyInfo_20220();
                zbiorczy.Bufor = "BUFOR";
                zbiorczy.Data = Convert.ToInt32(ConnectToSQL(connectionString, "SELECT DATEDIFF(day, '" + date1 + "','" + date2 + "'); "));
                zbiorczy.DataDokumentu = Convert.ToInt32(ConnectToSQL(connectionString, "SELECT DATEDIFF(day, '" + date1 + "','" + date2 + "'); "));
                zbiorczy.Dziennik = "PK";
                zbiorczy.GIDFirma = 0;
                zbiorczy.GIDLp = 0;
                zbiorczy.GIDNumer = 0;
                zbiorczy.GIDTyp = 0;
                zbiorczy.Opis = nazwa;
                zbiorczy.Tryb = 2;
                zbiorczy.Wersja = XLWersjaAPI;

                int blad = cdn_api.cdn_api.XLDodajDekretZbiorczy(sesja, zbiorczy);
                if (blad != 0)
                    MessageBox.Show("Błąd dodawania dekretu zbiorczego:" + blad, "Komunikat",
                         MessageBoxButtons.OK, MessageBoxIcon.Warning);


                pBar1.Visible = true;
                // Set Minimum to 1 to represent the first file being copied.
                pBar1.Minimum = 1;
                // Set Maximum to the total number of files to copy.
                pBar1.Maximum = dataTable.Rows.Count;
                // Set the initial value of the ProgressBar.
                pBar1.Value = 1;
                // Set the Step property to a value of 1 to represent each file being copied.
                pBar1.Step = 1;


                for (i = 1; i < dataTable.Rows.Count; i++)
                {

                    pBar1.PerformStep();

                    try
                    {

                        XLDekretInfo_20220 dekret = new XLDekretInfo_20220();
                        dekret.GIDFirma = zbiorczy.GIDFirma;
                        dekret.GIDLp = zbiorczy.GIDLp;
                        dekret.GIDNumer = zbiorczy.GIDNumer;
                        dekret.GIDTyp = zbiorczy.GIDTyp;
                        dekret.Tryb = 2;
                        dekret.Wersja = XLWersjaAPI;
                        dekret.Opis = Convert.ToString(dataTable.Rows[i][5]);


                        // MessageBox.Show("Kwota: " + dataTable.Rows[i][4], "Komunikat",
                        // MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        //MessageBox.Show("Kwota2: " + dataTable.Rows[i][3], "Komunikat",
                        // MessageBoxButtons.OK, MessageBoxIcon.Warning);

                        //MessageBox.Show("Konto: " +(dataTable.Rows[i][2]), "Komunikat",
                        // MessageBoxButtons.OK, MessageBoxIcon.Warning);

                        string konto = Convert.ToString(dataTable.Rows[i][2]).Trim();
                        if (konto.Length == 3)
                            konto += "-01";
                        else if (konto.Length == 4)
                            konto += "01";
                        if (dataTable.Rows[i][3].ToString().Trim() == "0")
                        {
                            //MessageBox.Show("Konto winien: " + Convert.ToString(dataTable.Rows[i][2]), "Komunikat",
                            // MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            dekret.KontoCredit = konto;
                            dekret.KontoDebet = "";
                            // MessageBox.Show("Kwota: " + Convert.ToString(dataTable.Rows[i][4]), "Komunikat",
                            //MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            dekret.Kwota = dataTable.Rows[i][4].ToString();//.Replace(",", ".");
                        }
                        else
                        {

                            dekret.KontoCredit = "";
                            //MessageBox.Show("Konto ma: " + Convert.ToString(dataTable.Rows[i][2]), "Komunikat",
                            //MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            dekret.KontoDebet = konto;
                            // MessageBox.Show("Kwota: " + Convert.ToString(dataTable.Rows[i][3]), "Komunikat",
                            //MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            dekret.Kwota = dataTable.Rows[i][3].ToString();//.Replace(",", ".");
                        }

                        blad = cdn_api.cdn_api.XLDodajDekretSkladowy(sesja, dekret);
                        if (blad != 0)
                        {

                            MessageBox.Show("Błąd dodawania dekretu składowego " + Convert.ToString(dataTable.Rows[i][5]) + " :" + blad, "Komunikat",
                             MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Błąd dodawania dekretu składowego " + Convert.ToString(dataTable.Rows[i][5]) + " :" + ex.Message, "Komunikat",
                                 MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }



                }

                blad = cdn_api.cdn_api.XLZamknijDekretZbiorczy(sesja, zbiorczy);
                if (blad != 0)
                    MessageBox.Show("Błąd zamykania dekretu zbiorczego :" + blad, "Komunikat",
                         MessageBoxButtons.OK, MessageBoxIcon.Warning);

                pBar1.Value = 1;
                pBar1.Visible = false;

                string numerQuery = "SELECT [DZK_Identyfikator] FROM [CDN].[Dziennik] WHERE [DZK_GIDTyp] = " + zbiorczy.GIDTyp + " and [DZK_GIDFirma]=" + zbiorczy.GIDFirma + " and [DZK_GIDNumer]=" + zbiorczy.GIDNumer + " and [DZK_GIDLp]=" + zbiorczy.GIDLp;
                string numer = ConnectToSQL(connectionString, numerQuery);

                cdn_api.cdn_api.XLLogout(sesja);
                MessageBox.Show("Dodano dekret zbiorczy: " + numer, "Komunikat",
                         MessageBoxButtons.OK, MessageBoxIcon.Information);
            }


        }


        public string ConnectToSQL(string connection_string, string query)
        {
            string str = "";
            try
            {
                using (SqlConnection connection = new SqlConnection(connection_string))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    adapter.SelectCommand = new SqlCommand(query, connection);
                    DataSet dataSet = new DataSet();
                    adapter.Fill(dataSet, "Query");
                    if (dataSet.Tables.Count == 0)
                    {
                        return "";
                    }
                    foreach (DataRow row in (InternalDataCollectionBase)dataSet.Tables[0].Rows)
                    {
                        str = row[0].ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                str = str + "Bład SQL:" + ex.ToString() + "\n";
            }
            return str;
        }
        #endregion

        #region funkcje wybranie pliku
        void WybraniePliku()
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog()
            {
                InitialDirectory = Directory.GetCurrentDirectory(),
                DefaultExt = "xlsx",
                Filter = "Pliki xlsx (*.xlsx)|*.xlsx|Pliki xls (*.xls)|*.xls|Wszystkie pliki (*.*)|*.*",
            };
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                tB1.Text = openFileDialog1.FileName;
                //MessageBox.Show(path);
            }
            return;
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            WybraniePliku();
        }
        private void TextBox1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            WybraniePliku();
        }
        #endregion

        #region funkcje pobranie danych z pliku
        bool SprawdzamCzyArkuszNieJestPusty(DataTable _dataTable)
        {
            if (_dataTable.Rows.Count == 0)
            {
                MessageBox.Show("Wybrany arkusz pusty.", "Komunikat",
                MessageBoxButtons.OK, MessageBoxIcon.Error);//błąd
                return true;
            }
            return false;
        }


        string znajdzNazwe(string path)
        {
            string[] nazwy = path.Split('\\');
            string nazwa = nazwy[nazwy.Length - 1];
            return nazwa.Substring(0, nazwa.IndexOf(".xls"));
        }

        private void title1_Click(object sender, EventArgs e)
        {

        }

        private void ONE034Form_Load(object sender, EventArgs e)
        {

        }

        /*
        public bool CzyTraNag(short typ)        //zwraca informacje, czy zaznaczony dokument będzie się znajdował w tabeli TraNag
        {
            if (typ == 1529
            || typ == 434
            || typ == 800
            || typ == 1521
            || typ == 1529
            || typ == 1824
            || typ == 1828
            || typ == 1832
            || typ == 1836
            || typ == 2033
            || typ == 2034
            || typ == 2037
            || typ == 2041
            || typ == 2042
            || typ == 2045
            || typ == 2768
            || typ == 2832
            || typ == 2976
            || typ == 2977
            || typ == 3344
            || typ == 3352
            || typ == 4145
            || typ == 7680
            || typ == 7684)
                return true;
            else return false;
        }*/

        #endregion
    }
}