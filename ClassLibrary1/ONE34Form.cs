using System;
using System.IO;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using CDNBase;
using ExcelDataReader;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;

namespace one034
{
    public partial class ONE034Form : Form
    {
        string connectionString = @"Data Source = SQLSERVER; Initial Catalog = DomarXL; User ID=AdminS;Password=6,eF#W451,t6,D@";
        List<string> listaDuplikatowIDTranzakcji = new List<string>();
        public ONE034Form()
        {
            InitializeComponent();
        }

        void Button1_Click(object sender, EventArgs e)
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
                List<ONE032Data> listaPrzyszlychRekordow = KompletujeListePrzyszlychRekordow(dataTable);

                #endregion

            #endregion

            #region baza sql
            //try
            //{
                foreach (var item in listaPrzyszlychRekordow)
                {
                    string strZwracanyPrzezInsertInto;
                    //moj insert dla elementu listy
                    strZwracanyPrzezInsertInto = InsertInto(//tu moje query
                    "INSERT INTO[dbo].[ONE_platnosci] ([Operator] ,[ID Transakcji] ,[Data wpłaty] ,[Kwota] ,[Odbiorca] ,[Adresat] ,[Miasto] ,[Kod] ,[Adres] ,[Email] ,[ID Paczki] ,[Zamówienie] ,[Dokument rozliczony])" +
                    "VALUES ('"+item.Operator +"' , '"+item.IDTranzakcji + "' , '" + item.DataWyplaty + "' , " + item.Kwota.ToString().Replace(",", ".") + " , '" + item.Odbiorca + "' , '" + item.Adresat + "' , '" + item.Miasto + "' , '" + item.Kod + "' , '" + item.Adres + "' , '" + item.Email + "' , '" + item.IDPaczki + "' , '" + item.Zamowienie + "' , '" + item.DokumentRozliczny +"' )"
                        );

                    if(strZwracanyPrzezInsertInto == "blad")
                    {
                        listaDuplikatowIDTranzakcji.Add(item.IDTranzakcji);
                    }
                }
            //}
            //catch (Exception ee)
            //{

            //    MessageBox.Show("Nie udało się dodać danych do bazy. Blad: "+ee.Message,"Komunikat");
            //}
            if(listaDuplikatowIDTranzakcji.Count != 0 )
            {
                string combinedListofPrimaryKeyDuplicats = string.Join(",\n", listaDuplikatowIDTranzakcji);

                MessageBox.Show("Nie udało się dodać poniższych rekordów. W bazie znajdują się już rekordy o takim ID Tranzakcji: \n"+combinedListofPrimaryKeyDuplicats,"Komunikat", MessageBoxButtons.OK,MessageBoxIcon.Warning);
            }

            #endregion

        }

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
        List<ONE032Data> KompletujeListePrzyszlychRekordow(DataTable _dataTable)
        {
            List<ONE032Data> _listaPrzyszlychRekordow = new List<ONE032Data>();

            string typ = SprawdzamWzorOperatora(_dataTable);


            int numPierwszegoWiersza = 1;
            if (typ == "P24") { numPierwszegoWiersza = 4; }
            for (int i = numPierwszegoWiersza; i < _dataTable.Rows.Count - 1; i++)
            {
                try
                {
                    DodajePozycjeDoListyWgWzoru(_listaPrzyszlychRekordow, typ, i, _dataTable);
                }
                catch (Exception exp)
                {
                    MessageBox.Show("Niepoprawny format zmiennych podanych w pliku arkusza.\n" + exp.Message, "Komunikat",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return null;
                }
            }


            return _listaPrzyszlychRekordow;
        }
        string SprawdzamWzorOperatora(DataTable _dataTable)
        { // zwraca kod operatora do późniejszego wyboru typu zczytywania danych
            string temp = Convert.ToString(_dataTable.Rows[0][0]);
            if (temp == "Lista transakcji wygenerowana z panelu administracyjnego systemu Przelewy24.pl")
            {
                return "P24";
            }
            else if (temp == "Nr_klienta")
            {
                return "DPD";
            }
            else if (temp == "LP")
            {
                return "DTS";
            }
            else
            {
                return null;
            }
        }
        void DodajePozycjeDoListyWgWzoru(List<ONE032Data> _lista, string kodWzoru, int i, DataTable _dataTable)
        {
            if (kodWzoru == "P24")
            {
                string zam;
                string temppp = Convert.ToString(_dataTable.Rows[i][14]);
                if (Convert.ToString(_dataTable.Rows[i][14]).Contains("/"))
                {
                    zam = Convert.ToString(_dataTable.Rows[i][14]).Substring(0, Convert.ToString(_dataTable.Rows[i][14]).IndexOf(" /"));
                }else if (Convert.ToString(_dataTable.Rows[i][14]).Contains("Zamówienie: ")){
                    zam = Convert.ToString(_dataTable.Rows[i][14]).Substring(12);
                }else if (Convert.ToString(_dataTable.Rows[i][14]).Contains("Numer zamowienia: ")){
                    zam = Convert.ToString(_dataTable.Rows[i][14]).Substring(18);
                }else{
                    zam = "Nie dało sie pobrac numeru zamowienia.";
                }
                MessageBox.Show(zam+"\n"+ Convert.ToString(_dataTable.Rows[i][14]));
                _lista.Add(new ONE032Data()
                {
                    Operator = kodWzoru,
                    IDTranzakcji = Convert.ToString(_dataTable.Rows[i][1]),
                    DataWyplaty = Convert.ToString(_dataTable.Rows[i][3]),
                    Kwota = Convert.ToDecimal(_dataTable.Rows[i][4]) / 100,
                    Odbiorca = "",
                    Adresat = Convert.ToString(_dataTable.Rows[i][5]),
                    Miasto = Convert.ToString(_dataTable.Rows[i][7]),
                    Kod = Convert.ToString(_dataTable.Rows[i][8]),
                    Adres = Convert.ToString(_dataTable.Rows[i][9]),
                    Email = Convert.ToString(_dataTable.Rows[i][10]),
                    IDPaczki = Convert.ToString(_dataTable.Rows[i][12]),


                    Zamowienie = zam,
                    DokumentRozliczny = ""


                });
            }
            else if (kodWzoru == "DTS")
            {
                _lista.Add(new ONE032Data()
                {
                    Operator = kodWzoru,
                    IDTranzakcji = Convert.ToString(_dataTable.Rows[i][1]),
                    DataWyplaty = Convert.ToString(_dataTable.Rows[i][6]),
                    Kwota = Convert.ToDecimal(_dataTable.Rows[i][7]),
                    Odbiorca = Convert.ToString(_dataTable.Rows[i][2]),
                    Adresat = Convert.ToString(_dataTable.Rows[i][3]),
                    Miasto = "",
                    Kod = "",
                    Adres = "",
                    Email = "",
                    IDPaczki = Convert.ToString(_dataTable.Rows[i][1]),
                    Zamowienie = Convert.ToString(_dataTable.Rows[i][4]),
                    DokumentRozliczny = ""


                });
            }
            else if (kodWzoru == "DPD")
            {
                _lista.Add(new ONE032Data()
                {
                    Operator = kodWzoru,
                    IDTranzakcji = Convert.ToString(_dataTable.Rows[i][4]),
                    DataWyplaty = Convert.ToString(_dataTable.Rows[i][7]),
                    Kwota = Convert.ToDecimal(_dataTable.Rows[i][6]),
                    Odbiorca = Convert.ToString(_dataTable.Rows[i][3]),
                    Adresat = Convert.ToString(_dataTable.Rows[i][8]),
                    Miasto = Convert.ToString(_dataTable.Rows[i][10]),
                    Kod = Convert.ToString(_dataTable.Rows[i][9]),
                    Adres = Convert.ToString(_dataTable.Rows[i][11]),
                    Email = "",
                    IDPaczki = Convert.ToString(_dataTable.Rows[i][15]),
                    Zamowienie = Convert.ToString(_dataTable.Rows[i][12]).Substring(Convert.ToString(_dataTable.Rows[i][12]).IndexOf(";;") + 3).Substring(0, Convert.ToString(_dataTable.Rows[i][12]).Substring(Convert.ToString(_dataTable.Rows[i][12]).IndexOf(";;") + 3).IndexOf(";")),
                    DokumentRozliczny = ""


                });

            }
            return;
        }
        #endregion

        //zakomentowane
        #region connekcje i logowania
        /*
        public bool GetConfig()
        {
            bool flaga = false;
            cnf = KonfigPobierz.GetData();
            XLUser = cnf.XLLogin_OpeIdent;
            XLDatabase = cnf.XLLogin_Baza;
            XLPass = cnf.XLLogin_OpeHaslo;

            XLWersjaAPI = 20202;
            XLTrybWsadowy = 1;
            XLWinieta = -1;
            XLUtworzWlasnaSesje = 0;

            XLSQLServer = cnf.XLServer;
            XLSQLDatabase = cnf.XLDatabase;

            if (XLUser != "" && XLDatabase != "" && XLPass != "") ;                                                              //tu nie ma żadnej funkcjonalności
            flaga = true;

            return flaga;
        }

        public int XLLogin()
        {
            return cdn_api.cdn_api.XLLogin(new XLLoginInfo_20202()
            {
                Wersja = XLWersjaAPI,
                ProgramID = "ONE0011",
                //PlikLog = (Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\api.log"),
                PlikLog = ("C:\\OneSupport\\ONE011\\api.log"),
                Baza = XLDatabase,
                OpeIdent = XLUser,
                OpeHaslo = XLPass,
                TrybWsadowy = XLTrybWsadowy,
                // TrybNaprawy = 0,
                UtworzWlasnaSesje = XLUtworzWlasnaSesje,
                Winieta = XLWinieta,
            }, ref sesja);
        }

        public void Logout()
        {
            cdn_api.cdn_api.XLLogout(sesja);
        }

        public string GetConnectionString()
        {
            XLPolaczenieInfo_20202 polaczenieInfo = new XLPolaczenieInfo_20202 { Wersja = 20202 };
            cdn_api.cdn_api.XLPolaczenie(polaczenieInfo);

            connection_string = polaczenieInfo.ConnectString + "UID=sabisup;PWD=#0ut2510gro0ne.34;";

            return connection_string;
        }
        */
        #endregion connekcje i logowania

        //zakomentowane
        #region funkcje sql


        public string InsertInto(string sql)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    connection.Open();
                    adapter.InsertCommand = new SqlCommand(sql, connection);
                    adapter.InsertCommand.ExecuteNonQuery();
                    //MessageBox.Show("Row inserted !! ");
                }
            }
            catch (Exception ex) when  (ex.Message.Substring(0, 1) == "V")
            {
                //listaDuplikatowIDTranzakcji.Add(item.IDTranzakcji);
                return "blad";
            }
            return sql;
        }
        /*
        public string Update(string sql)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(connection_string))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    connection.Open();
                    adapter.UpdateCommand = new SqlCommand(sql, connection);
                    adapter.UpdateCommand.ExecuteNonQuery();
                    //MessageBox.Show("Row updated !! ");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return sql;
        }

        public string InsertInto(string sql)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(connection_string))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    connection.Open();
                    adapter.InsertCommand = new SqlCommand(sql, connection);
                    adapter.InsertCommand.ExecuteNonQuery();
                    //MessageBox.Show("Row inserted !! ");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return sql;
        }

        public string DeleteFrom(string sql)       //Użycie delete i inserta nie jest optymalne, ale w niektórych przypadkach pozycje nie miały atrybutu. Po testach zmienić na update
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(connection_string))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    connection.Open();
                    adapter.DeleteCommand = new SqlCommand(sql, connection);
                    adapter.DeleteCommand.ExecuteNonQuery();
                    //MessageBox.Show("Row deleted!! ");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return sql;
        }

        public List<string> SelectList(string query, string szukana, string srcTableName)
        {
            List<string> str = new List<string>();
            string sttr = "";
            try
            {
                using (SqlConnection connection = new SqlConnection(connection_string))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    adapter.SelectCommand = new SqlCommand(query, connection);
                    DataSet dataSet = new DataSet();
                    adapter.Fill(dataSet, srcTableName);
                    foreach (DataRow row in (InternalDataCollectionBase)dataSet.Tables[0].Rows)
                        str.Add(row[szukana].ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(sttr = sttr + "Bład SQL:" + ex.ToString() + "\n");
            }
            return str;
        }

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
        }

        public string Select(string query, string szukana, string srcTableName)
        {
            string str = "";
            try
            {
                using (SqlConnection connection = new SqlConnection(connection_string))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    adapter.SelectCommand = new SqlCommand(query, connection);
                    DataSet dataSet = new DataSet();
                    adapter.Fill(dataSet, srcTableName);
                    foreach (DataRow row in (InternalDataCollectionBase)dataSet.Tables[0].Rows)
                        str = row[szukana].ToString();
                }
            }
            catch (Exception ex)
            {
                str = str + "Bład SQL:" + ex.ToString() + "\n";
            }
            return str;
        }
        */
        #endregion funkcje sql
    }
}
