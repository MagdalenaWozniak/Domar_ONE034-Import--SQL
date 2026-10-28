using Amazon.Runtime.Internal.Util;
using cdn_api;
using Hydra;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace one034
{
    public partial class FormWybor : Form
    {
        long KRPNumer;
        public FormWybor(long sValue)
        {
            InitializeComponent();
            KRPNumer = sValue;
        }

        private void buttonZapisy_Click(object sender, EventArgs e)
        {
            string baza = "";
            string connectionString = "";
            //int XLWersjaAPI = 20202;
            //int XLTrybWsadowy = 1;
           // int XLWinieta = -1;
            //int XLUtworzWlasnaSesje = 1;
            //int sesja = -1;
            string XLSQLServer = "";
            string XLSQLDatabase = "";//"ERPXL_PoloMarket"
            string XLUser = ""; //
            string XLPass = "";
            if (Runtime.ConfigurationDictionary.Baza.ToString().Substring(0, 6) == "Kopia2")
            {
                baza = "Kopia2";
                connectionString = @"Data Source=192.168.40.16;Initial Catalog=Kopia2; User ID=AdminS;Password=6,eF#W451,t6,D@";
            }
            else if (Runtime.ConfigurationDictionary.Baza.ToString().Substring(0, 7) == "KopiaXL")
            {
                baza = "KopiaXL";
                connectionString = @"Data Source=192.168.40.16;Initial Catalog=KopiaXL; User ID=AdminS;Password=6,eF#W451,t6,D@";
            }
            else if (Runtime.ConfigurationDictionary.Baza.ToString().Substring(0, 7) == "DomarXL")
            {
                baza = "DomarXL";
                connectionString = @"Data Source=192.168.40.16;Initial Catalog=DomarXL; User ID=AdminS;Password=6,eF#W451,t6,D@";
            }
            else if (Runtime.ConfigurationDictionary.Baza.ToString().Substring(0, 5) == "DOMAR")
            {
                baza = "DOMAR";
                connectionString = @"Data Source=SRV1;Initial Catalog=DOMAR; User ID=admin;Password=Dupa2003!";
                XLSQLServer = "192.168.137.149";
                XLSQLDatabase = "Domar2";//"ERPXL_PoloMarket"
                XLUser = "OS"; //
                XLPass = "NE!w0rk12!";
                
            }
            else
            {
                MessageBox.Show("Dodatek nie posiada funkcjonalności dla tej bazy.", "Komunikat",
                 MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            if (baza != "" && connectionString != "")
            {
                //logowanie(XLSQLServer, XLSQLDatabase, XLUser, XLPass);
                //cdn_api.cdn_api.XLPobierzNumerDokumentu();
                rozliczanieNajmu(baza, connectionString);
            }

        }

        private void buttonExcel_Click(object sender, EventArgs e)
        {

            if (Runtime.ConfigurationDictionary.Baza.ToString().Substring(0, 7) == "KopiaXL")
            {
                MessageBox.Show("Po wybraniu pliku nastąpi dodawanie danych do bazy KopiaXL", "Komunikat",
                 MessageBoxButtons.OK, MessageBoxIcon.Information);
                Process.Start("C:\\OneSupport\\ImportPlatnosci\\KopiaXL\\ONE034.exe");
            }
            else if (Runtime.ConfigurationDictionary.Baza.ToString().Substring(0, 7) == "DomarXL")
            {
                MessageBox.Show("Po wybraniu pliku nastąpi dodawanie danych do bazy DomarXL", "Komunikat",
                 MessageBoxButtons.OK, MessageBoxIcon.Information);
                Process.Start("C:\\OneSupport\\ImportPlatnosci\\DomarXL\\ONE034.exe");
            }
            else if (Runtime.ConfigurationDictionary.Baza.ToString().Substring(0, 5) == "DOMAR")
            {
                MessageBox.Show("Po wybraniu pliku nastąpi dodawanie danych do bazy DOMAR", "Komunikat",
                 MessageBoxButtons.OK, MessageBoxIcon.Information);
                Process.Start("C:\\Users\\magdaw\\Documents\\Domar\\Excel-Halmar\\one034 form — kopiaDziala\\build one34\\ONE034.exe");
            }
            else
            {
                MessageBox.Show("Dodatek nie posiada funkcjonalności dla tej bazy.", "Komunikat",
                 MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        void rozliczanieNajmu(string baza, string connectionString)
        {
            
            string countSelect = "Select count(*) FROM [" + baza + "].[CDN].[Zapisy] WHERE KAZ_KRPNumer="+KRPNumer+" and KAZ_CiagNumeracji='BP'and KAZ_Rozliczony=0 and (KAZ_Tresc like '%/N1%' or KAZ_Tresc like'%/N2%' or  KAZ_Tresc like '%_N1%' or KAZ_Tresc like'%_N2%')";
            int count = Convert.ToInt32(ConnectToSQL(connectionString, countSelect));

            if (count == 0)
            {
                MessageBox.Show("Brak nierozliczonych przychodów za najem w tym rejestrze", "Komunikat",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                List<ONE032Data> _lista = new List<ONE032Data>();
                string TypString = "Select [KAZ_GIDTyp] FROM [" + baza + "].[CDN].[Zapisy] WHERE KAZ_KRPNumer=" + KRPNumer + " and KAZ_CiagNumeracji='BP'and KAZ_Rozliczony=0 and (KAZ_Tresc like '%/N1%' or KAZ_Tresc like'%/N2%' or  KAZ_Tresc like '%_N1%' or KAZ_Tresc like'%_N2%')";
                List<string> typLista = new List<string>();
                typLista = SelectList(TypString, "Zapisy", connectionString);
                string FirmaString = "Select [KAZ_GIDFirma] FROM [" + baza + "].[CDN].[Zapisy] WHERE KAZ_KRPNumer=" + KRPNumer + " and KAZ_CiagNumeracji='BP'and KAZ_Rozliczony=0 and (KAZ_Tresc like '%/N1%' or KAZ_Tresc like'%/N2%' or  KAZ_Tresc like '%_N1%' or KAZ_Tresc like'%_N2%')";
                List<string> firmaLista = new List<string>();
                firmaLista = SelectList(FirmaString, "Zapisy", connectionString);
                string NumerString = "Select  [KAZ_GIDNumer] FROM [" + baza + "].[CDN].[Zapisy] WHERE KAZ_KRPNumer=" + KRPNumer + " and KAZ_CiagNumeracji='BP'and KAZ_Rozliczony=0 and (KAZ_Tresc like '%/N1%' or KAZ_Tresc like'%/N2%' or  KAZ_Tresc like '%_N1%' or KAZ_Tresc like'%_N2%')";
                List<string> numerLista = new List<string>();
                numerLista = SelectList(NumerString, "Zapisy", connectionString);
                string KwotaString = "Select  [KAZ_Kwota] FROM [" + baza + "].[CDN].[Zapisy] WHERE KAZ_KRPNumer=" + KRPNumer + " and KAZ_CiagNumeracji='BP'and KAZ_Rozliczony=0 and (KAZ_Tresc like '%/N1%' or KAZ_Tresc like'%/N2%' or  KAZ_Tresc like '%_N1%' or KAZ_Tresc like'%_N2%')";
                List<string> kwotaLista = new List<string>();
                kwotaLista = SelectList(KwotaString, "Zapisy", connectionString);
                string FakturaString = "Select [KAZ_Tresc] FROM [" + baza + "].[CDN].[Zapisy] WHERE KAZ_KRPNumer=" + KRPNumer + " and KAZ_CiagNumeracji='BP'and KAZ_Rozliczony=0 and (KAZ_Tresc like '%/N1%' or KAZ_Tresc like'%/N2%' or  KAZ_Tresc like '%_N1%' or KAZ_Tresc like'%_N2%')";
                List<string> fakturaLista = new List<string>();
                fakturaLista = SelectList(FakturaString, "Zapisy", connectionString);
                for (int i = 0; i < count; i++)
                {
                    fakturaLista[i] = fakturaLista[i].Replace("_", "/").Replace(" ", "");
                    string fakturaTemp = fakturaLista[i];
                    try
                    {
                        
                        int countFaktura = fakturaTemp.Length - fakturaTemp.Replace("FS-", "FS").Length;
                        //MessageBox.Show("W treści: " + fakturaLista[i] + " są " + countFaktura + " faktury", "Komunikat",
                        // MessageBoxButtons.OK, MessageBoxIcon.Information);
                        if (countFaktura == 1)
                        {
                            string koniec = "";
                            if (fakturaLista[i].Contains("/N1"))
                                koniec = "/N1";
                            else
                                koniec = "/N2";

                                //string Faktura = ConnectToSQL(connectionString, FakturaString).Replace("_", "/");
                                //Faktura = Faktura.Replace(" ", "");
                                _lista.Add(new ONE032Data()
                                {
                                    Operator = "Najemca",
                                    IDTranzakcji = numerLista[i],
                                    DataWyplaty = "",
                                    // DataWyplaty = Convert.ToString(_dataTable.Rows[i][6]).Substring(6, 4) + "-" + Convert.ToString(_dataTable.Rows[i][6]).Substring( 3, 2) + "-" + Convert.ToString(_dataTable.Rows[i][6]).Substring(0, 2),
                                    Kwota = Convert.ToDecimal(kwotaLista[i]),
                                    Odbiorca = firmaLista[i],
                                    Adresat = "",
                                    Miasto = "",
                                    Kod = "",
                                    Adres = "",
                                    Email = "",
                                    IDPaczki = typLista[i],
                                    Zamowienie = fakturaLista[i].Substring(fakturaLista[i].IndexOf("FS-"), fakturaLista[i].IndexOf(koniec) + 3 - fakturaLista[i].IndexOf("FS-")),
                                    DokumentRozliczny = ""


                                });
                        }
                        else
                        {
                            List<string> faktury = new List<string>();
                            fakturaTemp = fakturaLista[i];
                            for (int j = 0; j < countFaktura; j++)
                            {
                                
                                string faktura = fakturaTemp.Substring(fakturaTemp.IndexOf("FS-"), fakturaTemp.IndexOf("/N") + 3 - fakturaTemp.IndexOf("FS-"));
                                MessageBox.Show("Faktura nr  " + Convert.ToString(j + 1) + ": " + faktura, "Komunikat",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                                if (fakturaTemp.IndexOf("FS-") == -1)
                                {
                                    // MessageBox.Show("Indeks faktury nr  " + Convert.ToString(j + 1) + ": " + Convert.ToString(fakturaTemp.IndexOf("FS-")), "Komunikat",
                                    //  MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    break;
                                }
                                if (!faktury.Contains(faktura))
                                {
                                    faktury.Add(faktura);
                                    // MessageBox.Show("Dodano fakturę nr  " + Convert.ToString(j + 1) + ": " + faktura, "Komunikat",
                                    //MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }

                                fakturaTemp = fakturaTemp.Remove(fakturaTemp.IndexOf(faktura), fakturaTemp.IndexOf("/N") + 3 - fakturaTemp.IndexOf(faktura));
                                // MessageBox.Show("Faktura po usunięciu:  " + fakturaTemp, "Komunikat",
                                // MessageBoxButtons.OK, MessageBoxIcon.Information);

                            }
                            //MessageBox.Show(Convert.ToString(faktury.Count), "Komunikat",
                            //  MessageBoxButtons.OK, MessageBoxIcon.Information);
                            int k = 0;
                            foreach (var item in faktury)
                            {
                                _lista.Add(new ONE032Data()
                                {
                                    Operator = "Najemca",
                                    IDTranzakcji = numerLista[i] + Convert.ToString(k),
                                    DataWyplaty = "",
                                    // DataWyplaty = Convert.ToString(_dataTable.Rows[i][6]).Substring(6, 4) + "-" + Convert.ToString(_dataTable.Rows[i][6]).Substring( 3, 2) + "-" + Convert.ToString(_dataTable.Rows[i][6]).Substring(0, 2),
                                    Kwota = Convert.ToDecimal(kwotaLista[i]),
                                    Odbiorca = firmaLista[i],
                                    Adresat = "",
                                    Miasto = "",
                                    Kod = "",
                                    Adres = "",
                                    Email = "",
                                    IDPaczki = typLista[i],
                                    Zamowienie = item,
                                    DokumentRozliczny = ""
                                });
                                k++;
                            }
                        }
                    }
                    catch(Exception e)
                    {
                        MessageBox.Show("Błąd dla "+fakturaTemp+": "+ e.Message+ " "+e.StackTrace, "Komunikat",
                              MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }


                    foreach (var item in _lista)
                    {
                        string query1 = "SELECT [Rozliczono] FROM [CDN].[ONE_Platnosci] WHERE Operator='" + item.Operator + "' and IDTransakcji='" + item.IDTranzakcji + "'";
                        string query2 = "SELECT [Status] FROM [CDN].[ONE_Platnosci] WHERE Operator = '" + item.Operator + "' and IDTransakcji = '" + item.IDTranzakcji + "'";
                        string strZwracanyPrzezInsertInto;
                        string op = item.Operator;
                        //moj insert dla elementu listy
                        strZwracanyPrzezInsertInto = InsertInto(//tu moje query
                        "INSERT INTO[CDN].[ONE_Platnosci] ([Operator],[IDTransakcji],[Datawplaty],[Kwota],[Odbiorca],[Adresat],[Miasto],[Kod],[Adres],[Email],[IDPaczki],[Zamowienie],[Rozliczono])" +
                        "VALUES ('" + item.Operator + "' , '" + item.IDTranzakcji + "' , '" + item.DataWyplaty + "' , " + item.Kwota.ToString().Replace(",", ".") + " , '" + item.Odbiorca + "' , '" + item.Adresat + "' , '" + item.Miasto + "' , '" + item.Kod + "' , '" + item.Adres + "' , '" + item.Email + "' , '" + item.IDPaczki + "' , '" + item.Zamowienie + "' , '" + item.DokumentRozliczny + "' )", connectionString
                            );

                        if (strZwracanyPrzezInsertInto == "blad")
                        {
                            string rozliczono = ConnectToSQL(connectionString, query1);
                            if (rozliczono == "False")
                            {
                            string update = "UPDATE [CDN].[ONE_Platnosci] SET [Operator] = '"+item.Operator+"' , [Datawplaty]= '"+ item.DataWyplaty + "' , [Kwota]=" + item.Kwota.ToString().Replace(",", ".") +" , [Odbiorca]='"+ item.Odbiorca + "' , [Adresat]='" + item.Adresat + "' , [Miasto]= '" + item.Miasto + "' , [Kod]='" + item.Kod + "' , [Adres]='" + item.Adres + "' , [Email]='" + item.Email + "' , [IDPaczki]='" + item.IDPaczki + "' , [Zamowienie]='" + item.Zamowienie + "' , [Rozliczono] = '" + item.DokumentRozliczny + "' WHERE [IDTransakcji] = '" + item.IDTranzakcji + "'";
                            Update(update, connectionString);
                            procedura(connectionString, item.Operator, item.IDPaczki, item.IDTranzakcji);

                            MessageBox.Show("Transakcja " + item.IDTranzakcji + " Dla operatora " + item.Operator + " została wcześniej pobrana. Rekord został zaktualizowany.", "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                             
                            //listaDuplikatowIDTranzakcji.Add(item.IDTranzakcji);
                            }
                            else
                            {
                                string status = ConnectToSQL(connectionString, query2);
                                MessageBox.Show("Transakcja " + item.IDTranzakcji + " Dla operatora " + item.Operator + " Jest już rozliczona w systemie " + status, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                        else
                        {
                           procedura(connectionString, item.Operator, item.IDPaczki, item.IDTranzakcji);
                        }
                    }
                }
                MessageBox.Show("Zakończono rozliczanie przychodów za najem", "Komunikat",
               MessageBoxButtons.OK, MessageBoxIcon.Information);
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

            public List<string> SelectList(string query, string srcTableName, string connection_string)
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
                        {
                            str.Add(row[0].ToString());
                            // MessageBox.Show(row[0].ToString(), "Komunikat",
                            //  MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }

                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(sttr = sttr + "Bład SQL:" + ex.ToString() + "\n");
                }
                return str;
            }
        public string InsertInto(string sql, string connectionString)
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
            catch (Exception ex) when (ex.Message.Substring(0, 1) == "V")
            {
                //listaDuplikatowIDTranzakcji.Add(item.IDTranzakcji);
                return "blad";
            }
            return sql;
        }

        public string Update(string sql, string connection_string)
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

        void procedura(string connectionString, string OP, string idPaczki, string IDTran)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();

                    // 1.  create a command object identifying the stored procedure
                    SqlCommand cmd = new SqlCommand("cdn.ONE_RozliczPlatnosci", conn);

                    // 2. set the command object so it knows to execute a stored procedure
                    cmd.CommandType = CommandType.StoredProcedure;

                    // 3. add parameter to command, which will be passed to the stored procedure
                    cmd.Parameters.Add(new SqlParameter("@Operator", OP));
                    cmd.Parameters.Add(new SqlParameter("@idPaczki", idPaczki));
                    cmd.Parameters.Add(new SqlParameter("@idTransakcji", IDTran));

                    // execute the command
                    using (SqlDataReader rdr = cmd.ExecuteReader())
                    {
                        // iterate through results, printing each to console
                        while (rdr.Read())
                        {
                            //Console.WriteLine("Product: {0,-35} Total: {1,2}", rdr["ProductName"], rdr["Total"]);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Złe wywołanie procedury: " + ex.Message, "Komunikat",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        void logowanie(string XLSQLServer, string XLSQLDatabase, string XLUser, string XLPass)
        {
            XLLoginInfo_20202 nowyUzytkownik = new XLLoginInfo_20202();
            nowyUzytkownik.Wersja = 20202;
            nowyUzytkownik.Winieta = 1;
            nowyUzytkownik.TrybWsadowy = 1;
            nowyUzytkownik.UtworzWlasnaSesje = 1;
            nowyUzytkownik.SerwerKlucza = XLSQLServer;
            nowyUzytkownik.Baza = XLSQLDatabase;
            nowyUzytkownik.OpeIdent = XLUser;
            nowyUzytkownik.OpeHaslo = XLPass;
            //nowyUzytkownik.PortHTTPSerweraKlucza = 1;
            //nowyUzytkownik.RodzajSerweraKlucza = 0;
           // nowyUzytkownik.SerwerKlucza = "SRV1::5000139268";
            nowyUzytkownik.ProgramID = "ONES034 - ImportPlatnosci";

            int sesja = 1;
            int blad =cdn_api.cdn_api.XLLogin(nowyUzytkownik, ref sesja);
            if(blad!=0)
            {
                MessageBox.Show("Błąd logowania: "+blad, "Komunikat",
                   MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
    }

    


