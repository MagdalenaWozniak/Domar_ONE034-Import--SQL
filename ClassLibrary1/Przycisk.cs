using cdn_api;
using Hydra;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;


[assembly: CallbackAssemblyDescription(Name: "ONE034 - Import -> SQL",
                                           Description: "Importuje pliki excel do tabeli płatności SQL.",
                                           Author: "OneSupport",
                                           Version: "1.0",
                                           SystemVersion: "2020.2.2",
                                           Created: "01-05-2021")]

namespace one034
{
    [SubscribeProcedure(Procedures.KAZLista, "ONE034 - Import -> SQL")]
    public class importDoSql : Callback
    {
        [DllImport("user32.dll", EntryPoint = "keybd_event", CharSet = CharSet.Auto, ExactSpelling = true)]
        private static extern void keybd_event(byte vk, byte scan, int flags, int extrainfo);
        public static void PressOnce(Keys Key)
        {
            keybd_event(Convert.ToByte(Key), Convert.ToByte(Key), 0, 0);
            keybd_event(Convert.ToByte(Key), Convert.ToByte(Key), 0x2, 0);
        }
        private ClaWindow RapButton;
        private ClaWindow zakladkaTabWewnetrzne;
        private ClaWindow filtrBtn;
        private ClaWindow listaPoz;
        private int listID;
        private bool bSpacja = false;
        long sValue;

        public override void Init()
        {
            AddSubscription(true, 0, Events.OpenWindow, new TakeEventDelegate(OnOpenWindow));
            AddSubscription(false, 0, Events.ResizeWindow, new TakeEventDelegate(Resize));
        }

        private bool OnOpenWindow(Procedures ProcId, int ControlId, Events Event)
        {
            ClaWindow Parent = GetWindow();
            Rectangle WndBnd = GetWindow().Bounds;
            zakladkaTabWewnetrzne = GetWindow().AllChildren["?RaportTab"];
            listaPoz = GetWindow().AllChildren["?List"];
            RapButton = zakladkaTabWewnetrzne.Children.Add(ControlTypes.button);
            RapButton.Visible = true;
            RapButton.TextRaw = "SQ";
            RapButton.ToolTipRaw = "Wyświetl raport na podstawie zaznaczonych dokumentów.";
           

            filtrBtn = GetWindow().AllChildren["?Export"];                            //pobiera współrzędne przycisku zapisz
            RapButton.Bounds = new Rectangle(filtrBtn.Bounds.X - 20, filtrBtn.Bounds.Y, 17, 15);
            //tu sprawdzic
            Application.EnableVisualStyles();
            //Application.SetCompatibleTextRenderingDefault(false);
            RapButton.OnAfterAccepted += new TakeEventDelegate(zaczytaj_OnAfterAccepted);
            RapButton.OnAfterMouseDown += new TakeEventDelegate(BtnOnMouseDown);

            //pobrane z pola w zakładce Raporty
            return true;
        }

        private bool Resize(Procedures ProcId, int ControlId, Events Event)
        {
            filtrBtn = GetWindow().AllChildren["?Export"];
            RapButton.Bounds = new Rectangle(filtrBtn.Bounds.X - 20, filtrBtn.Bounds.Y, 17, 15);

            return true;
        }
        bool zaczytaj_OnAfterAccepted(Procedures ProcedureId, int ControlId, Events Event)
        {
            bSpacja = true;
            Runtime.WindowController.PostEvent(listaPoz.Id, Events.NewSelection);
            sValue = Zapisy.KAZ_KRPNumer;
           
            return true;
        }

            private bool BtnOnMouseDown(Procedures ProcId, int ControlId, Events Event)
        {
            //listID = GetWindow().AllChildren["?List"].Id;

            #region komentarze
            //Application.Run(new ONE034Form());
            //Process.Start("C: \\Users\\bartek\\opt\\one034 form\\build one34\\ONE34.exe");
            //Process.Start("C:\\OneSupport\\build one34\\ONE034.exe");
            /* MessageBox.Show(Runtime.ConfigurationDictionary.Baza.ToString(), "Komunikat",
                  MessageBoxButtons.OK, MessageBoxIcon.Information);*/


            /*if (Runtime.ConfigurationDictionary.Baza.ToString().Substring(0,7) == "KopiaXL")
            {
                MessageBox.Show("Po wybraniu pliku nastąpi dodawanie danych do bazy KopiaXL", "Komunikat",
                 MessageBoxButtons.OK, MessageBoxIcon.Information);
                Process.Start("C:\\OneSupport\\ImportPlatnosci\\KopiaXL\\ONE034.exe");
            }
            else if (Runtime.ConfigurationDictionary.Baza.ToString().Substring(0,7) == "DomarXL")
            {
                MessageBox.Show("Po wybraniu pliku nastąpi dodawanie danych do bazy DomarXL", "Komunikat",
                 MessageBoxButtons.OK, MessageBoxIcon.Information);
                Process.Start("C:\\OneSupport\\ImportPlatnosci\\DomarXL\\ONE034.exe");
            }
            else
            {
                MessageBox.Show("Dodatek nie posiada funkcjonalności dla tej bazy.", "Komunikat",
                 MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }*/

            // Process.Start("C:\\Users\\magdaw\\Documents\\Domar\\Excel-Halmar\\one034 form — kopia\\build one34\\ONE034.exe");
            //MessageBox.Show(Convert.ToString(listID), "Komunikat",
            //    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            #endregion
            try
            {
                
                
                //Application.Run(new Logowanie());
                //if(zalogowano)
                //MessageBox.Show(args0 +" i "+ args1);

                Application.Run(new FormWybor(sValue));

               
            }
            catch(Exception e)
            {
                MessageBox.Show("Błąd: "+e.Message);
            }
            return true;

        }

        public override void Cleanup()
        {
        }

    }

   
}
