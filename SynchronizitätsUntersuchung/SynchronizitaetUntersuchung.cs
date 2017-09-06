using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using System.Web.UI;
using System.Globalization;
using Newtonsoft.Json;
using System.Text;

namespace SynchronizitätsUntersuchung
{
    public partial class SynchronizitaetUntersuchung
    {
        // Variablen
        public static NameSpace NameSpace;
        public static Explorer CurrentExplorer;
        public static Folder folder;
        public static string User;
        public static bool Untersuchung_Erweitern = false;
        public static bool Cancel = false;
        public static String Pfad = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "/SynchronizitaetsUntersuchung/";
        public static bool Breite_Werte = false;

        private Dictionary<string, int> speicher;
        List<string> konversationen;
        Dictionary<string, Beziehung> beziehungen;
        ExchangeUser currentUser;
        StreamWriter logfile = null;

        /*
         * Initialisiere das Add-In
         */
        private void SynchronizitaetUntersuchungRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            speicher = new Dictionary<string, int>();
            konversationen = new List<string>();
            beziehungen = new Dictionary<string, Beziehung>();
            NameSpace = Globals.ThisAddIn.Application.GetNamespace("MAPI");
            CurrentExplorer = Globals.ThisAddIn.Application.ActiveExplorer();
        }

        private void btnUntersuchungStarten_Click(object sender, RibbonControlEventArgs e)
        {
            log("Starte Untersuchung");
            speicher = new Dictionary<string, int>();
            Zeige_Dialog();

            // Abbrechen
            if (Cancel)
            {
                return;
            }

            // Es wurde keine E-Mail angegeben
            if (User == "")
            {
                log("Keine E-Mail Adresse angegeben");
                MessageBox.Show("Um die Untersuchung durchzufügen, muss angegeben werden wie deine E-Mail Adresse lautet.");
                return;
            }

            Properties.Settings.Default.UserEmail = User;

            /*try
            {*/
                // Zunächst Ordner für Auswertung erstellen wenn dieser noch nicht existiert
                Directory.CreateDirectory(Pfad);
                // Dann prüfen ob erweitert oder von vorne untersucht werden soll
                if (Untersuchung_Erweitern) Daten_Aus_Datei_Lesen();
                MessageBox.Show("Die Synchronizitätsuntersuchung wird gestartet. Diese kann abhängig von der Größe des Mailordners mehrere Minuten dauern.", "Start", MessageBoxButtons.OK, MessageBoxIcon.Information);
                // Dann Untersuchung starten
                Starte_Untersuchung();
                Auswerten();
                Erstelle_html_datei();
                Daten_In_Datei_Schreiben();
                MessageBox.Show("Die Synchronizitätsuntersuchung wurde erfolgreich abgeschlossen.", "Erfolg", MessageBoxButtons.OK, MessageBoxIcon.Information);

            /*foreach(KeyValuePair<string, Beziehung> bez in beziehungen)
            {
                Debug.WriteLine("Beziehung: " + bez.Key);

                foreach(KeyValuePair<string, Konversation> konv in bez.Value.Konversationen)
                {
                    Debug.WriteLine("\tKonversation: " + konv.Value.Thema + " mit " + konv.Value.Laenge + " Emails");
                }
            }*/
            /*}
            catch (System.Exception)
            {
                MessageBox.Show("Es ist ein Fehler aufgetreten, bitte führe die Untersuchung erneut durch.", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }*/
        }

        private void Zeige_Dialog()
        {
            Form dialog = new Form()
            {
                Width = 700,
                Height = 200,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = "Deine E-Mail Adresse, die du in dem zu untersuchenden Postfach verwendest.",
                StartPosition = FormStartPosition.CenterScreen
            };

            Properties.Settings.Default.UserEmail = "maik.schmaddebeck@tu-clausthal.de";

            Label textLabel = new Label() { Left = 40, Top = 15, Width = 610, Text = "E-Mail Adresse" };
            TextBox email = new TextBox() { Left = 40, Top = 35, Width = 610, Text = Properties.Settings.Default.UserEmail };
            CheckBox override_btn = new CheckBox() { Text = "Komplett neue Untersuchung durchführen und alle bestehende Daten überschreiben", Left = 40, Top = 90, Width = 610, Checked = false };
            CheckBox eigenewerte_btn = new CheckBox() { Text = "Für die Untersuchung breitere Werte verwenden", Left = 40, Top = 65, Width = 610, Checked = false };
            Button confirmation = new Button() { Text = "Ok", Left = 340, Width = 100, Top = 120, DialogResult = DialogResult.OK };
            Button cancel = new Button() { Text = "Abbrechen", Left = 450, Width = 100, Top = 120, DialogResult = DialogResult.Cancel };

            dialog.Controls.Add(email);
            dialog.Controls.Add(textLabel);
            dialog.Controls.Add(override_btn);
            dialog.Controls.Add(eigenewerte_btn);
            dialog.Controls.Add(confirmation);
            dialog.Controls.Add(cancel);
            dialog.AcceptButton = confirmation;
            dialog.CancelButton = cancel;

            DialogResult result = dialog.ShowDialog();

            switch (result)
            {
                // put in how you want the various results to be handled
                // if ok, then something like var x = dialog.MyX;
                case DialogResult.OK:
                    log("E-Mail Adresse \"" + email.Text + "\" angegeben.");
                    log("CheckBox \"Überschreiben\": " + override_btn.Checked);
                    log("CheckBox \"BreiteWerte\": " + eigenewerte_btn.Checked);
                    User = email.Text;
                    Untersuchung_Erweitern = !override_btn.Checked;
                    Breite_Werte = eigenewerte_btn.Checked;
                    break;
                default:
                    Cancel = true;
                    log("Dialog abgebrochen");
                    break;
            }

        }

        private void Daten_Aus_Datei_Lesen()
        {    
            // Existiert noch keine JSON Datei dann den Schritt zum Laden überpsringen
            if(!File.Exists(Pfad + "SynchronizitaetsUntersuchung.json")) {
                return;
            }

            // Open the text file using a stream reader.
            using (StreamReader file = new StreamReader(Pfad + "SynchronizitaetsUntersuchung.json"))
            {
                // Read the stream to a string, and write the string to the console.
                String json = file.ReadToEnd();
                JsonTextReader reader = new JsonTextReader(new StringReader(json));

                string partner = "";
                Dictionary<string,Beziehung> liste_beziehungen = new Dictionary<string, Beziehung>();
                    
                // Speicher auslesen
                while (reader.TokenType != JsonToken.EndArray && reader.Read())
                {
                    if (reader.Value != null && reader.Value.ToString() == "folder")
                    {
                        reader.ReadAsString();
                        string folder = reader.Value.ToString();
                        reader.Read();
                        reader.ReadAsInt32();
                        int lesezeichen = (int)reader.Value;
                        speicher[folder] = lesezeichen;
                    }
                }
                reader.Read();

                // Beziehungen auslesen
                Beziehung beziehung = null;
                while (reader.TokenType != JsonToken.EndArray && reader.Read())
                {
                    if (reader.Value != null && reader.Value.ToString() == "partner")
                    {
                        reader.ReadAsString();
                        partner = reader.Value.ToString();
                        beziehung = new Beziehung(partner);
                    }

                    // Antworten auslesen
                    if (reader.Value != null && reader.Value.ToString() == "antworten")
                    {
                        while (reader.TokenType != JsonToken.EndArray && reader.Read())
                        {
                            if (reader.Value != null && reader.Value.ToString() == "antwortzeit")
                            {
                                reader.ReadAsDouble(); ;
                                double antwortzeit = (double)reader.Value;
                                reader.Read();
                                reader.ReadAsString();
                                string tageszeit = reader.Value.ToString();
                                reader.Read();
                                reader.Read();
                                DayOfWeek wochentag = (DayOfWeek)Enum.ToObject(typeof(DayOfWeek), reader.Value);

                                Antwort antwort = new Antwort(antwortzeit, tageszeit, wochentag);
                                beziehung.Antworten.Add(antwort);
                            }
                        }
                        reader.Read();
                    }

                    // Konversationen auslesen
                    if (reader.Value != null && reader.Value.ToString() == "konversationen")
                    {
                        while (reader.TokenType != JsonToken.EndArray && reader.Read())
                        {
                            if (reader.Value != null && reader.Value.ToString() == "id")
                            {
                                reader.ReadAsString();
                                string id = reader.Value.ToString();
                                Konversation konv = new Konversation(id);
                                reader.Read();
                                reader.Read();
                                reader.Read();

                                // Antwortzeiten auslesen
                                while (reader.TokenType != JsonToken.EndArray && reader.Read())
                                {
                                    if (reader.Value != null && reader.Value.ToString() == "antwortzeit")
                                    {
                                        reader.ReadAsDouble();
                                        double antwortzeit = (double)reader.Value;
                                        konv.Antwort_Hinzufuegen(antwortzeit);
                                    }
                                }
                                reader.Read();

                                beziehung.Konversation_Hinzufuegen(konv);
                            }
                        }
                        reader.Read();
                    }

                    if (reader.TokenType == JsonToken.EndObject)
                    {
                        liste_beziehungen.Add(partner, beziehung);
                    }
                }

                beziehungen = liste_beziehungen;
            }
        }

        private void Starte_Untersuchung() 
        {
            Properties.Settings.Default.UserEmail = User;

            // Outlook Account auslesen, mitdem im Outlook angemeldet ist
            
            AddressEntry addrEntry = NameSpace.CurrentUser.AddressEntry;
            if (addrEntry.Type == "EX")
            {
                currentUser = addrEntry.GetExchangeUser();
            }

            //Folder root = NameSpace.Session.DefaultStore.GetRootFolder() as Folder;
            Folder root = CurrentExplorer.CurrentFolder as Folder;
            UntersucheOrdner(root);
            
        }

        private void UntersucheOrdner(Folder folder)
        {
            log("Untersuche [Ordner] " + folder.Name);
            Konversation konversation;
            string thema;
            string partner;

            bool letzteMailWurdeEmpfangen;
            DateTime creationTimeDerLetztenEmpfangenenMail; // Mit jetzigem TimeStamp versehen, da irgendwas initialisiert werden muss
            DateTime receivedTimeDerLetztenEmpfangenenMail; // Mit jetzigem TimeStamp versehen, da irgendwas initialisiert werden muss

            bool letzteMailWurdeGesendet;
            DateTime sentOnDerLetztenGesendetenMail; // Mit jetzigem TimeStamp versehen, da irgendwas initialisiert werden muss
            DateTime receivedTimeDerLetztenGesendetenMail; // Mit jetzigem TimeStamp versehen, da irgendwas initialisiert werden muss

            // Setze das Lesezeichen, falls für den ausgewählten Ordner bereits eine Untersuchung durchgeführt wurde
            int lesezeichen = 1;
            if (speicher.ContainsKey(folder.EntryID) && speicher[folder.EntryID] <= folder.Items.Count)
            {
                lesezeichen = speicher[folder.EntryID];
                log("Lesezeichen bei " + lesezeichen + " gesetzt.");
            }

            /*
             * Handelt es sich nicht um einen Ordner der E-Mails enthält, beende das Programm
             */
            if (folder.DefaultItemType != OlItemType.olMailItem)
            {
                return;
            }

            /*
             * Iteriere über jedes Objekt in dem zuvor ausgewählten Ordner
             */
            for (int item = lesezeichen; item <= folder.Items.Count; item++)
            {
                speicher[folder.EntryID] = item;
                MailItem mail = folder.Items[item] as MailItem;

                /*
                 * Überspringe das aktuelle Item, wenn es sich dabei nicht um eine E-Mail handelt
                 */
                if (mail == null)
                {
                    continue;
                }

                /*
                 * Hole die zur E-Mail gehörende Konversation und prüfe ob die Konversation bereits abgehandelt wurde.
                 * Wenn nicht, füge sie in das Dictionary aller gefundenen Konversationen ein, mit denen später weitergearbeitet wird.
                 */
                Conversation k = mail.GetConversation();
                if (k != null && !konversationen.Contains(k.ConversationID))
                {
                    /*
                     * Es handelt sich um eine neue Konversation die zu bearbeiten ist, initialisiere deshalb alle nötigen Variablen
                     */
                    konversation = new Konversation(k.ConversationID);
                    thema = null;
                    partner = null;

                    letzteMailWurdeEmpfangen = false;
                    creationTimeDerLetztenEmpfangenenMail = DateTime.Now; // Mit jetzigem TimeStamp versehen, da irgendwas initialisiert werden muss
                    receivedTimeDerLetztenEmpfangenenMail = DateTime.Now; // Mit jetzigem TimeStamp versehen, da irgendwas initialisiert werden muss

                    letzteMailWurdeGesendet = false;
                    sentOnDerLetztenGesendetenMail = DateTime.Now; // Mit jetzigem TimeStamp versehen, da irgendwas initialisiert werden muss
                    receivedTimeDerLetztenGesendetenMail = DateTime.Now; // Mit jetzigem TimeStamp versehen, da irgendwas initialisiert werden muss

                    Table table = k.GetTable();
                    
                    // Wenn die Konversation mindestens aus 2 E-Mails besteht
                    if (table.GetRowCount() > 1)
                    {
                        dynamic[,] temp = table.GetArray(table.GetRowCount());
                        string entryID;

                        List<MailItem> mailItems = new List<MailItem>();

                        for (int i = 0; i < temp.GetLength(0); i++)
                        {
                            entryID = ((object)temp[i, 0]).ToString();

                            MailItem mailitem = NameSpace.GetItemFromID(entryID, folder.StoreID) as MailItem;
                            if (mailitem != null)
                            {
                                mailItems.Add(mailitem);
                            }
                        }
                        mailItems.Sort((x, y) => DateTime.Compare(x.ReceivedTime, y.ReceivedTime));

                        // Iteriere alle E-Mails der Konversation
                        foreach (MailItem mailItem in mailItems)
                        {
                            String senderEmailAddress = "";

                            // Die Mail nur betrachten wenn sie abgesendet wurde (also kein Draft ist)
                            if (mailItem.Sent)
                            {

                                if (string.IsNullOrEmpty(thema))
                                {
                                    thema = mailItem.ConversationTopic;
                                    konversation.Thema = thema;
                                    log("[Konversation] \"" + thema + "\"");
                                }

                                if (!string.IsNullOrEmpty(mailItem.SenderEmailAddress))
                                {
                                    senderEmailAddress = mailItem.SenderEmailAddress;
                                }

                                // Es handelt sich um eine gesendete E-Mail
                                if (senderEmailAddress == User || (currentUser != null && senderEmailAddress.Equals(currentUser.Address, StringComparison.OrdinalIgnoreCase)))
                                {
                                    log("\t[E-Mail] von " + senderEmailAddress + " am " + mailItem.SentOn);
                                    /*
                                     * Aktualisiere den Synchronizitätswert für diese Konversation, in dem die Zeit zwischen
                                     * Empfangen der letzten E-Mail und Senden der aktuellen E-Mail berechnet und mit dem letzten 
                                     * Synchronizitätswert addiert wird.
                                     */
                                    if (letzteMailWurdeEmpfangen)
                                    {
                                        // Sollte die E-Mail beantwortet worden sein, bevor diese auf dem Computer geladen wurde (z.B. per Mobil beantwortet), dann nimm die
                                        // Zeit des Eingangs auf dem Server zur Berechnung
                                        DateTime antwort_zeitstempel = (mailItem.SentOn < creationTimeDerLetztenEmpfangenenMail) ? receivedTimeDerLetztenEmpfangenenMail : creationTimeDerLetztenEmpfangenenMail;
                                        TimeSpan antwort_antwortzeit = mailItem.SentOn - antwort_zeitstempel;
                                        TimeSpan konversation_antwortzeit = mailItem.SentOn - receivedTimeDerLetztenEmpfangenenMail;

                                        // Es wird eine neue Antwort erstellt und diese der passenden Beziehung zugeordnet
                                        beziehungen[partner].Antworten.Add(new Antwort(antwort_antwortzeit.TotalSeconds, antwort_zeitstempel, mailItem.SentOn));
                                        // Neue Antwort der Konversation hinzufügen
                                        beziehungen[partner].Konversationen[konversation.Id].Antwort_Hinzufuegen(konversation_antwortzeit.TotalSeconds);
                                    }

                                    // Wenn mehr als eine Nachricht hintereinander gesendet werden, ohne dass selbst eine E-Mail empfangen wird, dann
                                    // überschreib den Zeitstempel nicht mit der neuen gesendeten E-Mail, sondern behalte den der älteren bei
                                    if (!letzteMailWurdeGesendet)
                                    {
                                        receivedTimeDerLetztenGesendetenMail = mailItem.ReceivedTime;
                                        sentOnDerLetztenGesendetenMail = mailItem.SentOn;
                                    }

                                    // Auf False setzen, da es sich hier um eine gesendete Mail handelt
                                    letzteMailWurdeEmpfangen = false;
                                    letzteMailWurdeGesendet = true;
                                }
                                // Es handelt sich um eine empfangene E-Mail
                                else
                                {
                                    log("\t[E-Mail] von " + senderEmailAddress + " am " + mailItem.ReceivedTime + " (Server) " + mailItem.CreationTime + " (Outlook)");

                                    // Setze bei der ersten Empfangenen E-Mail den Kommunikationspartner und erstelle falls diese noch nicht vorhandne ist eine neue Beziehung
                                    if (string.IsNullOrEmpty(partner))
                                    {
                                        partner = mailItem.SenderName;
                                        if (!beziehungen.ContainsKey(partner))
                                        {
                                            beziehungen.Add(partner, new Beziehung(partner));
                                        }
                                        // Füge der Beziehung sofort die Konversation hinzu
                                        beziehungen[partner].Konversation_Hinzufuegen(konversation);

                                    }

                                    // Zähle die E-Mail als Empfangen
                                    beziehungen[partner].AnzahlErhalteneEmails++;

                                    // Aktualisiere Gesprächs Synchronizität
                                    if (letzteMailWurdeGesendet)
                                    {
                                        // Neue Antwort der Konversation hinzufügen
                                        TimeSpan antwortzeit = mailItem.ReceivedTime - sentOnDerLetztenGesendetenMail;
                                        beziehungen[partner].Konversationen[konversation.Id].Antwort_Hinzufuegen(antwortzeit.TotalSeconds);
                                    }

                                    // Wenn mehr als eine Nachricht hintereinander empfangen werden, ohne dass selbst eine E-Mail abgeschickt wird, dann
                                    // überschreib den Zeitstempel nicht mit der neuen empfangenen E-Mail, sondern behalte den der älteren bei
                                    if (!letzteMailWurdeEmpfangen)
                                    {
                                        creationTimeDerLetztenEmpfangenenMail = mailItem.CreationTime;
                                        receivedTimeDerLetztenEmpfangenenMail = mailItem.ReceivedTime;
                                    }

                                    // Es handelt sich um eine empfangene Mail, deshalb Schalter setzen
                                    letzteMailWurdeEmpfangen = true;
                                    letzteMailWurdeGesendet = false;
                                }


                            }

                        }

                    }

                    konversationen.Add(konversation.Id);

                }

            }

            // Iteriere über alle Unterordner rekursiv
            Folders childFolders = folder.Folders;

            if (childFolders.Count > 0)
            {
                foreach (Folder childFolder in childFolders)
                {
                    // Write the folder path.
                    Debug.WriteLine("##############################################################");
                    Debug.WriteLine(childFolder.FolderPath);
                    Debug.WriteLine("##############################################################");
                    // Call EnumerateFolders using childFolder.
                    UntersucheOrdner(childFolder);
                }
            }

        }

        private void Auswerten()
        {
            foreach (KeyValuePair<string, Beziehung> pair in beziehungen)
            {
                pair.Value.Beziehung_Auswerten();
            }
        }

        private void Erstelle_html_datei()
        {
            // Initialize StringWriter instance.
            StringWriter stringWriter = new StringWriter();
            var beziehungen_sortiert = from pair in beziehungen orderby pair.Value.Antworten.Count descending select pair;

            int sehr_schnell = 0;
            int schnell = 0;
            int normal = 0;
            int langsam = 0;
            int sehr_langsam = 0;

            int komplettSynchron = 0;
            int groeßtenteilsSynchron = 0;
            int eherSynchron = 0;
            int gleichmaeßigSynchronUndAsynchron = 0;
            int eherAsynchron = 0;
            int groesstenteilsAsynchron = 0;
            int komplettAsynchron = 0;
            Beziehung beziehung = new Beziehung("");

            Dictionary<Synchronizitaet, int> gesamt_synchornizitaet = new Dictionary<Synchronizitaet, int>();
            gesamt_synchornizitaet[Synchronizitaet.KomplettSynchron] = 0;
            gesamt_synchornizitaet[Synchronizitaet.GroesstenteilsSynchron] = 0;
            gesamt_synchornizitaet[Synchronizitaet.EherSynchron] = 0;
            gesamt_synchornizitaet[Synchronizitaet.GleichmaeßigSynchronUndAsynchron] = 0;
            gesamt_synchornizitaet[Synchronizitaet.EherAsynchron] = 0;
            gesamt_synchornizitaet[Synchronizitaet.GroesstenteilsAsynchron] = 0;
            gesamt_synchornizitaet[Synchronizitaet.KomplettAsynchron] = 0;

            Dictionary <Antwortzeit, int> gesamt_antwortzeit = new Dictionary<Antwortzeit, int>();
            gesamt_antwortzeit[Antwortzeit.SehrSchnell] = 0;
            gesamt_antwortzeit[Antwortzeit.Schnell] = 0;
            gesamt_antwortzeit[Antwortzeit.Normal] = 0;
            gesamt_antwortzeit[Antwortzeit.Langsam] = 0;
            gesamt_antwortzeit[Antwortzeit.SehrLangsam] = 0;

            Dictionary<DayOfWeek, List<double>> gesamt_tagesabhaengigeAntworten = new Dictionary<DayOfWeek, List<double>>();
            gesamt_tagesabhaengigeAntworten[DayOfWeek.Monday] = new List<double>();
            gesamt_tagesabhaengigeAntworten[DayOfWeek.Tuesday] = new List<double>();
            gesamt_tagesabhaengigeAntworten[DayOfWeek.Wednesday] = new List<double>();
            gesamt_tagesabhaengigeAntworten[DayOfWeek.Thursday] = new List<double>();
            gesamt_tagesabhaengigeAntworten[DayOfWeek.Friday] = new List<double>();
            gesamt_tagesabhaengigeAntworten[DayOfWeek.Saturday] = new List<double>();
            gesamt_tagesabhaengigeAntworten[DayOfWeek.Sunday] = new List<double>();

            Dictionary<string, List<double>> gesamt_tageszeitabhaengigeAntworten = new Dictionary<string, List<double>>();
            gesamt_tageszeitabhaengigeAntworten[Tageszeitabhaengigkeit.Morgen] = new List<double>();
            gesamt_tageszeitabhaengigeAntworten[Tageszeitabhaengigkeit.Vormittag] = new List<double>();
            gesamt_tageszeitabhaengigeAntworten[Tageszeitabhaengigkeit.Mittag] = new List<double>();
            gesamt_tageszeitabhaengigeAntworten[Tageszeitabhaengigkeit.Nachmittag] = new List<double>();
            gesamt_tageszeitabhaengigeAntworten[Tageszeitabhaengigkeit.Abend] = new List<double>();
            gesamt_tageszeitabhaengigeAntworten[Tageszeitabhaengigkeit.Nacht] = new List<double>();

            string string_tagesabhaengigkeit = "";

            string string_tageszeitabhaengigkeit = "";

            // Put HtmlTextWriter in using block because it needs to call Dispose.
            using (HtmlTextWriter writer = new HtmlTextWriter(stringWriter, String.Empty))
            {

                writer.RenderBeginTag(HtmlTextWriterTag.Html);

                writer.RenderBeginTag(HtmlTextWriterTag.Head);

                // CSS Bootstrap
                writer.AddAttribute(HtmlTextWriterAttribute.Rel, "stylesheet");
                writer.AddAttribute(HtmlTextWriterAttribute.Href, "https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0-beta/css/bootstrap.min.css");
                writer.RenderBeginTag(HtmlTextWriterTag.Link);
                writer.RenderEndTag();

                // Javascript JQuery
                writer.AddAttribute(HtmlTextWriterAttribute.Src, "https://ajax.googleapis.com/ajax/libs/jquery/3.2.1/jquery.min.js");
                writer.RenderBeginTag(HtmlTextWriterTag.Script);
                writer.RenderEndTag();

                // Javascript Plotly
                writer.AddAttribute(HtmlTextWriterAttribute.Src, "https://cdn.plot.ly/plotly-1.2.0.min.js");
                writer.RenderBeginTag(HtmlTextWriterTag.Script);
                writer.RenderEndTag();

                // CSS
                writer.RenderBeginTag(HtmlTextWriterTag.Style);
                writer.WriteLine("a {text-decoration: none;}");
                writer.WriteLine(".BeziehungenContainer {float:left; height: 100%; width: 250px; position: fixed; left: 0; top: 0; border-style: ridge;}");
                writer.WriteLine(".BeziehungenInhaltsverzeichnis {height: 100%; overflow: auto;}");
                writer.WriteLine(".Diagramme {margin-left: 250px;}");
                writer.RenderEndTag();

                writer.RenderEndTag(); // head

                writer.RenderBeginTag(HtmlTextWriterTag.Body);

                writer.AddAttribute(HtmlTextWriterAttribute.Class, "BeziehungenContainer");
                writer.RenderBeginTag(HtmlTextWriterTag.Div);

                writer.AddAttribute(HtmlTextWriterAttribute.Class, "BeziehungenInhaltsverzeichnis");
                writer.RenderBeginTag(HtmlTextWriterTag.Div);

                /*
                 * INHALTSVERZEICHNIS
                 */
                int counter = 0;

                // ALle Beziehungen interieren und für jeden einen Inhaltsverzeichnis eintrag anlegen, zusätzlich mit Index 0 ein Gesamt Eintrag
                for (int i = 0; (i < beziehungen_sortiert.Count() + 1) && (i == 0 || beziehungen_sortiert.ElementAt(i - 1).Value.Antworten.Count() > 0); i++)
                {
                    counter = i - 1;

                    if (i > 0)
                    {
                        beziehung = beziehungen_sortiert.ElementAt(counter).Value;
                    }

                    writer.AddAttribute(HtmlTextWriterAttribute.Href, "#beziehung" + i);
                    writer.RenderBeginTag(HtmlTextWriterTag.A); // Link zum öffnen des DIVs
                    if (i == 0)
                    {
                        writer.WriteEncodedText("Gesamt");
                    }
                    else
                    {
                        writer.WriteEncodedText(beziehungen_sortiert.ElementAt(counter).Key);
                    }
                    writer.RenderEndTag(); // Link zum öffnen des DIVs
                    writer.WriteBreak();
                }

                writer.RenderEndTag(); // Beziehungen Inhaltsverzeichnis DIV
                writer.RenderEndTag(); // Beziehungen Container DIV

                /*
                 * DIAGRAMME
                 */
                writer.AddAttribute(HtmlTextWriterAttribute.Class, "Diagramme");
                writer.RenderBeginTag(HtmlTextWriterTag.Div);

                for (int i = 0; (i < beziehungen_sortiert.Count() + 1) && (i == 0 || beziehungen_sortiert.ElementAt(i - 1).Value.Antworten.Count() > 0); i++)
                {
                    counter = i - 1;
                    if (i > 0)
                    {
                        beziehung = beziehungen_sortiert.ElementAt(counter).Value;
                    }

                    writer.AddAttribute(HtmlTextWriterAttribute.Class, "container");
                    writer.AddAttribute(HtmlTextWriterAttribute.Id, "beziehung" + i);
                    writer.RenderBeginTag(HtmlTextWriterTag.Div); // Beziehung DIV
                    writer.RenderBeginTag(HtmlTextWriterTag.H2);

                    if (i == 0)
                    {
                        writer.WriteEncodedText("Gesamt");
                        writer.RenderEndTag();
                    }
                    else
                    {
                        writer.WriteEncodedText(beziehungen_sortiert.ElementAt(counter).Key);
                        writer.RenderEndTag();
                        writer.WriteEncodedText("Insgesamt hast du " + beziehung.Konversationen.Count() + " Konversationen mit einer durchschnittlichen Konversationslänge von " + beziehung.DurchschnittlicheKonversationslaenge + " E-Mails mit "
                            + beziehung.Partner + " geführt. Unabhängig von den Konversationen hast du insgesamt " + beziehung.Antworten.Count + " Antworten gesendet.");
                    }

                    if (i > 0 && beziehung.Antworten.Count < 1)
                    {
                        writer.WriteEncodedText(" Aufgrund von zuwenigen Daten können zu dieser Beziehung keine Statistiken erstellt werden.");
                    }
                    else
                    {
                        writer.AddAttribute(HtmlTextWriterAttribute.Class, "row");
                        writer.RenderBeginTag(HtmlTextWriterTag.Div);

                        writer.AddAttribute(HtmlTextWriterAttribute.Class, "col");
                        writer.RenderBeginTag(HtmlTextWriterTag.Div);
                        writer.AddAttribute(HtmlTextWriterAttribute.Id, "synchronizitaet" + i);
                        writer.RenderBeginTag(HtmlTextWriterTag.Div); // synchronizitaet DIV
                        if (i > 0 && beziehung.Konversationen.Count < 1)
                        {
                            writer.WriteEncodedText("Es wurden keine Konversationen mit " + beziehung.Partner + " geführt.");
                        }
                        writer.RenderEndTag(); // synchronizitaet DIV
                        writer.RenderEndTag(); // col

                        writer.AddAttribute(HtmlTextWriterAttribute.Class, "col");
                        writer.RenderBeginTag(HtmlTextWriterTag.Div);
                        writer.AddAttribute(HtmlTextWriterAttribute.Id, "kuchendiagramm" + i);
                        writer.RenderBeginTag(HtmlTextWriterTag.Div); // Kuchendiagramm DIV
                        writer.RenderEndTag(); // Kuchendiagramm DIV
                        writer.RenderEndTag(); // col

                        writer.RenderEndTag(); // row


                        writer.AddAttribute(HtmlTextWriterAttribute.Class, "row");
                        writer.RenderBeginTag(HtmlTextWriterTag.Div);

                        writer.AddAttribute(HtmlTextWriterAttribute.Class, "col");
                        writer.RenderBeginTag(HtmlTextWriterTag.Div);
                        writer.AddAttribute(HtmlTextWriterAttribute.Id, "tagesabhaengikeit" + i);
                        writer.RenderBeginTag(HtmlTextWriterTag.Div); // Tagesabhängigkeit DIV
                        writer.RenderEndTag(); // Tagesabhängigkeit DIV
                        writer.RenderEndTag(); // col

                        writer.AddAttribute(HtmlTextWriterAttribute.Class, "col");
                        writer.RenderBeginTag(HtmlTextWriterTag.Div);
                        writer.AddAttribute(HtmlTextWriterAttribute.Id, "tageszeitabhaengikeit" + i);
                        writer.RenderBeginTag(HtmlTextWriterTag.Div); // Tageszeitabhängigkeit DIV
                        writer.RenderEndTag(); // Tageszeitabhängigkeit DIV
                        writer.RenderEndTag(); // col

                        writer.RenderEndTag(); // row
                    }

                    writer.RenderEndTag(); // Beziehung DIV

                    writer.RenderBeginTag(HtmlTextWriterTag.Hr);
                }
                writer.RenderEndTag(); // Diagramme DIV


                /*
                * Javascript 
                */
                writer.RenderBeginTag(HtmlTextWriterTag.Script);
                string data = "";
                string layout = "";

                // Iteriere alle Beziehungen rückwärts um Gesamtanzahlen zählen zu können und Gesamtstatistiken am Ende zu berechnen
                for (int i = beziehungen_sortiert.Count(); (i >= 0); i--)
                {
                    if (i == 0 || beziehungen_sortiert.ElementAt(i - 1).Value.Antworten.Count > 0)
                    {
                        counter = i - 1;
                        
                        NumberFormatInfo nfi = new NumberFormatInfo();
                        nfi.NumberDecimalSeparator = ".";
                        if (i > 0)
                        {
                            beziehung = beziehungen_sortiert.ElementAt(counter).Value;
                            komplettSynchron = beziehung.VerteilungGespraeche[Synchronizitaet.KomplettSynchron];
                            groeßtenteilsSynchron = beziehung.VerteilungGespraeche[Synchronizitaet.GroesstenteilsSynchron];
                            eherSynchron = beziehung.VerteilungGespraeche[Synchronizitaet.EherSynchron];
                            gleichmaeßigSynchronUndAsynchron = beziehung.VerteilungGespraeche[Synchronizitaet.GleichmaeßigSynchronUndAsynchron]; 
                            eherAsynchron = beziehung.VerteilungGespraeche[Synchronizitaet.EherAsynchron];
                            groesstenteilsAsynchron = beziehung.VerteilungGespraeche[Synchronizitaet.GroesstenteilsAsynchron];
                            komplettAsynchron = beziehung.VerteilungGespraeche[Synchronizitaet.KomplettAsynchron];
                            
                            sehr_schnell = beziehung.VerteilungAntworten[Antwortzeit.SehrSchnell];
                            schnell = beziehung.VerteilungAntworten[Antwortzeit.Schnell];
                            normal = beziehung.VerteilungAntworten[Antwortzeit.Normal];
                            langsam = beziehung.VerteilungAntworten[Antwortzeit.Langsam];
                            sehr_langsam = beziehung.VerteilungAntworten[Antwortzeit.SehrLangsam];

                            gesamt_synchornizitaet[Synchronizitaet.KomplettSynchron] += komplettSynchron;
                            gesamt_synchornizitaet[Synchronizitaet.GroesstenteilsSynchron] += groeßtenteilsSynchron;
                            gesamt_synchornizitaet[Synchronizitaet.EherSynchron] += eherSynchron;
                            gesamt_synchornizitaet[Synchronizitaet.GleichmaeßigSynchronUndAsynchron] += gleichmaeßigSynchronUndAsynchron;
                            gesamt_synchornizitaet[Synchronizitaet.EherAsynchron] += eherAsynchron;
                            gesamt_synchornizitaet[Synchronizitaet.GroesstenteilsAsynchron] += groesstenteilsAsynchron;
                            gesamt_synchornizitaet[Synchronizitaet.KomplettAsynchron] += komplettAsynchron;

                            gesamt_antwortzeit[Antwortzeit.SehrSchnell] += sehr_schnell;
                            gesamt_antwortzeit[Antwortzeit.Schnell] += schnell;
                            gesamt_antwortzeit[Antwortzeit.Normal] += normal;
                            gesamt_antwortzeit[Antwortzeit.Langsam] += langsam;
                            gesamt_antwortzeit[Antwortzeit.SehrLangsam] += sehr_langsam;

                            string_tagesabhaengigkeit = "";
                            foreach (KeyValuePair<DayOfWeek, List<double>> tagesabhaengigkeit in beziehung.TagesabhaengigeAntworten)
                            {
                                string_tagesabhaengigkeit += "{y: [" + string.Join(", ", tagesabhaengigkeit.Value.Select(x => (x / 3600).ToString(nfi))) + "], type: 'box', name: '" + tagesabhaengigkeit.Key + "', boxmean: true},";
                                gesamt_tagesabhaengigeAntworten[tagesabhaengigkeit.Key].AddRange(tagesabhaengigkeit.Value);
                            }

                            string_tageszeitabhaengigkeit = "";
                            foreach (KeyValuePair<string, List<double>> tageszeitabhaengigkeit in beziehung.TageszeitabhaengigeAntworten)
                            {
                                string_tageszeitabhaengigkeit += "{y: [" + string.Join(", ", tageszeitabhaengigkeit.Value.Select(x => (x / 3600).ToString(nfi))) + "], type: 'box', name: '" + tageszeitabhaengigkeit.Key + "', boxmean: true},";
                                gesamt_tageszeitabhaengigeAntworten[tageszeitabhaengigkeit.Key].AddRange(tageszeitabhaengigkeit.Value);
                            }

                            
                        }
                        else
                        {
                            komplettSynchron = gesamt_synchornizitaet[Synchronizitaet.KomplettSynchron];
                            groeßtenteilsSynchron = gesamt_synchornizitaet[Synchronizitaet.GroesstenteilsSynchron];
                            eherSynchron = gesamt_synchornizitaet[Synchronizitaet.EherSynchron];
                            gleichmaeßigSynchronUndAsynchron = gesamt_synchornizitaet[Synchronizitaet.GleichmaeßigSynchronUndAsynchron];
                            eherAsynchron = gesamt_synchornizitaet[Synchronizitaet.EherAsynchron];
                            groesstenteilsAsynchron = gesamt_synchornizitaet[Synchronizitaet.GroesstenteilsAsynchron];
                            komplettAsynchron = gesamt_synchornizitaet[Synchronizitaet.KomplettAsynchron];

                            sehr_schnell = gesamt_antwortzeit[Antwortzeit.SehrSchnell];
                            schnell = gesamt_antwortzeit[Antwortzeit.Schnell];
                            normal = gesamt_antwortzeit[Antwortzeit.Normal];
                            langsam = gesamt_antwortzeit[Antwortzeit.Langsam];
                            sehr_langsam = gesamt_antwortzeit[Antwortzeit.SehrLangsam];

                            string_tagesabhaengigkeit = "";
                            foreach (KeyValuePair<DayOfWeek, List<double>> tagesabhaengigkeit in gesamt_tagesabhaengigeAntworten)
                            {
                                string_tagesabhaengigkeit += "{y: [" + string.Join(", ", tagesabhaengigkeit.Value.Select(x => (x / 3600).ToString(nfi))) + "], type: 'box', name: '" + tagesabhaengigkeit.Key + "', boxmean: true},";
                            }

                            string_tageszeitabhaengigkeit = "";
                            foreach (KeyValuePair<string, List<double>> tageszeitabhaengigkeit in gesamt_tageszeitabhaengigeAntworten)
                            {
                                string_tageszeitabhaengigkeit += "{y: [" + string.Join(", ", tageszeitabhaengigkeit.Value.Select(x => (x / 3600).ToString(nfi))) + "], type: 'box', name: '" + tageszeitabhaengigkeit.Key + "', boxmean: true},";
                            }
                        }

                        /*
                        * Gespräche Verteilung
                        */
                        if (beziehung.Konversationen.Count >= 1 || i == 0)
                        {
                            data = "[{" +
                                "values: [" + komplettSynchron + "," + groeßtenteilsSynchron + "," + eherSynchron + "," + gleichmaeßigSynchronUndAsynchron + "," + eherAsynchron + "," + groesstenteilsAsynchron + "," + komplettAsynchron + "], " +
                                "labels: ['" +
                                        SynchronizitaetLabel.KomplettSynchron + "', '" +
                                        SynchronizitaetLabel.GroesstenteilsSynchron + "', '" +
                                        SynchronizitaetLabel.EherSynchron + "', '" +
                                        SynchronizitaetLabel.GleichmaeßigSynchronUndAsynchron + "', '" +
                                        SynchronizitaetLabel.EherAsynchron + "', '" +
                                        SynchronizitaetLabel.GroesstenteilsAsynchron + "', '" +
                                        SynchronizitaetLabel.KomplettAsynchron +
                                        "'], " +
                                "type: 'pie', " +
                                "marker: {" +
                                    "colors: ['" +
                                        SynchronizitaetFarbe.KomplettSynchron + "', '" +
                                        SynchronizitaetFarbe.GroesstenteilsSynchron + "', '" +
                                        SynchronizitaetFarbe.EherSynchron + "', '" +
                                        SynchronizitaetFarbe.GleichmaeßigSynchronUndAsynchron + "', '" +
                                        SynchronizitaetFarbe.EherAsynchron + "', '" +
                                        SynchronizitaetFarbe.GroesstenteilsAsynchron + "', '" +
                                        SynchronizitaetFarbe.KomplettAsynchron +
                                        "']" +
                                "}" +
                            "}]";
                            layout = "{title: 'Prozentuale Verteilung der Synchronizität der Gespräche'}";
                            writer.WriteLineNoTabs("Plotly.newPlot('" + "synchronizitaet" + i + "', " + data + "," + layout + "); ");
                        }

                        /*
                        * Antwortzeiten Verteilung
                        */
                        if (beziehung.Antworten.Count >= 1 || i == 0)
                        {
                            data = "[{" +
                                "values: [" + sehr_schnell + "," + schnell + "," + normal + "," + langsam + "," + sehr_langsam + "], " +
                                "labels: ['" +
                                        AntwortzeitLabel.SehrSchnell + "', '" +
                                        AntwortzeitLabel.Schnell + "', '" +
                                        AntwortzeitLabel.Normal + "', '" +
                                        AntwortzeitLabel.Langsam + "', '" +
                                        AntwortzeitLabel.SehrLangsam +
                                        "'], " +
                                "type: 'pie', " +
                                "marker: {" +
                                    "colors: ['" +
                                        AntwortzeitFarbe.SehrSchnell + "', '" +
                                        AntwortzeitFarbe.Schnell + "', '" +
                                        AntwortzeitFarbe.Normal + "', '" +
                                        AntwortzeitFarbe.Langsam + "', '" +
                                        AntwortzeitFarbe.SehrLangsam +
                                        "']" +
                                "}" +
                            "}]";
                            layout = "{title: 'Prozentuale Verteilung deiner Antwortzeiten'}";
                            writer.WriteLineNoTabs("Plotly.newPlot('" + "kuchendiagramm" + i + "', " + data + "," + layout + "); ");

                            /*
                            * Antwortzeiten nach Wochentagesabhängigkeit per Boxplot
                            */
                            data = "[" + string_tagesabhaengigkeit + "]";
                            layout = "{title: 'Deine Antwortzeiten (in Stunden) nach Tag des E-Mail-Eingangs'}";
                            writer.WriteLineNoTabs("Plotly.newPlot('" + "tagesabhaengikeit" + i + "', " + data + "," + layout + "); ");

                            /*
                            * Antwortzeiten nach Tageszeitabhängigkeit per Boxplot
                            */
                            data = "[" + string_tageszeitabhaengigkeit + "]";

                            layout = "{title: 'Deine Antwortzeiten (in Stunden) nach Tageszeit des E-Mail-Eingangs'}";
                            writer.WriteLineNoTabs("Plotly.newPlot('" + "tageszeitabhaengikeit" + i + "', " + data + "," + layout + "); ");

                        }
                    }

                }
                writer.RenderEndTag(); // Javascript

                writer.RenderEndTag(); // body
                writer.RenderEndTag(); // html
            }


            // Return the result.
            File.WriteAllText(Pfad + "SynchronizitaetsUntersuchung.html", stringWriter.ToString());
        }

        private void Daten_In_Datei_Schreiben()
        {
            StreamWriter file = new StreamWriter(Pfad + "SynchronizitaetsUntersuchung.json");
            StringBuilder sb = new StringBuilder();
            JsonWriter jw = new JsonTextWriter(new StringWriter(sb));
            jw.Formatting = Formatting.Indented;

            jw.WriteStartObject();

            jw.WritePropertyName("speicher");
            jw.WriteStartArray();
            foreach(KeyValuePair<string,int> pair in speicher)
            {
                jw.WriteStartObject();
                jw.WritePropertyName("folder");
                jw.WriteValue(pair.Key);
                jw.WritePropertyName("lesezeichen");
                jw.WriteValue(pair.Value);
                jw.WriteEndObject();
            }
            jw.WriteEndArray();

            jw.WritePropertyName("beziehungen");
            jw.WriteStartArray();
            foreach (KeyValuePair<string, Beziehung> beziehung in beziehungen)
            {
                jw.WriteStartObject();

                jw.WritePropertyName("partner");
                jw.WriteValue(beziehung.Key);

                jw.WritePropertyName("antworten");
                jw.WriteStartArray();
                foreach (Antwort antwort in beziehung.Value.Antworten)
                {
                    jw.WriteStartObject();
                    jw.WritePropertyName("antwortzeit");
                    jw.WriteValue(antwort.Antwortzeit);
                    jw.WritePropertyName("tageszeit");
                    jw.WriteValue(antwort.Tageszeit);
                    jw.WritePropertyName("wochentag");
                    jw.WriteValue(antwort.Wochentag);
                    jw.WriteEndObject();
                }
                jw.WriteEndArray();

                jw.WritePropertyName("konversationen");
                jw.WriteStartArray();
                foreach (KeyValuePair<string, Konversation> konversation in beziehung.Value.Konversationen)
                {
                    jw.WriteStartObject();
                    jw.WritePropertyName("id");
                    jw.WriteValue(konversation.Key);

                    jw.WritePropertyName("antwortzeiten");
                    jw.WriteStartArray();
                    foreach (double antwortzeit in konversation.Value.Antwortzeiten)
                    {
                        jw.WriteStartObject();
                        jw.WritePropertyName("antwortzeit");
                        jw.WriteValue(antwortzeit);
                        jw.WriteEndObject();
                    }
                    jw.WriteEndArray();
                    jw.WriteEndObject();
                }
                jw.WriteEndArray();
                jw.WriteEndObject();
            }
            jw.WriteEndArray();
            jw.WriteEndObject();

            file.WriteLine(sb.ToString());
            file.Close();
        }

        private void log(String message) 
        {
            if(logfile == null)
            {
                logfile = File.AppendText(Pfad + "logfile.txt");
            }
            
            logfile.WriteLine("{0}: {1}", DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture), message);
            Debug.WriteLine("{0}: {1}", DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture), message);
        }
    }
}