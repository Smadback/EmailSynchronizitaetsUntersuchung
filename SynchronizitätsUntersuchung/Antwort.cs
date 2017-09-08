using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SynchronizitätsUntersuchung
{
    class Antwort
    {

        public double Antwortzeit;
        public DayOfWeek Wochentag;
        public string Tageszeit;
        public Antwortzeit Synchronizität { get; private set; }

        private DateTime zeitstempel_empfangene_mail;
        private DateTime zeitstempel_gesendete_antwort;

        public Antwort(double antwortzeit, DateTime zeitstempel_empfangene_mail, DateTime zeitstempel_gesendete_antwort)
        {
            Antwortzeit = antwortzeit;
            Wochentag = zeitstempel_empfangene_mail.DayOfWeek;
            Tageszeit = Tageszeitabhaengigkeit.GetTageszeit(zeitstempel_empfangene_mail.TimeOfDay.Hours);

            Synchronizität = Helper.GetAntwortzeit(Antwortzeit);

            this.zeitstempel_empfangene_mail = zeitstempel_empfangene_mail;
            this.zeitstempel_gesendete_antwort = zeitstempel_gesendete_antwort;
        }

        public Antwort(double antwortzeit, string tageszeit, DayOfWeek wochentag)
        {
            Antwortzeit = antwortzeit;
            Wochentag = wochentag;
            Tageszeit = tageszeit;
            Synchronizität = Helper.GetAntwortzeit(Antwortzeit);
        }

        public void Synchronizitaet_Setzen()
        {
            Synchronizität = Helper.GetAntwortzeit(Antwortzeit);
        }

        public override string ToString()
        {
            return "[Antwort] nach " + TimeSpan.FromSeconds(Antwortzeit).ToString("%d'd '%h'h '%m'min '%s'sec'") + ". Empfangen am " + Wochentag + " " + zeitstempel_empfangene_mail + " zum " + Tageszeit + ". Antwort am " + zeitstempel_gesendete_antwort + ".";
        }

    }
}