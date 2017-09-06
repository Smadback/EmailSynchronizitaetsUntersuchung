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
        public Antwortzeit Synchronizität { get; }

        private DateTime zeitstempel_empfangene_mail;
        private DateTime zeitstempel_gesendete_antwort;

        public Antwort(double antwortzeit, DateTime zeitstempel_empfangene_mail, DateTime zeitstempel_gesendete_antwort)
        {
            Antwortzeit = antwortzeit;
            Wochentag = zeitstempel_empfangene_mail.DayOfWeek;
            Tageszeit = Tageszeitabhaengigkeit.GetTageszeit(zeitstempel_empfangene_mail.TimeOfDay.Hours);

            Synchronizität = getSynchronizität();

            this.zeitstempel_empfangene_mail = zeitstempel_empfangene_mail;
            this.zeitstempel_gesendete_antwort = zeitstempel_gesendete_antwort;
        }

        public Antwort(double antwortzeit, string tageszeit, DayOfWeek wochentag)
        {
            Antwortzeit = antwortzeit;
            Wochentag = wochentag;
            Tageszeit = tageszeit;
            Synchronizität = getSynchronizität();
        }


        private Antwortzeit getSynchronizität()
        {
            if (Antwortzeit > 86400) // mehr als 24 Stunden
            {
                return SynchronizitätsUntersuchung.Antwortzeit.SehrLangsam;
            }
            else if (Antwortzeit > 14400) // zwischen 4 Stunden und 24 Stunden
            {
                return SynchronizitätsUntersuchung.Antwortzeit.Langsam;
            }
            else if (Antwortzeit > 3600) // zwischen 30 Minuten und 4 Stunden
            {
                return SynchronizitätsUntersuchung.Antwortzeit.Normal;
            }
            else if (Antwortzeit > 240) // zwischen 5 Minuten und 30 Minuten
            {
                return SynchronizitätsUntersuchung.Antwortzeit.Schnell;
            }
            else // weniger als 3 Minuten
            {
                return SynchronizitätsUntersuchung.Antwortzeit.SehrSchnell;
            }
        }

        public override string ToString()
        {
            return "[Antwort] nach " + TimeSpan.FromSeconds(Antwortzeit).ToString("%d'd '%h'h '%m'min '%s'sec'") + ". Empfangen am " + Wochentag + " " + zeitstempel_empfangene_mail + " zum " + Tageszeit + ". Antwort am " + zeitstempel_gesendete_antwort + ".";
        }

    }
}
