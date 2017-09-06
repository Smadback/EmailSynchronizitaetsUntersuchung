using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SynchronizitätsUntersuchung
{
    class Beziehung
    {

        /*
         * Properties einer Beziehung
         */
        public string Partner { get; private set; }

        public Dictionary<string, Konversation> Konversationen { get; private set; }
        public List<Antwort> Antworten;
        public int AnzahlErhalteneEmails;
        public Dictionary<Antwortzeit, int> VerteilungAntworten;
        public Dictionary<Synchronizitaet, int> VerteilungGespraeche;
        public Dictionary<DayOfWeek, List<double>> TagesabhaengigeAntworten;
        public Dictionary<string, List<double>> TageszeitabhaengigeAntworten;
        public int DurchschnittlicheKonversationslaenge;
        
        /*
         * Konstruktor
         */
        public Beziehung(string partner)
        {
            Partner = partner;
            Konversationen = new Dictionary<string, Konversation>();
            Antworten = new List<Antwort>();
            VerteilungAntworten = new Dictionary<Antwortzeit, int>();
            VerteilungGespraeche = new Dictionary<Synchronizitaet, int>();
            TagesabhaengigeAntworten = new Dictionary<DayOfWeek, List<double>>();
            TagesabhaengigeAntworten[DayOfWeek.Monday] = new List<double>();
            TagesabhaengigeAntworten[DayOfWeek.Tuesday] = new List<double>();
            TagesabhaengigeAntworten[DayOfWeek.Wednesday] = new List<double>();
            TagesabhaengigeAntworten[DayOfWeek.Thursday] = new List<double>();
            TagesabhaengigeAntworten[DayOfWeek.Friday] = new List<double>();
            TagesabhaengigeAntworten[DayOfWeek.Saturday] = new List<double>();
            TagesabhaengigeAntworten[DayOfWeek.Sunday] = new List<double>();

            TageszeitabhaengigeAntworten = new Dictionary<string, List<double>>();
            TageszeitabhaengigeAntworten[Tageszeitabhaengigkeit.Morgen] = new List<double>();
            TageszeitabhaengigeAntworten[Tageszeitabhaengigkeit.Vormittag] = new List<double>();
            TageszeitabhaengigeAntworten[Tageszeitabhaengigkeit.Mittag] = new List<double>();
            TageszeitabhaengigeAntworten[Tageszeitabhaengigkeit.Nachmittag] = new List<double>();
            TageszeitabhaengigeAntworten[Tageszeitabhaengigkeit.Abend] = new List<double>();
            TageszeitabhaengigeAntworten[Tageszeitabhaengigkeit.Nacht] = new List<double>();
        }

        /*
         * Füge eine neue Konversation der enstprechenden Beziehung hinzu
         */
        public void Konversation_Hinzufuegen(Konversation konversation)
        {
            if ( !Konversationen.ContainsKey( konversation.Id ) )
            {
                Konversationen[konversation.Id] = konversation;
            }
        }

        public void Beziehung_Auswerten()
        {
            VerteilungAntworten[Antwortzeit.SehrSchnell] = Antworten.Count(antwort => antwort.Synchronizität == Antwortzeit.SehrSchnell);
            VerteilungAntworten[Antwortzeit.Schnell] = Antworten.Count(antwort => antwort.Synchronizität == Antwortzeit.Schnell);
            VerteilungAntworten[Antwortzeit.Normal] = Antworten.Count(antwort => antwort.Synchronizität == Antwortzeit.Normal);
            VerteilungAntworten[Antwortzeit.Langsam] = Antworten.Count(antwort => antwort.Synchronizität == Antwortzeit.Langsam);
            VerteilungAntworten[Antwortzeit.SehrLangsam] = Antworten.Count(antwort => antwort.Synchronizität == Antwortzeit.SehrLangsam);

            foreach(Antwort antwort in Antworten)
            {
                TagesabhaengigeAntworten[antwort.Wochentag].Add(antwort.Antwortzeit);
                TageszeitabhaengigeAntworten[antwort.Tageszeit].Add(antwort.Antwortzeit);
            }

            // Alle Konversationen entfernen, die weniger als 2 Antworten haben, denn dann ist es kein Gespräch
            var toRemove = Konversationen.Where(pair => pair.Value.Antwortzeiten.Count() < 2)
                         .Select(pair => pair.Key)
                         .ToList();

            foreach (var key in toRemove)
            {
                Konversationen.Remove(key);
            }

            if (Konversationen.Count > 0)
            {
                int gesamt_länge = 0;
                foreach (KeyValuePair<string, Konversation> konv in Konversationen)
                {
                    konv.Value.Konversation_Auswerten();
                    gesamt_länge += konv.Value.Laenge;
                }

                DurchschnittlicheKonversationslaenge = gesamt_länge / Konversationen.Count;
            }

            VerteilungGespraeche[Synchronizitaet.KomplettSynchron] = Konversationen.Count(konversation => konversation.Value.Synchronizitaet == Synchronizitaet.KomplettSynchron);
            VerteilungGespraeche[Synchronizitaet.GroesstenteilsSynchron] = Konversationen.Count(konversation => konversation.Value.Synchronizitaet == Synchronizitaet.GroesstenteilsSynchron);
            VerteilungGespraeche[Synchronizitaet.EherSynchron] = Konversationen.Count(konversation => konversation.Value.Synchronizitaet == Synchronizitaet.EherSynchron);
            VerteilungGespraeche[Synchronizitaet.GleichmaeßigSynchronUndAsynchron] = Konversationen.Count(konversation => konversation.Value.Synchronizitaet == Synchronizitaet.GleichmaeßigSynchronUndAsynchron);
            VerteilungGespraeche[Synchronizitaet.EherAsynchron] = Konversationen.Count(konversation => konversation.Value.Synchronizitaet == Synchronizitaet.EherAsynchron);
            VerteilungGespraeche[Synchronizitaet.GroesstenteilsAsynchron] = Konversationen.Count(konversation => konversation.Value.Synchronizitaet == Synchronizitaet.GroesstenteilsAsynchron);
            VerteilungGespraeche[Synchronizitaet.KomplettAsynchron] = Konversationen.Count(konversation => konversation.Value.Synchronizitaet == Synchronizitaet.KomplettAsynchron);
        }
    }
}