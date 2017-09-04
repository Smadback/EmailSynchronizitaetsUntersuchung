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
        public Dictionary<Synchronizität, int> VerteilungGespräche;
        public Dictionary<DayOfWeek, List<double>> TagesabhängigeAntworten;
        public Dictionary<string, List<double>> TageszeitabhängigeAntworten;
        public int DurchschnittlicheKonversationslänge;
        
        /*
         * Konstruktor
         */
        public Beziehung(string partner)
        {
            Partner = partner;
            Konversationen = new Dictionary<string, Konversation>();
            Antworten = new List<Antwort>();
            VerteilungAntworten = new Dictionary<Antwortzeit, int>();
            VerteilungGespräche = new Dictionary<Synchronizität, int>();
            TagesabhängigeAntworten = new Dictionary<DayOfWeek, List<double>>();
            TagesabhängigeAntworten[DayOfWeek.Monday] = new List<double>();
            TagesabhängigeAntworten[DayOfWeek.Tuesday] = new List<double>();
            TagesabhängigeAntworten[DayOfWeek.Wednesday] = new List<double>();
            TagesabhängigeAntworten[DayOfWeek.Thursday] = new List<double>();
            TagesabhängigeAntworten[DayOfWeek.Friday] = new List<double>();
            TagesabhängigeAntworten[DayOfWeek.Saturday] = new List<double>();
            TagesabhängigeAntworten[DayOfWeek.Sunday] = new List<double>();

            TageszeitabhängigeAntworten = new Dictionary<string, List<double>>();
            TageszeitabhängigeAntworten[Tageszeitabhängigkeit.Morgen] = new List<double>();
            TageszeitabhängigeAntworten[Tageszeitabhängigkeit.Vormittag] = new List<double>();
            TageszeitabhängigeAntworten[Tageszeitabhängigkeit.Mittag] = new List<double>();
            TageszeitabhängigeAntworten[Tageszeitabhängigkeit.Nachmittag] = new List<double>();
            TageszeitabhängigeAntworten[Tageszeitabhängigkeit.Abend] = new List<double>();
            TageszeitabhängigeAntworten[Tageszeitabhängigkeit.Nacht] = new List<double>();
        }

        /*
         * Füge eine neue Konversation der enstprechenden Beziehung hinzu
         */
        public void Konversation_Hinzufügen(Konversation konversation)
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
                TagesabhängigeAntworten[antwort.Wochentag].Add(antwort.Antwortzeit);
                TageszeitabhängigeAntworten[antwort.Tageszeit].Add(antwort.Antwortzeit);
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
                    gesamt_länge += konv.Value.Länge;
                }

                DurchschnittlicheKonversationslänge = gesamt_länge / Konversationen.Count;
            }

            VerteilungGespräche[Synchronizität.KomplettSynchron] = Konversationen.Count(konversation => konversation.Value.Synchronizität == Synchronizität.KomplettSynchron);
            VerteilungGespräche[Synchronizität.GrößtenteilsSynchron] = Konversationen.Count(konversation => konversation.Value.Synchronizität == Synchronizität.GrößtenteilsSynchron);
            VerteilungGespräche[Synchronizität.EherSynchron] = Konversationen.Count(konversation => konversation.Value.Synchronizität == Synchronizität.EherSynchron);
            VerteilungGespräche[Synchronizität.GleichmäßigSynchronUndAsynchron] = Konversationen.Count(konversation => konversation.Value.Synchronizität == Synchronizität.GleichmäßigSynchronUndAsynchron);
            VerteilungGespräche[Synchronizität.EherAsynchron] = Konversationen.Count(konversation => konversation.Value.Synchronizität == Synchronizität.EherAsynchron);
            VerteilungGespräche[Synchronizität.GrößtenteilsAsynchron] = Konversationen.Count(konversation => konversation.Value.Synchronizität == Synchronizität.GrößtenteilsAsynchron);
            VerteilungGespräche[Synchronizität.KomplettAsynchron] = Konversationen.Count(konversation => konversation.Value.Synchronizität == Synchronizität.KomplettAsynchron);
        }
    }
}