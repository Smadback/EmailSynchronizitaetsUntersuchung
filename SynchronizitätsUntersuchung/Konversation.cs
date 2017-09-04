using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;
using System.Linq;

namespace SynchronizitätsUntersuchung
{
    class Konversation
    {
        // Properties einer Konversation
        public string Id { get; private set; }
        public List<double> Antwortzeiten { get; }

        // Wird berechnet
        public Dictionary<Antwortzeit, int> AnzahlAntwortzeiten { get; }
        public Antwortzeit MedianAntwortzeit { get; private set; }
        public Synchronizität Synchronizität { get; private set; }
        public int Länge;

        // irrelevant
        public string Thema;

        /*
         * Konstruktor
         */
        public Konversation(String id)
        {
            Id = id;
            Antwortzeiten = new List<double>();
            AnzahlAntwortzeiten = new Dictionary<Antwortzeit, int>();
        }

        /**
         * Neue Antwort der Konversation hinzufügen
         */
        public void Antwort_Hinzufügen(double antwortzeit)
        {
            // Die Antwortzeit in die Liste hinzufügen
            Antwortzeiten.Add(antwortzeit);
            // Anzahl der Antwortzeit um 1 erhöhen
            if(AnzahlAntwortzeiten.ContainsKey(Helper.GetAntwortzeit(antwortzeit)))
            {
                AnzahlAntwortzeiten[Helper.GetAntwortzeit(antwortzeit)]++;
            } else
            {
                AnzahlAntwortzeiten[Helper.GetAntwortzeit(antwortzeit)] = 1;
            }
        }

        /**
         * Wähle aus allen Antwortzeiten der Konversation den Median und setze diesen als Synchronizitätswert
         */
        public void Konversation_Auswerten()
        {
            Länge = Antwortzeiten.Count + 1;
            int länge = Antwortzeiten.Count;
            int halfIndex = länge / 2;
            int unteresQuartilIndex = (int)(länge * 0.25);
            int oberesQuartilIndex = (int)(länge * 0.75);
            // Alle Antwortzeiten absteigend sortieren, also von langsamen bis zu schnellen Antworten
            List<double> antwortzeiten_sortiert = (from antwort in Antwortzeiten orderby antwort descending select antwort).ToList();
            double median = 0;
            double unteres_quartil = 0;
            double oberes_quartil = 0;

            if (länge > 1)
            {
                // Median
                if ((länge % 2) == 0)
                {
                    // Median ist der Mittelwert der beiden Antwortzeiten in der Mitte der Liste (gibt nicht genau einen, da gerade Anzahl an Werten)
                    median = (Antwortzeiten[halfIndex] + Antwortzeiten[halfIndex - 1] / 2);
                }
                else
                {
                    // Median ist die Antwortzeit genau in der Mitte der Liste
                    median = Antwortzeiten[halfIndex];
                }

                // Quartile
                if ((länge % 4) == 0)
                {
                    // Unteres Quartil
                    unteres_quartil = (Antwortzeiten[unteresQuartilIndex] + Antwortzeiten[unteresQuartilIndex - 1] / 2);
                    // Oberes Quartil
                    oberes_quartil = (Antwortzeiten[oberesQuartilIndex] + Antwortzeiten[oberesQuartilIndex - 1] / 2);
                }
                else
                {
                    // Unteres Quartil
                    unteres_quartil = (Antwortzeiten[unteresQuartilIndex]);
                    // Oberes Quartil
                    oberes_quartil = (Antwortzeiten[oberesQuartilIndex]);
                }
            }

            MedianAntwortzeit = Helper.GetAntwortzeit(median);
            Antwortzeit UnteresQuartilAntwortzeit = Helper.GetAntwortzeit(unteres_quartil);
            Antwortzeit OberesQuartilAntwortzeit = Helper.GetAntwortzeit(oberes_quartil);

            // Abhängig vom Median und Quartilen speziellere Auswertung vornehmen
            if (MedianAntwortzeit == Antwortzeit.SehrSchnell || MedianAntwortzeit == Antwortzeit.Schnell)
            {
                if(Helper.GetAntwortzeit(Antwortzeiten[(int)(länge * 0.1)]) == Antwortzeit.Schnell || Helper.GetAntwortzeit(Antwortzeiten[(int)(länge * 0.1)]) == Antwortzeit.SehrSchnell)
                {
                    Synchronizität = Synchronizität.KomplettSynchron;
                }
                else if (UnteresQuartilAntwortzeit == Antwortzeit.Schnell || UnteresQuartilAntwortzeit == Antwortzeit.SehrSchnell)
                {
                    Synchronizität = Synchronizität.GrößtenteilsSynchron;
                }
                else
                {
                    // Median bei schnell: Gleichmäßig Synchron und Asynchron
                    Synchronizität = Synchronizität.EherSynchron;
                }
            }
            // Median bei normal: Gleichmäßig Synchron und Asynchron
            else if (MedianAntwortzeit == Antwortzeit.Normal)
            {
                Synchronizität = Synchronizität.GleichmäßigSynchronUndAsynchron;
            }
            else if (MedianAntwortzeit == Antwortzeit.Langsam || MedianAntwortzeit == Antwortzeit.SehrLangsam)
            {
                if (Helper.GetAntwortzeit(Antwortzeiten[(int)(länge * 0.9)]) == Antwortzeit.Langsam || Helper.GetAntwortzeit(Antwortzeiten[(int)(länge * 0.9)]) == Antwortzeit.SehrLangsam)
                {
                    Synchronizität = Synchronizität.KomplettAsynchron;
                }
                else if (OberesQuartilAntwortzeit == Antwortzeit.Langsam || OberesQuartilAntwortzeit == Antwortzeit.SehrLangsam)
                {
                    Synchronizität = Synchronizität.GrößtenteilsAsynchron;
                }
                else
                {
                    Synchronizität = Synchronizität.EherAsynchron;
                }
            } else
            {
                Synchronizität = Synchronizität.Undefiniert;
            }

        }
    }
}