using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SynchronizitätsUntersuchung
{
    enum Antwortzeit
    {
        SehrSchnell, Schnell, Normal, Langsam, SehrLangsam, Undefiniert
    }

    enum Synchronizitaet
    {
        KomplettSynchron, GroesstenteilsSynchron, EherSynchron, GleichmaeßigSynchronUndAsynchron, EherAsynchron, GroesstenteilsAsynchron, KomplettAsynchron, Undefiniert
    }

    class Helper
    {
        /**
         * zeit in Sekunden
         */
        public static Antwortzeit GetAntwortzeit(double zeit)
        {
            // Wenn Breite Werte ausgewählt wurde
            if(SynchronizitaetUntersuchung.Breite_Werte)
            {
                if (zeit == 0)
                {
                    return Antwortzeit.Undefiniert;
                }
                else if (zeit > 295200)
                {
                    return Antwortzeit.SehrLangsam; // mehr als 3 Tage
                }
                else if (zeit > 86400)
                {
                    return Antwortzeit.Langsam; // zwischen 1 Tag und 3 Tage
                }
                else if (zeit > 3600)
                {
                    return Antwortzeit.Normal; // zwischen 1 Stunden und 1 Tag
                }
                else if (zeit > 300)
                {
                    return Antwortzeit.Schnell; // zwischen 5 Minuten und 1 Stunde
                }
                else
                {
                    return Antwortzeit.SehrSchnell; // weniger als 5 Minuten
                }
            }
            // Wenn nicht Breite Werte ausgewählt wurde, und stattdessen statistische Werte verwendet werden sollenr
            else
            {
                if (zeit == 0)
                {
                    return Antwortzeit.Undefiniert;
                }
                else if (zeit > 86400) 
                {
                    return Antwortzeit.SehrLangsam; // mehr als 24 Stunden
                }
                else if (zeit > 14400) 
                {
                    return Antwortzeit.Langsam; // zwischen 4 Stunden und 24 Stunden
                }
                else if (zeit > 1800)
                {
                    return Antwortzeit.Normal; // zwischen 30 Minuten und 4 Stunden
                }
                else if (zeit > 180)
                {
                    return Antwortzeit.Schnell; // zwischen 3 Minuten und 30 Minuten
                }
                else
                {
                    return Antwortzeit.SehrSchnell; // weniger als 3 Minuten
                }
            }
        }
    }

    class AntwortzeitFarbe
    {
        public static readonly string SehrSchnell = "#00b159";
        public static readonly string Schnell = "#00aedb";
        public static readonly string Normal = "#ffc425";
        public static readonly string Langsam = "#f37735";
        public static readonly string SehrLangsam = "#d11141";
    }

    class SynchronizitaetFarbe
    {
        public static readonly string KomplettSynchron = "#00b159";
        public static readonly string GroesstenteilsSynchron = "#00aedb";
        public static readonly string EherSynchron = "orange";
        public static readonly string GleichmaeßigSynchronUndAsynchron = "#ffc425";
        public static readonly string EherAsynchron = "brown";
        public static readonly string GroesstenteilsAsynchron = "#f37735";
        public static readonly string KomplettAsynchron = "#d11141";
    }

    class AntwortzeitLabel
    {
        public static readonly string SehrSchnell = "Sehr schnell";
        public static readonly string Schnell = "Schnell";
        public static readonly string Normal = "Normal";
        public static readonly string Langsam = "Langsam";
        public static readonly string SehrLangsam = "Sehr langsam";
    }

    class SynchronizitaetLabel
    {
        public static readonly string KomplettSynchron = "Komplett synchron";
        public static readonly string GroesstenteilsSynchron = "Größtenteils synchron";
        public static readonly string EherSynchron = "Eher Synchron";
        public static readonly string GleichmaeßigSynchronUndAsynchron = "Gleichmäßig synchron und asynchron";
        public static readonly string EherAsynchron = "Eher asynchron";
        public static readonly string GroesstenteilsAsynchron = "Größtenteils asynchron";
        public static readonly string KomplettAsynchron = "Komplett asynchron";
    }

    class Tageszeitabhaengigkeit
    {
        public static readonly string Morgen = "Morgen";
        public static readonly string Vormittag = "Vormittag";
        public static readonly string Mittag = "Mittag";
        public static readonly string Nachmittag = "Nachmittag";
        public static readonly string Abend = "Abend";
        public static readonly string Nacht = "Nacht";

        public static string GetTageszeit (int stunden)
        {
            string tageszeit = "";

            if (stunden > 0 && stunden < 4)
            {
                tageszeit = "Nacht";
            }
            else if (stunden > 4 && stunden < 8)
            {
                tageszeit = "Morgen";
            }
            else if (stunden > 8 && stunden < 12)
            {
                tageszeit = "Vormittag";
            }
            else if (stunden > 12 && stunden < 16)
            {
                tageszeit = "Mittag";
            }
            else if (stunden > 16 && stunden < 20)
            {
                tageszeit = "Nachmittag";
            }
            else
            {
                tageszeit = "Abend";
            }
            return tageszeit;
        }
    }
}
