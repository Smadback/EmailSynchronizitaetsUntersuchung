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

    enum Synchronizität
    {
        KomplettSynchron, GrößtenteilsSynchron, EherSynchron, GleichmäßigSynchronUndAsynchron, EherAsynchron, GrößtenteilsAsynchron, KomplettAsynchron, Undefiniert
    }

    class Helper
    {
        public static Antwortzeit GetAntwortzeit(double zeit)
        {
            if(zeit == 0)
            {
                return Antwortzeit.Undefiniert;
            }
            else if (zeit > 86400) // mehr als 24 Stunden
            {
                return Antwortzeit.SehrLangsam;
            }
            else if (zeit > 14400) // zwischen 4 Stunden und 24 Stunden
            {
                return Antwortzeit.Langsam;
            }
            else if (zeit > 3600) // zwischen 30 Minuten und 4 Stunden
            {
                return Antwortzeit.Normal;
            }
            else if (zeit > 240) // zwischen 3 Minuten und 30 Minuten
            {
                return Antwortzeit.Schnell;
            }
            else // weniger als 3 Minuten
            {
                return Antwortzeit.SehrSchnell;
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

    class SynchronizitätFarbe
    {
        public static readonly string KomplettSynchron = "#00b159";
        public static readonly string GrößtenteilsSynchron = "#00aedb";
        public static readonly string EherSynchron = "orange";
        public static readonly string GleichmäßigSynchronUndAsynchron = "#ffc425";
        public static readonly string EherAsynchron = "brown";
        public static readonly string GrößtenteilsAsynchron = "#f37735";
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

    class SynchronizitätLabel
    {
        public static readonly string KomplettSynchron = "Komplett synchron";
        public static readonly string GrößtenteilsSynchron = "Größtenteils synchron";
        public static readonly string EherSynchron = "Eher Synchron";
        public static readonly string GleichmäßigSynchronUndAsynchron = "Gleichmäßig synchron und asynchron";
        public static readonly string EherAsynchron = "Eher asynchron";
        public static readonly string GrößtenteilsAsynchron = "Größtenteils asynchron";
        public static readonly string KomplettAsynchron = "Komplett asynchron";
    }

    class Tageszeitabhängigkeit
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
