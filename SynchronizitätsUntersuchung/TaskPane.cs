using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace SynchronizitätsUntersuchung
{
    public partial class TaskPane : UserControl
    {
        public TaskPane()
        {
            InitializeComponent();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string[] datei = File.ReadAllLines(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/SynchronizitaetsUntersuchung/Auswertung" + ".txt");
            int funktion = 1;
            Dictionary<int, Tuple<string, string>> die_synchronsten = new Dictionary<int, Tuple<string, string>>();

            foreach(string zeile in datei)
            {
                if(zeile == "" || zeile == null)
                {
                    continue;
                }

                if(zeile.Contains("/1"))
                {
                    funktion = 1;
                    continue;
                }
                else if (zeile.Contains("/2"))
                {
                    funktion = 2;
                    continue;
                }
                else  if (zeile.Contains("/3"))
                {
                    funktion = 3;
                    continue;
                }
                else if (zeile.Contains("/4"))
                {
                    funktion = 4;
                    continue;
                }
                else if (zeile.Contains("/5"))
                {
                    funktion = 5;
                    continue;
                }
                else if (zeile.Contains("/6"))
                {
                    funktion = 6;
                    continue;
                }

                if (funktion == 1 || funktion == 2)
                {
                    string position = zeile.Split('[', ']')[1];
                    string person = zeile.Split('"', '"')[1];
                    string synchronizität = zeile.Split('(', ')')[1];

                    //this.textBox1.Text += this.textBox1.Text + "\n" + position + ", " + person + ", " + synchronizität;

                    string[] row1 = { person, synchronizität };
                    listView1.Items.Add(position).SubItems.AddRange(row1);

                }
                
            }

        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
