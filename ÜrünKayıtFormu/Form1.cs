using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ÜrünKayıtFormu
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int barkodDurum = ListeBarkodDenetle(listView1, textBox1.Text);
            if (textBox1.Text != " ")
            {
                string[] elemanlar = { textBox1.Text, comboBox1.Text, textBox2.Text, textBox3.Text };
                ListViewItem veriler = new ListViewItem(elemanlar);

                listView1.Items.Add(veriler);
                Temizle();
            }

            else
            {
                MessageBox.Show("Geçersiz barkod girişi!");
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            listView1.FullRowSelect = true; //tıklayınca satırın hepsini seçme.
            listView2.FullRowSelect = true;

            ListViewKolonlar(listView1);
            ListViewKolonlar(listView2);
        }

        void ListViewKolonlar(ListView Liste)
        {
            Liste.Columns.Add("Barkod");
            Liste.Columns.Add("Kategori");
            Liste.Columns.Add("Ürün Adı");
            Liste.Columns.Add("Fiyat");
        }

        int ListeBarkodDenetle(ListView Liste, string Barkod)
        {
            foreach (ListViewItem veri in Liste.Items)
            {

                if (Barkod == veri.Text)
                {
                    return 1; //eklenmek istenen barkod listede var.
                }
            }

              return 0; //eklenmek istenen barkod listede yok.

        }

        void Temizle()
        {
            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            comboBox1.Text = " ";
        }

        ListViewItem IptalListe;
        int i = 2;
        int j = 1;

        void ExceleAktar(ListView Liste)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application uygulama = new Microsoft.Office.Interop.Excel.Application();
                uygulama.Visible = true;
                Workbook kitap = uygulama.Workbooks.Add(XlSheetType.xlWorksheet);
                Worksheet sayfa = (Worksheet)uygulama.ActiveSheet;
                sayfa.Cells[1, 1] = "Barkod";
                sayfa.Cells[1, 2] = "Kategori";
                sayfa.Cells[1, 3] = "Ürün Adı";
                sayfa.Cells[1, 4] = "Fiyat";

                foreach(ListViewItem item in Liste.Items)
                {
                    sayfa.Cells[i, j] = item.Text.ToString();
                    foreach(ListViewItem.ListViewSubItem sb in item.SubItems) //listViewlere eklemek için.
                    {
                        sayfa.Cells[i, j] = sb.Text.ToString();
                        j++;

                    }
                    j = 1;
                    i++;
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Uygulamada hata var.");
                
            }
        }

        private void verileriGösterToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ExceleAktar(listView1);
        }

        private void silToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach(ListViewItem veri in listView1.SelectedItems)
            {
                string[] İptalEdilenler = { veri.SubItems[0].Text, veri.SubItems[1].Text, veri.SubItems[2].Text, veri.SubItems[3].Text };

                listView2.Items.Add(IptalListe = new ListViewItem(İptalEdilenler));
               
            }
            listView1.SelectedItems[0].Remove();
        }
    }
}
