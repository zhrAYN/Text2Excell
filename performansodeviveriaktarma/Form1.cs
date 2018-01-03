using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel= Microsoft.Office.Interop.Excel;

namespace CsharpExcelVeriAktarma
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void listviewDoldur()
        {
            try
            {
                // 20 tane ürünü dinamik olarak eklemsi için kulandık. artırıp azaltabilirsiniz.
                int urunSayisi = 20;
                // For dönüsü ile ürün sayısı kadar ekleme işlemi yapıyoruz
                for (int i = 0; i < urunSayisi; i++)
                {
                    // eklenecek ürünler için örnek bir dizi tanımlıyoruz
                    string[] urunler = {
                                       "UrunID"+i.ToString(),
                                       "Urun Adı " + i.ToString(),
                                       "Adet :" + i.ToString()
                                        };
                    ListViewItem urun = new ListViewItem(urunler);
                    listView1.Items.Add(urun);// listview1 adında olan listview nesnemize urun itemini yüklüyoruz
                } // for bitis
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ürün ekleme işlemi sırasında bir hata meydana geldi.\n Hata içeriği:" + ex.ToString());
            }
        }

        /// <summary>
        ///  ListView verilerini dinamik olarak excel dosyasına aktarır
        /// </summary>
        /// <param name="lw">Aktarım Yapılacak ListView nesnesinin IDsi</param>
        public void excelAktar(ListView lw, ProgressBar pb = null)
        {
            try
            {
               
                Excel.Application Excelapp = new Excel.Application();
                Excel.Workbook wb = Excelapp.Workbooks.Add(Excel.XlSheetType.xlWorksheet);
                Excel.Worksheet ws = (Excel.Worksheet)Excelapp.ActiveSheet; // çalışma alanı aktif çalışma alanı

                Excelapp.Visible = true; // görünürlük aktif
                #region manuelBaslikalani
                // alanları manuel olarak yazıyoruz
                /*ws.Cells[1, 1] = "Stok Kodu";6
                ws.Cells[1, 2] = "Barkod";
                ws.Cells[1, 3] = "Stok Adı";
                ws.Cells[1, 4] = "Stok Miktar";
                ws.Cells[1, 4] = "Stok Grubu";
                ws.Cells[1, 5] = "Stok Markası";
                */
                #endregion
                // Eğer Progress bar nesnesi null değil ise sıfırlama ve ayarlama işlemini gerçekleştir
                if (pb != null)
                {
                    pb.Maximum = Convert.ToInt32(lw.Items.Count.ToString());
                    pb.Value = 0;
                }

                // Şimdi ise dinamik olarak colon bilgilerini alıyoruz ekliyoruz
                for (int i = 0; i < lw.Columns.Count; i++)
                {
                    // alanları manuel olarak yazıyoruz
                    ws.Cells[1, i + 1] = lw.Columns[i].Text.ToString();
                }
                // Şimdi de lw içerisindeki verileri dinamik olarak aktarıyoruz
                int _i = 2; // 2. satırdan itibaren içerikleri doldurmaya başla
                int j = 1;
                foreach (ListViewItem item in lw.Items)
                {
                    ws.Cells[_i, j] = item.Text.ToString();
                    foreach (ListViewItem.ListViewSubItem subitem in item.SubItems)
                    {
                        ws.Cells[_i, j] = subitem.Text.ToString();
                        j++;
                    }
                    j = 1;
                    _i++;
                    // Eğer Progress bar nesnesi null değil ise artırma işlemini yap
                    if (pb != null)
                    {
                        pb.Value = _i - 2;
                    }

                }
                Excelapp.Columns.AutoFit(); // column sütunları yazı boyutuna göre ayarlıyor
                Excelapp.AlertBeforeOverwriting = false; // aktarama işlemi sırasında alabileceğimiz hatalara karşı önlem olarak hata bastırma işlemi yapılıyor
            }
            catch (Exception)
            {

                throw;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            listviewDoldur();
        }

        private void btnAktar_Click(object sender, EventArgs e)
        {
            excelAktar(listView1);
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void btnAc_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "Belge Seçiniz...";
            openFileDialog1.Filter = "Excel 2003 dosyaları (*.xls)|*.xls|Excel 2010-2016 dosyaları (*.xlsx)|*.xlsx";
            openFileDialog1.FilterIndex = 1;
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel) 
            MessageBox.Show("Dosya Seçilemedi");
            else 
            {
                Excel.Application Excelapp = new Excel.Application();
                Excelapp.Workbooks.Open(openFileDialog1.FileName);
                Excelapp.Visible = true;
            }
                        
        }
        //Excele ekranına ve dosyamıza ekleme yapmak için kod satırlarımız...                                     
        private void button1_Click(object sender, EventArgs e)
        {
            string  UrunID    = textBox1.Text;
            string  UrunAdı   = textBox2.Text;
            string  Adet      = textBox3.Text;
            string Başlıkekle = textBox4.Text;

            string[] row = { "UrunID" +UrunID, "Urun Adı " +UrunAdı, "Adet:" +Adet,  };
            var satir = new ListViewItem(row);
            listView1.Items.Add(satir);

        }

        private void listView1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
           
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            textBox4.Clear();

        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            
            string aktar = textBox4.Text;
            listView1.Columns.Add(aktar , 100);


        }

        private void button4_Click(object sender, EventArgs e)
        {

        }
    }
}
