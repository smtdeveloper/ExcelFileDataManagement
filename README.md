# ExcelFileDataManagement


<H1> SMTcoder </H1>

<h3> Ekran görüntüleri </h3>

![bandicam 2021-05-18 04-59-22-414](https://user-images.githubusercontent.com/74311713/118579436-c3a0ba00-b796-11eb-86fa-9320cb47cb6e.jpg)
![bandicam 2021-05-18 05-02-58-932](https://user-images.githubusercontent.com/74311713/118579443-c6031400-b796-11eb-9528-0299f33b89ee.jpg)
![bandicam 2021-05-18 05-04-09-025](https://user-images.githubusercontent.com/74311713/118579449-c8656e00-b796-11eb-8bd1-1c2e63442c4c.jpg)

<br>
<br>

C# ile Excel Dosyasına Bağlanma (OleDbConnection ile)
3 sene önce54 Yorum
Bu yazımızda OledbConnection kullanarak Excel dosyasına bağlanıp Select (Veri çekme), İnsert (Veri Ekleme), Update (Güncelleme) işlemlerini gerçekleştireceğiz ve Excel dosyasındaki verilerin Datagridview de görüntülenmesini sağlayacağız.


 
Ayrıca C# dilinde yazılmış daha fazla örnek ve konular için C# Dersleri yazısını yada sağ üstte bulunan site içinde arama panelini kullanabilirsiniz.


Örneğimizde D sürücüsünde bulunan ve Öğrenci listesi tutan “ogrenci.xlsx” isimli bir excel dosyasına bağlanıp bu işlemleri gerçekleştireceğiz. Excel dosyamızı aşağıdaki şekilde hazırlıyoruz.


Daha sonra formumuzu aşağıdaki şekilde tasarlayalım.


 
excel_datagrid


 
Kodlamaya başlayalım. İlk olarak bağlantı sağlayabilmek için;

using System.Data.OleDb;
1
using System.Data.OleDb;
ekliyoruz.

Daha sonra verileri Getir butonuna çift tıklayarak excel verilerimizin DataGridView üzerinde görünmesini sağlamak amacıyla aşağıdaki kodları yazıyoruz. OledbConnection bağlantı cümlesinde HDR= YES yaparak ilk satırın sütun başlığı olarak ayarlanmasını sağlıyoruz.

private void button1_Click(object sender, EventArgs e)
{
OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\ogrenci.xlsx; Extended Properties='Excel 12.0 xml;HDR=YES;'"); 
baglanti.Open();  //www.yazilimkodlama.com
OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [Sayfa1$]", baglanti);
DataTable dt = new DataTable();
da.Fill(dt);
dataGridView1.DataSource = dt.DefaultView;
baglanti.Close();
}
1
2
3
4
5
6
7
8
9
10
private void button1_Click(object sender, EventArgs e)
{
OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\ogrenci.xlsx; Extended Properties='Excel 12.0 xml;HDR=YES;'"); 
baglanti.Open();  //www.yazilimkodlama.com
OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [Sayfa1$]", baglanti);
DataTable dt = new DataTable();
da.Fill(dt);
dataGridView1.DataSource = dt.DefaultView;
baglanti.Close();
}

Ekle komutuna basınca Textbox’ lara girmiş olduğumuz değerlerin ilgili Excel sütunlarına kayıt işlemi için aşağıdaki kodları yazıyoruz.

OleDbCommand komut = new OleDbCommand();
OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\ogrenci.xlsx; Extended Properties='Excel 12.0 xml;HDR=YES;'");
baglanti.Open();
komut.Connection = baglanti; //www.yazilimkodlama.com
string sql = "Insert into [Sayfa1$] (NUMARA,AD,SOYAD,SINIFI) values('" + textBox1.Text + "','" + textBox2.Text + "','"+textBox3.Text+"','"+textBox4.Text+"')";
komut.CommandText = sql;
komut.ExecuteNonQuery();
baglanti.Close();
1
2
3
4
5
6
7
8
OleDbCommand komut = new OleDbCommand();
OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\ogrenci.xlsx; Extended Properties='Excel 12.0 xml;HDR=YES;'");
baglanti.Open();
komut.Connection = baglanti; //www.yazilimkodlama.com
string sql = "Insert into [Sayfa1$] (NUMARA,AD,SOYAD,SINIFI) values('" + textBox1.Text + "','" + textBox2.Text + "','"+textBox3.Text+"','"+textBox4.Text+"')";
komut.CommandText = sql;
komut.ExecuteNonQuery();
baglanti.Close();
Güncelleme işlemini TextBox1′ e girdiğimiz Öğrenci Numarasına göre yapalım. Örneğin 155 nolu Öğrencinin bilgilerini değiştirmek gibi. Bunun için Güncelle butonuna aşağıdaki kodları yazabiliriz.

OleDbCommand komut = new OleDbCommand();
OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\ogrenci.xlsx; Extended Properties='Excel 12.0 xml;HDR=YES'");
baglanti.Open();
komut.Connection = baglanti;
string sql = "Update  [Sayfa1$] set AD='"+textBox2.Text+"',SOYAD='"+textBox3.Text+"',SINIFI='"+textBox4.Text+"' WHERE NUMARA="+textBox1.Text+"";
komut.CommandText = sql; //www.yazilimkodlama.com
komut.ExecuteNonQuery();
baglanti.Close();
1
2
3
4
5
6
7
8
OleDbCommand komut = new OleDbCommand();
OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\ogrenci.xlsx; Extended Properties='Excel 12.0 xml;HDR=YES'");
baglanti.Open();
komut.Connection = baglanti;
string sql = "Update  [Sayfa1$] set AD='"+textBox2.Text+"',SOYAD='"+textBox3.Text+"',SINIFI='"+textBox4.Text+"' WHERE NUMARA="+textBox1.Text+"";
komut.CommandText = sql; //www.yazilimkodlama.com
komut.ExecuteNonQuery();
baglanti.Close();
Yukarıdaki örnekte bağlantıyı tekrar tekrar yazmak yerine Public olarak tanımlayabilirsiniz. İsterseniz veri seçme için bir metot oluşturarak Güncelleme ve Ekleme işlemlerinden sonra veya form açıldığında datagrid’ in güncellenmesini sağlayabilirsiniz.

Kodların bu şekilde düzenlenmiş hali ise aşağıdaki şekilde olacaktır.

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;


namespace csharp_excel_baglanti
{
public partial class Form1 : Form
{
public Form1()
{
InitializeComponent(); //www.yazilimkodlama.com
}
OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\ogrenci.xlsx; Extended Properties='Excel 12.0 xml;HDR=YES;'");
public void doldur()
{
baglanti.Open();
OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [Sayfa1$]", baglanti);
DataTable dt = new DataTable();
da.Fill(dt);
 dataGridView1.DataSource = dt.DefaultView;
baglanti.Close();
}
private void button1_Click(object sender, EventArgs e)
{

 doldur();

    }

private void button2_Click(object sender, EventArgs e)
{
OleDbCommand komut = new OleDbCommand();
baglanti.Open();
komut.Connection = baglanti;
string sql = "Insert into [Sayfa1$] (NUMARA,AD,SOYAD,SINIFI) values('" + textBox1.Text + "','" + textBox2.Text + "','"+textBox3.Text+"','"+textBox4.Text+"')";
komut.CommandText = sql;
komut.ExecuteNonQuery();
baglanti.Close();//www.yazilimkodlama.com
doldur();
}

private void button3_Click(object sender, EventArgs e)
{
OleDbCommand komut = new OleDbCommand();
baglanti.Open();
komut.Connection = baglanti;
string sql = "Update  [Sayfa1$] set AD='"+textBox2.Text+"',SOYAD='"+textBox3.Text+"',SINIFI='"+textBox4.Text+"' WHERE NUMARA="+textBox1.Text+"";
komut.CommandText = sql;
komut.ExecuteNonQuery();
baglanti.Close();//www.yazilimkodlama.com
doldur();
}

private void Form1_Load(object sender, EventArgs e)
{
doldur();
}
}
}
1
2
3
4
5
6
7
8
9
10
11
12
13
14
15
16
17
18
19
20
21
22
23
24
25
26
27
28
29
30
31
32
33
34
35
36
37
38
39
40
41
42
43
44
45
46
47
48
49
50
51
52
53
54
55
56
57
58
59
60
61
62
63
64
65
66
67
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
 
 
namespace csharp_excel_baglanti
{
public partial class Form1 : Form
{
public Form1()
{
InitializeComponent(); //www.yazilimkodlama.com
}
OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\ogrenci.xlsx; Extended Properties='Excel 12.0 xml;HDR=YES;'");
public void doldur()
{
baglanti.Open();
OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [Sayfa1$]", baglanti);
DataTable dt = new DataTable();
da.Fill(dt);
 dataGridView1.DataSource = dt.DefaultView;
baglanti.Close();
}
private void button1_Click(object sender, EventArgs e)
{
 
 doldur();
 
    }
 
private void button2_Click(object sender, EventArgs e)
{
OleDbCommand komut = new OleDbCommand();
baglanti.Open();
komut.Connection = baglanti;
string sql = "Insert into [Sayfa1$] (NUMARA,AD,SOYAD,SINIFI) values('" + textBox1.Text + "','" + textBox2.Text + "','"+textBox3.Text+"','"+textBox4.Text+"')";
komut.CommandText = sql;
komut.ExecuteNonQuery();
baglanti.Close();//www.yazilimkodlama.com
doldur();
}
 
private void button3_Click(object sender, EventArgs e)
{
OleDbCommand komut = new OleDbCommand();
baglanti.Open();
komut.Connection = baglanti;
string sql = "Update  [Sayfa1$] set AD='"+textBox2.Text+"',SOYAD='"+textBox3.Text+"',SINIFI='"+textBox4.Text+"' WHERE NUMARA="+textBox1.Text+"";
komut.CommandText = sql;
komut.ExecuteNonQuery();
baglanti.Close();//www.yazilimkodlama.com
doldur();
}
 
private void Form1_Load(object sender, EventArgs e)
{
doldur();
}
}
}
Kaynak:

www.csharp-console-examples.com


 <h3> <a href="https://sametakca.com/">  web sitem </a> </h3> 
 
<br> <br>
<h3> Sosyal Medya Hesaplarım 😛 </h3>
<br>

<a href="https://www.instagram.com/smtcoder/">
instagram - @SMTcoder 
</a>
<br>

<a href="https://www.linkedin.com/in/samet-akca-2a4bbb1a8/">
linkedin
</a>
<br>

<a href="https://www.youtube.com/channel/UCZXmqpZJ3ax5Uzm0pXeVqMg">
youtube
</a>

<br>

<br> <br>
<h3> Projelerim 😛 </h3>
<br>

<a href="https://play.google.com/store/apps/developer?id=Samet+Akca&gl=TR">
Google Play uygulamalarım
</a>
<br>
<a href="https://www.tabbs.co/Samet">
 Tüm Projeler 
</a>


<br>
<br>


Projeye yıldız Vermeyi Unutmayın  🚀
Teşekkürler! ❤️
