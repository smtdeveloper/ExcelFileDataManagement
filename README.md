# ExcelFileDataManagement


<H1> SMTcoder </H1>
Projeye yıldız Vermeyi Unutmayın 🚀 Teşekkürler! ❤️
<h3> Ekran görüntüleri </h3>

![bandicam 2021-05-18 04-59-22-414](https://user-images.githubusercontent.com/74311713/118579436-c3a0ba00-b796-11eb-86fa-9320cb47cb6e.jpg)
![bandicam 2021-05-18 05-02-58-932](https://user-images.githubusercontent.com/74311713/118579443-c6031400-b796-11eb-9528-0299f33b89ee.jpg)
![bandicam 2021-05-18 05-04-09-025](https://user-images.githubusercontent.com/74311713/118579449-c8656e00-b796-11eb-8bd1-1c2e63442c4c.jpg)

<br>
<br>

<h2> Nasıl Yapılır - C# ile Excel Dosyasına Bağlanma (OleDbConnection ile) </h2>


 OledbConnection kullanarak Excel dosyasına bağlanıp Select (Veri çekme), İnsert (Veri Ekleme), Update (Güncelleme) işlemlerini gerçekleştireceğiz ve Excel dosyasındaki verilerin Datagridview de görüntülenmesini sağlayacağız.



Örneğimizde D sürücüsünde bulunan ve data listesi tutan “data.xlsx” isimli bir excel dosyasına bağlanıp bu işlemleri gerçekleştireceğiz. Excel dosyamızı aşağıdaki şekilde hazırlıyoruz.


Daha sonra formumuzu aşağıdaki şekilde tasarlayalım.

![bandicam 2021-05-18 04-59-22-414](https://user-images.githubusercontent.com/74311713/118579436-c3a0ba00-b796-11eb-86fa-9320cb47cb6e.jpg)
 



 
Kodlamaya başlayalım. İlk olarak bağlantı sağlayabilmek için;


<h3> using System.Data.OleDb; </h3>
ekliyoruz.

Daha sonra verileri Getir butonuna çift tıklayarak excel verilerimizin DataGridView üzerinde görünmesini sağlamak amacıyla aşağıdaki kodları yazıyoruz. OledbConnection bağlantı cümlesinde HDR= YES yaparak ilk satırın sütun başlığı olarak ayarlanmasını sağlıyoruz.

private void button1_Click(object sender, EventArgs e)
{
OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\ogrenci.xlsx; Extended Properties='Excel 12.0 xml;HDR=YES;'"); 
baglanti.Open();  
OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [Sayfa1$]", baglanti);
DataTable dt = new DataTable();
da.Fill(dt);
dataGridView1.DataSource = dt.DefaultView;
baglanti.Close();
}


private void button1_Click(object sender, EventArgs e)
{
OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\ogrenci.xlsx; Extended Properties='Excel 12.0 xml;HDR=YES;'"); 
baglanti.Open(); 
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
komut.Connection = baglanti; 
string sql = "Insert into [Sayfa1$] (NUMARA,AD,SOYAD,SINIFI) values('" + textBox1.Text + "','" + textBox2.Text + "','"+textBox3.Text+"','"+textBox4.Text+"')";
komut.CommandText = sql;
komut.ExecuteNonQuery();
baglanti.Close();


OleDbCommand komut = new OleDbCommand();
OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\ogrenci.xlsx; Extended Properties='Excel 12.0 xml;HDR=YES;'");
baglanti.Open();
komut.Connection = baglanti; 
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
komut.CommandText = sql; 
komut.ExecuteNonQuery();
baglanti.Close();


OleDbCommand komut = new OleDbCommand();
OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\ogrenci.xlsx; Extended Properties='Excel 12.0 xml;HDR=YES'");
baglanti.Open();
komut.Connection = baglanti;
string sql = "Update  [Sayfa1$] set AD='"+textBox2.Text+"',SOYAD='"+textBox3.Text+"',SINIFI='"+textBox4.Text+"' WHERE NUMARA="+textBox1.Text+"";
komut.CommandText = sql; 
komut.ExecuteNonQuery();
baglanti.Close();
Yukarıdaki örnekte bağlantıyı tekrar tekrar yazmak yerine Public olarak tanımlayabilirsiniz. İsterseniz veri seçme için bir metot oluşturarak Güncelleme ve Ekleme işlemlerinden sonra veya form açıldığında datagrid’ in güncellenmesini sağlayabilirsiniz.




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
