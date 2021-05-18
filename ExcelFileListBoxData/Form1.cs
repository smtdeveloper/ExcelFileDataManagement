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

namespace ExcelFileListBoxData
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)//veri getir BTN
        {
            //SMTcoder
            OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\YAZILIM/data.xlsx; Extended Properties='Excel 12.0 xml;HDR=YES;'");
            baglanti.Open();  
            OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [Sayfa1$]", baglanti);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt.DefaultView;
            baglanti.Close();
        }

        private void button2_Click(object sender, EventArgs e)//Ekle BTN
        {
            //SMTcoder
            OleDbCommand komut = new OleDbCommand();
            OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\YAZILIM/data.xlsx; Extended Properties='Excel 12.0 xml;HDR=YES;'");
            baglanti.Open();
            komut.Connection = baglanti;
            string sql = "Insert into [Sayfa1$] (id,userId,name,lastName) values('" + textBox1.Text + "','" + textBox2.Text + "','" + textBox3.Text + "','" + textBox4.Text + "')";
            komut.CommandText = sql;
            komut.ExecuteNonQuery();
            baglanti.Close();

            button1.Click += new EventHandler(button1_Click);

            //SMTcoder
            OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [Sayfa1$]", baglanti);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt.DefaultView;
            baglanti.Close();

            textBox1.Clear();

            textBox2.Clear();

            textBox3.Clear();

            textBox4.Clear();
        }

        public void button3_Click(object sender, EventArgs e) //Güncelle BTN
        {
            //SMTcoder
            OleDbCommand komut = new OleDbCommand();
            OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\YAZILIM/data.xlsx; Extended Properties='Excel 12.0 xml;HDR=YES'");
            baglanti.Open();
            komut.Connection = baglanti;
            string sql = "Update  [Sayfa1$] set userId='" + textBox2.Text + "',name='" + textBox3.Text + "',lastName='" + textBox4.Text + "' WHERE id=" + textBox1.Text + "";
            komut.CommandText = sql; 
            komut.ExecuteNonQuery();
            baglanti.Close();

            //SMTcoder
            OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [Sayfa1$]", baglanti);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt.DefaultView;
            baglanti.Close();

            textBox1.Clear();

            textBox2.Clear();

            textBox3.Clear();

            textBox4.Clear();
           

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e) //Silme BTN
        {
            //SMTcoder
            OleDbCommand komut = new OleDbCommand();
            OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\YAZILIM/data.xlsx; Extended Properties='Excel 12.0 xml;HDR=YES'");
            baglanti.Open();
            komut.Connection = baglanti;
            string sql = "Delete from  [Sayfa1$] WHERE id=" + textBox1.Text + "";
            komut.CommandText = sql;
            komut.ExecuteNonQuery();
            baglanti.Close();

         

            //SMTcoder
            OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [Sayfa1$]", baglanti);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt.DefaultView;
            baglanti.Close();

            textBox1.Clear();
        }
    }
}
