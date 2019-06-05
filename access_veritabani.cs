//DATAREADER
OleDbConnection baglan = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=KISIADRESLER.accdb;Persist Security Info = False;");

private void button1_Click(object sender, EventArgs e)
{
  baglan.Open();
  listBox1.Items.Clear();
  OleDbCommand komut = new OleDbCommand("Select * from Kisiler", baglan);
  OleDbDataReader oku = komut.ExecuteReader();
  while (oku.Read()) {    listBox1.Items.Add(oku["ID"]+"-"+ oku["AD"]+" "+ oku["SOYAD"]);     }
  baglan.Close();
}

//DATAREADER2
OleDbConnection baglan = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=KISIADRESLER.accdb;Persist Security Info = False;");
private void button1_Click(object sender, EventArgs e)
{
  baglan.Open();
  listBox1.Items.Clear();
  OleDbCommand komut = new OleDbCommand("Select * from Kisiler", baglan);
  OleDbDataReader oku = komut.ExecuteReader();
  while (oku.Read()) 
  {
    listBox1.Items.Add(
    oku["ID"] + "-" + 
    oku["AD"] + " " + 
    oku["SOYAD"]);
  }
  baglan.Close();
}

//DATAADAPTOR
OleDbConnection baglan = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=KISIADRESLER.accdb;Persist Security Info = False;");
private void button1_Click(object sender, EventArgs e)
{
  listBox1.Items.Clear();
  OleDbCommand komut = new OleDbCommand("Select * from Kisiler", baglan);
  OleDbDataAdapter adaptor = new OleDbDataAdapter(komut);
  DataTable tablo = new DataTable();
  adaptor.Fill(tablo);
  foreach (DataRow satır in tablo.Rows)
  {
    listBox1.Items.Add(satır["ID"] + "-" + satır["AD"] + " " + satır["SOYAD"]);
  }
}

//DATAGRIDVIEW
OleDbConnection baglan = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=KISIADRESLER.accdb;Persist Security Info = False;");        
private void button1_Click(object sender, EventArgs e)
{
  listBox1.Items.Clear();
  OleDbCommand komut = new OleDbCommand("Select * from Kisiler", baglan);
  OleDbDataAdapter adaptor = new OleDbDataAdapter(komut);

  DataTable tablo = new DataTable();
  adaptor.Fill(tablo);
  dataGridView1.DataSource = tablo;
}

//INSERT 1
OleDbConnection baglan = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=KISIADRESLER.accdb;Persist Security Info = False;");
private void button1_Click(object sender, EventArgs e)
{
  baglan.Open();
  OleDbCommand komut = new OleDbCommand("Insert into Kisiler(AD,SOYAD) Values('ATA','BAK')", baglan);
  komut.ExecuteNonQuery();//Komut çalıştırır kaç aded olduğunu da söyler
  baglan.Close();
}

//INSERT 2
OleDbConnection baglan = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=KISIADRESLER.accdb;Persist Security Info = False;");
private void button1_Click(object sender, EventArgs e)
{
  baglan.Open();                  
  OleDbCommand komut = new OleDbCommand("Insert into Kisiler(AD,SOYAD) Values(?,?)", baglan);
  komut.Parameters.AddWithValue("?", "HURŞİT");//Sırasıyla ekler
  komut.Parameters.AddWithValue("?", "ŞEKERPARE");

  komut.ExecuteNonQuery();//Komut çalıştırır kaç adet olduğunu da söyler
  baglan.Close();
}

//INSERT 3
OleDbConnection baglan = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=KISIADRESLER.accdb;Persist Security Info = False;");
private void button1_Click(object sender, EventArgs e)
{
  baglan.Open();
  listBox1.Items.Clear();
  OleDbCommand komut = new OleDbCommand("Insert into Kisiler(AD,SOYAD) Values(@AD,@SOYAD)", baglan);
  komut.Parameters.AddWithValue("@SOYAD", "ŞEKERPARE"); //kayıtların sırası farklı olursa sonuç farklı olur
  komut.Parameters.AddWithValue("@AD", "HURŞİT");//şekerpare hurşit diye kaydeder.
            
  komut.ExecuteNonQuery();//Komut çalıştırır kaç adet olduğunu da söyler
  baglan.Close();
}

//INSERT 4
OleDbConnection baglan = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=KISIADRESLER.accdb;Persist Security Info = False;");
private void button1_Click(object sender, EventArgs e)
{
  baglan.Open();
  listBox1.Items.Clear();
  OleDbCommand komut = new OleDbCommand("Insert into Kisiler(AD,SOYAD) Values('ATA','BAK')", baglan);
  komut.ExecuteNonQuery();//Komut çalıştırır kaç aded olduğunu da söyler

  komut = new OleDbCommand("SELECT AD,SOYAD FROM Kisiler", baglan);
  OleDbDataReader oku = komut.ExecuteReader();
  while (oku.Read()) { listBox1.Items.Add(oku["AD"]+"-"+ oku["SOYAD"]); }
  baglan.Close();
}

//DELETE
OleDbConnection baglan = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=KISIADRESLER.accdb;Persist Security Info = False;");
private void button1_Click(object sender, EventArgs e)
{
  baglan.Open();

  OleDbCommand komut = new OleDbCommand("delete from Kisiler where ID=?", baglan);
  komut.Parameters.AddWithValue("?", 35);//önemli nokta id yazamazsın ID olacak
  komut.ExecuteNonQuery();

  baglan.Close();
}

//UPDATE
OleDbConnection baglan = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=KISIADRESLER.accdb;Persist Security Info = False;");
private void button1_Click(object sender, EventArgs e)
{
  baglan.Open();

  OleDbCommand komut = new OleDbCommand("UPDATE Kisiler SET AD=? , SOYAD=? where ID=?", baglan);
  komut.Parameters.AddWithValue("?", "GÜL");
  komut.Parameters.AddWithValue("?", "ŞEN");
  komut.Parameters.AddWithValue("?", 3);

  komut.ExecuteNonQuery();
  baglan.Close();
}

//UPDATE2(SEÇİLİ)
OleDbConnection baglan = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=KISIADRESLER.accdb;Persist Security Info = False;");
private void button1_Click(object sender, EventArgs e)
{
  string seciliKayit = listBox1.SelectedItem.ToString();
  int seciliKayitId = int.Parse(seciliKayit.Split('-')[0]);
  baglan.Open();

  OleDbCommand komut = new OleDbCommand("UPDATE Kisiler SET AD=? , SOYAD=? where ID=?", baglan);
  komut.Parameters.AddWithValue("?", "GÜL");
  komut.Parameters.AddWithValue("?", "ŞEN");
  komut.Parameters.AddWithValue("?", seciliKayitId);

  komut.ExecuteNonQuery();
  baglan.Close();
}
