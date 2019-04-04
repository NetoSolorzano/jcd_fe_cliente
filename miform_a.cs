using System;
using System.Windows.Forms;
using System.Drawing;
using System.Configuration;
using System.Data;
using System.Xml;
using System.Text;
using System.IO;
using System.IO.Compression;
using System.Timers;
using MySql.Data;
using MySql.Data.MySqlClient;

namespace viza_fact_elect_clt
{
	public class miform : Form
	{
        static string nomform = "miform";
        #region conexion a la base de datos
        // own database connection
        static string serv = ConfigurationManager.AppSettings["serv"].ToString();
        static string port = ConfigurationManager.AppSettings["port"].ToString();
        static string usua = ConfigurationManager.AppSettings["user"].ToString();
        static string cont = ConfigurationManager.AppSettings["pass"].ToString();
        static string data = ConfigurationManager.AppSettings["data"].ToString();
        static string ctl = ConfigurationManager.AppSettings["ConnectionLifeTime"].ToString();
        static string sped = ConfigurationManager.AppSettings["spedvta"].ToString();    // series electronicas
        static string tloa = ConfigurationManager.AppSettings["timeload"].ToString();   // milisegundos para jalar datos
        string DB_CONN_STR = "server=" + serv + ";port=" + port + ";uid=" + usua + ";pwd=" + cont + ";database=" + data + 
            ";ConnectionLifeTime=" + ctl + ";";
        #endregion
        #region declaracion de objetos
        // Objects
		private Panel marco1;       // marco para los datos de conexion a la base del cliente
		private Label lb_serv;      // direccion del servidor cliente
        private Label lb_base;      // nombre de la base de datos
        private Label lb_tabl;      // nombre de la tabla cabecera de docs.venta
        private Label lb_tabd;      // invoice detail table name
        private Label lb_usua;      // usuario de la base de datos cliente
        private Label lb_pass;      // contraseña base cliente
        private Label lb_port;      // puerto del servidor cliente
		private TextBox tx_serv;
        private TextBox tx_base;
        private TextBox tx_tabl;    // header invoice table
        private TextBox tx_tabd;    // detail invoice table
        private TextBox tx_usua;
        private TextBox tx_pass;
        private TextBox tx_port;
        //
        private Panel marco2;       // marco para los parametros de datos de la tabla del cliente
        private Label lb_letr;      // letras de factura electrónica
        private Label lb_digs;      // # digitos de la serie, incluyendo la letra, osea F001
        private Label lb_time;      // tiempo de espera para obtener los datos del cliente
        private TextBox tx_letr;
        private TextBox tx_digs;
        private TextBox tx_time;
        //
		private Button btn_;
        //
        private Panel marco3;           // marco para los filtros de entrada para jalar los datos propios
        private Label lb_ano;    // filtro de año
        private Label lb_mes;    // filtro de mes
        private Label lb_dia;    // filtro de día
        private Label lb_spe;       // serie del punto de emision
        private Label lb_usu;       // usuario que genera los documentos
        private Label lb_loc;       // local donde esta el usuario
        private TextBox tx_ano;
        private TextBox tx_mes;
        private TextBox tx_dia;
        private TextBox tx_spe;
        private TextBox tx_usu;
        private TextBox tx_loc;
        //
        private Panel marco4;           // marco para la grilla
        private DataGridView grilla;    // grilla de los datos actuales
        //
        private Panel marco5;           // botones de comando
        private Button bt_print;
        private Button bt_anu;

        #endregion
        #region declaracion de variables
        // main form defaults margins
        int ancho = 980;
        int largo = 550;
        // temporizador
        private System.Windows.Forms.Timer temporizador;
        // datos del cliente del sistema
        private string nomclie = "";
        private string rucclie = "";
        private string con_jaladatos;
        #endregion
        public miform ()                // load
		{
			Text = "Cliente de Interfase para Facturación Electrónica";
			this.Size = new Size (ancho, largo);
            this.MinimumSize = this.Size;
            this.MaximumSize = new Size(ancho + 300, largo + 200);
            this.FormClosing += miform_FormClosing;
            jalainfo();     // jalamos la info del cliente
			render ();
            init();
            jala_marco1();
            jala_marco2();
            jala_marco3();
            bt_anu.Visible = false;     // no anulamos nada de momento
            btn_.Focus();
		}
        private void init()
        {
            tx_ano.MaxLength = 4;
            tx_mes.MaxLength = 2;
            tx_dia.MaxLength = 2;
        }
		private void render()           // dibuja los objetos en la pantalla
		{
			marco1 = new Panel ();
            marco1.Location = new Point(this.Left, this.Top);
            marco1.Size = new Size(this.Width, 60);
            marco1.Anchor = (AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top);
            marco1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            marco1.BackColor = Color.AliceBlue;
            //
			lb_serv = new Label { Text = "Servidor: ", Location = new Point(this.Left + 10 , this.Top + 5) };
            tx_serv = new TextBox { Location = new Point(lb_serv.Width + 10, this.Top + 3), Width = 150, ReadOnly = true };
            lb_base = new Label { Text = "Base datos: ", Location = new Point(300, this.Top + 5) };
            tx_base = new TextBox { Location = new Point(lb_base.Left + lb_base.Width + 10, this.Top + 5), Width = 100, ReadOnly = true };
            lb_tabl = new Label { Text = "Tabla Vtas.: ", Location = new Point(430, this.Top + 5), Anchor = AnchorStyles.None };
            tx_tabl = new TextBox { Location = new Point(lb_tabl.Left + lb_tabl.Width + 10, this.Top + 5), Width = 100, ReadOnly = true, Anchor = AnchorStyles.None };
            lb_tabd = new Label { Text = "Detalle Vtas.", Location = new Point(700, this.Top + 5), Anchor = AnchorStyles.None };
            tx_tabd = new TextBox { Location = new Point(lb_tabd.Left + lb_tabd.Width + 10, this.Top + 5), Width = 100, ReadOnly = true, Anchor = AnchorStyles.None };
            lb_usua = new Label { Text = "Usuario: ", Location = new Point(this.Left + 10, this.Top + 35) };
            tx_usua = new TextBox { Location = new Point(lb_usua.Width + 10, this.Top + 33), Width = 100, ReadOnly = true };
            lb_pass = new Label { Text = "Contraseña: ", Location = new Point(300, this.Top + 35) };
            tx_pass = new TextBox { Location = new Point(lb_pass.Left + lb_pass.Width + 10, this.Top + 35), ReadOnly = true, PasswordChar = '*' };
            lb_port = new Label { Text = "Puerto: ", Location = new Point(430, this.Top + 35), Anchor = (AnchorStyles.None) };
            tx_port = new TextBox { Location = new Point(lb_port.Left + lb_port.Width + 10, this.Top + 35), ReadOnly = true, Anchor = (AnchorStyles.None) };
            //
            marco2 = new Panel();
            marco2.Location = new Point(this.Left, marco1.Height+5);
            marco2.Size = new Size(this.Width, 30);
            marco2.Anchor = (AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top);
            marco2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            marco2.BackColor = Color.AliceBlue;
            lb_letr = new Label { Text = "Letras: ", Location = new Point(this.Left + 10, this.Top + 5) };
            tx_letr = new TextBox { Location = new Point(lb_letr.Width + 10, this.Top + 3), Width = 100, ReadOnly = true };
            lb_digs = new Label { Text = "#Digs.serie: ", Location = new Point(300, this.Top + 5) };
            tx_digs = new TextBox { Location = new Point(lb_digs.Left + lb_digs.Width + 10, this.Top + 5), Width = 100, ReadOnly = true };
            lb_time = new Label { Text = "Tiempo (seg): ", Location = new Point(430, this.Top + 5), Anchor = AnchorStyles.None };
            tx_time = new TextBox { Location = new Point(lb_time.Left + lb_time.Width + 10, this.Top + 5), Width = 100, ReadOnly = true, Anchor = AnchorStyles.None };
            //
            marco3 = new Panel();
            marco3.Location = new Point(this.Left, marco2.Top + marco2.Height + 5);
            marco3.Size = new Size(this.Width, 30);
            marco3.Anchor = (AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top);
            marco3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            marco3.BackColor = Color.AliceBlue;
            lb_ano = new Label { Text = "Año: ", Location = new Point(this.Left + 10, this.Top + 5),Width = 30 };
            tx_ano = new TextBox { Location = new Point(lb_ano.Width + 10, this.Top + 3), Width = 50, ReadOnly = false, TextAlign = HorizontalAlignment.Center };
            lb_mes = new Label { Text = "Mes: ", Location = new Point(tx_ano.Left + tx_ano.Width + 10, this.Top + 5),Width = 30 };
            tx_mes = new TextBox { Location = new Point(lb_mes.Left + lb_mes.Width + 10, this.Top + 3), Width = 50, ReadOnly = false, TextAlign = HorizontalAlignment.Center };
            lb_dia = new Label { Text = "Día: ", Location = new Point(tx_mes.Left + tx_mes.Width + 10, this.Top + 5),Width = 30 };
            tx_dia = new TextBox { Location = new Point(lb_dia.Left + lb_dia.Width + 10, this.Top + 3), Width = 50, ReadOnly = false, TextAlign = HorizontalAlignment.Center };
            lb_spe = new Label { Text = "Series Electrónicas: ", Location = new Point(tx_dia.Left + tx_dia.Width + 20, this.Top + 5), Width = 100 };
            tx_spe = new TextBox { Location = new Point(lb_spe.Left + lb_spe.Width + 10, this.Top + 3), Width = 150, ReadOnly = true, TextAlign = HorizontalAlignment.Center };
            lb_usu = new Label { Text = "Usuario: ", Location = new Point(tx_spe.Left + tx_spe.Width + 20, this.Top + 5), Width = 100 };
            tx_usu = new TextBox { Location = new Point(lb_usu.Left + lb_usu.Width + 10, this.Top + 3), Width = 150, ReadOnly = true, TextAlign = HorizontalAlignment.Center };
            lb_loc = new Label { Text = "Tienda: ", Location = new Point(tx_usu.Left + tx_usu.Width + 20, this.Top + 5), Width = 100 };
            tx_loc = new TextBox { Location = new Point(lb_loc.Left + lb_loc.Width + 10, this.Top + 3), Width = 150, ReadOnly = true, TextAlign = HorizontalAlignment.Center };
            //
            btn_ = new Button { Text = "Obtiene datos", Location = new Point(this.Width - 110, 2), BackColor = Color.BlueViolet, Anchor = (AnchorStyles.Right | AnchorStyles.Top), AutoSize = true };
			btn_.Click += Btn_Click; // Handle the event.
			// Attach these objects to the graphics window.
            this.Controls.Add(marco1);
            marco1.Controls.Add(lb_serv);
            marco1.Controls.Add(tx_serv);
            marco1.Controls.Add(lb_base);
            marco1.Controls.Add(tx_base);
            marco1.Controls.Add(lb_tabl);
            marco1.Controls.Add(tx_tabl);
            marco1.Controls.Add(lb_tabd);
            marco1.Controls.Add(tx_tabd);
            marco1.Controls.Add(lb_usua);
            marco1.Controls.Add(tx_usua);
            marco1.Controls.Add(lb_pass);
            marco1.Controls.Add(tx_pass);
            marco1.Controls.Add(lb_port);
            marco1.Controls.Add(tx_port);
            //
            this.Controls.Add(marco2);
            marco2.Controls.Add(lb_letr);
            marco2.Controls.Add(tx_letr);
            marco2.Controls.Add(lb_digs);
            marco2.Controls.Add(tx_digs);
            marco2.Controls.Add(lb_time);
            marco2.Controls.Add(tx_time);
            //
            this.Controls.Add(marco3);
            marco3.Controls.Add(lb_ano);
            marco3.Controls.Add(tx_ano);
            marco3.Controls.Add(lb_mes);
            marco3.Controls.Add(tx_mes);
            marco3.Controls.Add(lb_dia);
            marco3.Controls.Add(tx_dia);
            marco3.Controls.Add(lb_spe);
            marco3.Controls.Add(tx_spe);
            marco3.Controls.Add(btn_);
            //
            marco4 = new Panel();
            marco4.Location = new Point(this.Left, marco3.Top + marco3.Height + 5);
            marco4.Size = new Size(ancho, largo-10);
            marco4.Anchor = (AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top);    // AnchorStyles.Bottom | 
            marco4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            marco4.BackColor = Color.AliceBlue;
            this.Controls.Add(marco4);
            grilla = new DataGridView();
            grilla.Location = new Point(this.Left, this.Top);
            grilla.Size = new Size(marco4.Width, marco4.Height);
            grilla.Anchor = (AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top);
            grilla.ReadOnly = true;
            marco4.Controls.Add(grilla);
            //
            marco5 = new Panel();
            marco5.Location = new Point(this.Left, marco4.Top + marco4.Height + 5);
            marco5.Size = new Size(ancho, 30);
            marco5.Anchor = (AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top);    //  |  
            marco5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            marco5.BackColor = Color.AliceBlue;
            this.Controls.Add(marco5);
            bt_print = new Button { Text = "Imprime DV", Location = new Point(ancho - 200, 6), BackColor = Color.BlueViolet, Anchor = (AnchorStyles.Right | AnchorStyles.Bottom), AutoSize = true };
            bt_print.Click += Bt_print_Click;
            bt_anu = new Button { Text = "Anula DV", Location = new Point(ancho - 110, 6), BackColor = Color.BlueViolet, Anchor = (AnchorStyles.Right | AnchorStyles.Bottom), AutoSize = true };
            bt_anu.Click += Bt_anu_Click;
            marco5.Controls.Add(bt_print);
            marco5.Controls.Add(bt_anu);
		}
        private void jalainfo()         // obtiene informacion inicial del sistema, ruc, nombre cliente, etc.
        {
            MySqlConnection conn = new MySqlConnection(DB_CONN_STR);
            try
            {
                conn.Open();
                string consulta = "select * from baseconf limit 1";
                MySqlCommand micon = new MySqlCommand(consulta, conn);
                MySqlDataReader dr = micon.ExecuteReader();
                if (dr.Read())
                {
                    nomclie = dr.GetString("Cliente");
                    rucclie = dr.GetString("Ruc");
                }
                dr.Close();
            }
            catch (MySqlException ex)
            {
                MessageBox.Show(ex.Message, "Error de conexión");
                Application.Exit();
                return;
            }
            conn.Close();
        }
        private void jala_marco1()      // jala los datos de conexion a la base del cliente
        {
            MySqlConnection conn = new MySqlConnection(DB_CONN_STR);
            try
            {
                conn.Open();
                string consulta = "select * from clientconf where id=1";
                MySqlCommand micon = new MySqlCommand(consulta,conn);
                MySqlDataReader dr = micon.ExecuteReader();
                if(dr.HasRows)
                {
                    if(dr.Read()){
                        tx_serv.Text = dr.GetString(1);
                        tx_base.Text = dr.GetString(2);
                        tx_tabl.Text = dr.GetString(6);
                        tx_port.Text = dr.GetString(3);
                        tx_usua.Text = dr.GetString(4);
                        tx_pass.Text = dr.GetString(5);
                        tx_tabd.Text = dr.GetString(7);
                        // tabla de totales o pie de los docs
                        dr.Close();
                    }
                }else{
                    dr.Close();
                    MessageBox.Show("No esta configurado el acceso al cliente","Error de configuración");
                }
            }
            catch(MySqlException ex)
            {
                MessageBox.Show(ex.Message,"Error de conexión");
                Application.Exit();
                return;
            }
            conn.Close();
        }
        private void jala_marco2()      // jala parametros para obtener los datos del cliente
        {
            MySqlConnection conn = new MySqlConnection(DB_CONN_STR);
            conn.Open();
            try
            {
                string consulta = "select campo,param,valor from enlaces where formulario=@form and param='Default'";
                MySqlCommand micon = new MySqlCommand(consulta, conn);
                micon.Parameters.AddWithValue("@form", nomform);
                MySqlDataAdapter da = new MySqlDataAdapter(micon);
                DataTable dt = new DataTable();
                da.Fill(dt);
                foreach(DataRow row in dt.Rows){
                    if (row[0].ToString() == "letras") tx_letr.Text = row[2].ToString();
                    if (row[0].ToString() == "digitos") tx_digs.Text = row[2].ToString();
                    if (row[0].ToString() == "tiempo") tx_time.Text = row[2].ToString();
                }
            }
            catch (MySqlException ex)
            {
                MessageBox.Show(ex.Message, "Error de conexión");
                Application.Exit();
                return;
            }
            conn.Close();
            if (tloa != "") tx_time.Text = tloa;    // si esta configurado el cliente, vale esa, si no esta config. vale la del servidor
        }
        private void jala_marco3()      // jalamos los parametros de filtro de entrada
        {
            tx_spe.Text = sped;         // jalamos la serie del punto de emision
            tx_ano.Text = DateTime.Now.Date.Year.ToString();
            tx_mes.Text = DateTime.Now.Date.Month.ToString();
            tx_dia.Text = DateTime.Now.Date.Day.ToString();
            tx_usu.Text = "";   // jalamos de donde ??? el usuario
            tx_loc.Text = "";   // jalamos el local del usuario
        }

        private void jala_datos()       // obtiene los datos del cliente
        {
            // conexion a la base del cliente
            string CONN_CLTE = "server=" + tx_serv.Text.Trim() + ";port=" + tx_port.Text.Trim() + ";uid=" + tx_usua.Text.Trim() + 
                ";pwd=" + tx_pass.Text.Trim() + ";database=" + tx_base.Text.Trim() + ";ConnectionLifeTime=" + ctl + ";";
            // jalamos los datos
            MySqlConnection conc = new MySqlConnection(CONN_CLTE);
            try
            {
                string parte = "";
                if (tx_dia.Text == "" && tx_mes.Text == "") parte = "";
                if (tx_dia.Text == "" && tx_mes.Text != "") parte = " month(a.fechope)=@mes and";
                if (tx_dia.Text != "" && tx_mes.Text != "") parte = " month(a.fechope)=@mes and day(a.fechope)=@dia and";
                if (tx_spe.Text != "") parte = parte + " a.servta in (" + tx_spe.Text.Trim() + ")";
                conc.Open();
                // obtenemos los datos del cliente, cabecera de ventas
                string consulta = "select a.iddocvtas,a.usercaja,a.fechope,a.tipcam,a.local,a.docvta,a.servta,a.corrvta,a.doccli,a.numdcli," +
                    "a.nomcli,a.clidv,a.telef,a.direc1,a.direc2,a.dist,a.pedido,a.observ,a.observ2,a.moneda,a.subtot,a.igv,a.doctot,a.status," +
                    "a.codcli,a.dia,a.useranul,a.fechanul,a.marca1 " +
                    "from " + tx_tabl.Text.Trim() + " a " +
                    "where year(a.fechope)=@yea and " + parte;
                MySqlCommand micon = new MySqlCommand(consulta, conc);
                micon.Parameters.AddWithValue("@yea", tx_ano.Text);
                micon.Parameters.AddWithValue("@mes", tx_mes.Text);
                micon.Parameters.AddWithValue("@dia", tx_dia.Text);
                MySqlDataAdapter da = new MySqlDataAdapter(micon);
                DataTable dt = new DataTable();
                da.Fill(dt);
                // obtenemos los datos del cliente, detalle de ventas
                consulta = "select b.tipdv,b.servta,b.corvta,b.codprd,b.descrip,b.precio,b.cantid,b.total,b.precdol,b.totdol,b.local,b.pedido,b.status,b.marca1 " +
                    "from " + tx_tabd.Text.Trim() + " b left join " + tx_tabl.Text.Trim() +
                    "a on a.docvta=b.tipdv and a.servta=b.servta and a.corrvta=b.corvta " +
                    "where year(a.fechope)=@yea and " + parte;
                micon = new MySqlCommand(consulta, conc);
                micon.Parameters.AddWithValue("@yea", tx_ano.Text);
                micon.Parameters.AddWithValue("@mes", tx_mes.Text);
                micon.Parameters.AddWithValue("@dia", tx_dia.Text);
                MySqlDataAdapter dad = new MySqlDataAdapter(micon);
                DataTable dtd = new DataTable();
                dad.Fill(dtd);
                // insertamos la cabecera
                foreach (DataRow row in dt.Rows)
                {
                    string vid = row[0].ToString();     // ID del registro
                    string usca = row[1].ToString();     // usuario de la caja
                    string vfec = string.Format("{0:yyyy-MM-dd}",row[2]); // fecha del documento
                    string vtca = row[3].ToString(); // tipo de cambio
                    string vloc = row[4].ToString();   // local donde se emitio el doc
                    string vdvt = row[5].ToString(); // cod.doc.vta
                    string vsvt = row[6].ToString(); // serie
                    string vcvt = row[7].ToString(); // correlativo
                    string vdcl = row[8].ToString();    // tip doc cliente
                    string vndc = row[9].ToString();    // num doc cliente
                    string vrsc = row[10].ToString();    // nombre cliente 
                    string vcdv = row[11].ToString();    // nombre cliente del doc.vta
                    string vtcl = row[12].ToString();    // telefono del cliente
                    string vdic = row[13].ToString();    // direcc. cliente
                    string vdc2 = row[14].ToString();    // direcc. cliente
                    string vdsc = row[15].ToString();   // dist. cliente
                    string vpcl = row[16].ToString();   // pedido/contrato
                    string vobs = row[17].ToString();   // comentarios del doc.vta.
                    string vob2 = row[18].ToString();   // comentarios 2 del doc.vta.
                    string vmon = row[19].ToString();   // moneda del documento
                    string vsub = row[20].ToString();   // sub total del doc
                    string vigv = row[21].ToString();   // igv del doc
                    string vtot = row[22].ToString();   // total del doc
                    string vsta = row[23].ToString();   // situacion de doc
                    string vccl = row[24].ToString();   // codigo del cliente
                    string vdtc = row[25].ToString();   // fecha de creacion del doc
                    string vuan = row[26].ToString();   // usuario que anula
                    string vfan = row[27].ToString();   // fecha de la anulacion
                    string vma1 = row[28].ToString();   // marca1
                    // insertamos la tabla propia
                    string inserta = "insert into madocvtas (usercaja,fechope,tipcam,local,docvta,servta,corrvta,doccli,numdcli,direc1,direc2,nomcli," +
                        "telef,codcli,clidv,pedido,observ,observ2,moneda,subtot,igv,doctot,status,dia,useranul,fechanul,marca1,dist,cdr) values (" +
                        "@usca,@vfec,@vtca,@vloc,@vdvt,@vsvt,@vcvt,@vdcl,@vndc,@vdic,@vdc2,@vrsc," +
                        "@vtcl,@vccl,@vcdv,@vpcl,@vobs,@vob2,@vmon,@vsub,@vigv,@vtot,@vsta,@vdtc,@vuan,@vfan,@vma1,@vdsc,@cdr)";
                    MySqlConnection conn = new MySqlConnection(DB_CONN_STR);
                    try
                    {
                        conn.Open();
                        MySqlCommand mins = new MySqlCommand(inserta, conn);
                        mins.Parameters.AddWithValue("@usca", usca);    // usuario
                        mins.Parameters.AddWithValue("@vfec", vfec);    // fecha
                        mins.Parameters.AddWithValue("@vtca", vtca);    // tipo de cambio
                        mins.Parameters.AddWithValue("@vloc", vloc);    // local de emision
                        mins.Parameters.AddWithValue("@vdvt", vdvt);    // tipo doc venta
                        mins.Parameters.AddWithValue("@vsvt", vsvt);    // serie
                        mins.Parameters.AddWithValue("@vcvt", vcvt);    // correlativo
                        mins.Parameters.AddWithValue("@vdcl", vdcl);    // tipo doc cliente
                        mins.Parameters.AddWithValue("@vndc", vndc);    // numero doc cliente
                        mins.Parameters.AddWithValue("@vdic", vdic);    // direccion
                        mins.Parameters.AddWithValue("@vdi2", vdc2);    // direccion 2
                        mins.Parameters.AddWithValue("@vrsc", vrsc);    // nombre del cliente
                        mins.Parameters.AddWithValue("@vtcl", vtcl);    // telefono
                        mins.Parameters.AddWithValue("@vcdv", vcdv);    // codigo cliente
                        mins.Parameters.AddWithValue("@vpcl", vpcl);    // pedido/contrato con el cliente
                        mins.Parameters.AddWithValue("@vobs", vobs);    // observaciones
                        mins.Parameters.AddWithValue("@vob2", vob2);    // observaciones 2
                        mins.Parameters.AddWithValue("@vmon", vmon);    // moneda
                        mins.Parameters.AddWithValue("@vsub", vsub);    // sub total
                        mins.Parameters.AddWithValue("@vigv", vigv);    // igv
                        mins.Parameters.AddWithValue("@vtot", vtot);    // total    
                        mins.Parameters.AddWithValue("@vsta", vsta);    // estado del doc
                        mins.Parameters.AddWithValue("@vdtc", vdtc);    // fecha creacion del doc
                        mins.Parameters.AddWithValue("@vuan", vuan);    // usuario que anulo doc
                        mins.Parameters.AddWithValue("@vfan", vfan);    // fecha de anulacion
                        mins.Parameters.AddWithValue("@vma1", vma1);    // marca 1
                        mins.Parameters.AddWithValue("@vdsc", vdsc);    // distrito
                        mins.Parameters.AddWithValue("@cdr", "9");    // cdr
                        mins.ExecuteNonQuery();
                    }
                    catch (MySqlException ex)
                    {
                        MessageBox.Show(ex.Message, "Error de conexión al servidor propio");
                        Application.Exit();
                        return;
                    }
                    conn.Close();
                    // marcamos los registros obtenidos en la tabla del cliente 
                    string actua = "update " + tx_tabl.Text.Trim() + " set cdr=@cdr where id=@vid";
                    MySqlCommand miact = new MySqlCommand(actua, conc);
                    miact.Parameters.AddWithValue("@cdr", "9");
                    miact.Parameters.AddWithValue("@vid", vid);
                    miact.ExecuteNonQuery();
                }
                // insertamos el detalle 
                foreach (DataRow row in dtd.Rows)
                {
                    //b.tipdv,b.servta,b.corvta,b.codprd,b.descrip,b.precio,b.cantid,b.total,b.precdol,b.totdol,b.local,b.pedido,b.status,b.marca1
                    string vdvt = row[0].ToString(); // cod.doc.vta
                    string vsvt = row[1].ToString(); // serie
                    string vcvt = row[2].ToString(); // correlativo
                    string vcpd = row[3].ToString();    // codigo del producto
                    string vpre = row[4].ToString();    // precio 
                    string vcan = row[5].ToString();    // cantidad 
                    string vtot = row[6].ToString();   // total del doc
                    string vtod = row[7].ToString();   // total del doc dolares
                    string vloc = row[8].ToString();   // local
                    string vped = row[9].ToString();   // pedido/contrato
                    string vsta = row[10].ToString();   // situacion de doc
                    string vmar = row[11].ToString();   // marca
                    // insertamos la tabla propia
                    string inserta = "insert into detavtas (tipdv,servta,corvta,codprd,descrip,precio,cantid,total,precdol,totdol,local,pedido,status,marca1) values (" +
                        "@vdvt,@vsvt,@vcvt,@vcpd,@vpre,@vcan,@vtot,@vtod,@vloc,@vped,@vsta,@vmar)";
                    MySqlConnection conn = new MySqlConnection(DB_CONN_STR);
                    try
                    {
                        conn.Open();
                        MySqlCommand mins = new MySqlCommand(inserta, conn);
                        mins.Parameters.AddWithValue("@vdvt", vdvt);    // cod.doc.vta
                        mins.Parameters.AddWithValue("@vsvt", vsvt);    // serie
                        mins.Parameters.AddWithValue("@vcvt", vcvt);    // correlativo
                        mins.Parameters.AddWithValue("@vcpd", vcpd);    // codigo del producto
                        mins.Parameters.AddWithValue("@vpre", vpre);    // precio 
                        mins.Parameters.AddWithValue("@vcan", vcan);    // cantidad
                        mins.Parameters.AddWithValue("@vtot", vtot);    // total del doc
                        mins.Parameters.AddWithValue("@vtod", vtod);    // total del doc dolares
                        mins.Parameters.AddWithValue("@vloc", vloc);    // local
                        mins.Parameters.AddWithValue("@vped", vped);    // pedido/contrato
                        mins.Parameters.AddWithValue("@vsta", vsta);    // situacion de doc
                        mins.Parameters.AddWithValue("@vmar", vmar);    // marca
                        mins.ExecuteNonQuery();
                    }
                    catch (MySqlException ex)
                    {
                        MessageBox.Show(ex.Message, "Error de conexión al servidor propio");
                        Application.Exit();
                        return;
                    }
                    conn.Close();
                }
            }
            catch (MySqlException ex)
            {
                MessageBox.Show(ex.Message, "Error en conexión con el cliente");
                Application.Exit();
                return;
            }
            // cerramos la conexión con el cliente
            conc.Close();
        }
        private void muestra_datos()    // muestra los datos de la tabla propia en la grilla usando los filtros de entrada
        {
            string condicion="";
            if (tx_mes.Text != "00") condicion = " and month(fechope)=@mes";
			if (tx_dia.Text != "00") condicion = " and day(fechope)=@dia";
            if (tx_spe.Text != "") condicion = condicion + " and servta in (" + tx_spe.Text.Trim() + ")";   // falta usuario y fecha
            condicion = condicion + " and usercaja=@ucaj and local=@loca";
            MySqlConnection conn = new MySqlConnection(DB_CONN_STR);
            try
            {
                conn.Open();
                string consulta = "select fechope,tipcam,docvta,servta,corrvta,doccli,numdcli,direc1,direc2,nomcli," +
                        "telef,codcli,clidv,pedido,observ,observ2,moneda,subtot,igv,doctot,status,dia,useranul,fechanul,marca1,dist,cdr,impreso " +
                        "from docvtas where year(fechope)=@yea" + condicion;
				MySqlCommand micon = new MySqlCommand(consulta,conn);
				micon.Parameters.AddWithValue("@yea", tx_ano.Text);
				micon.Parameters.AddWithValue("@mes", tx_mes.Text);
				micon.Parameters.AddWithValue("@dia", tx_dia.Text);
                micon.Parameters.AddWithValue("@ucaj", tx_usu.Text);
                micon.Parameters.AddWithValue("@loca", tx_loc.Text);
				MySqlDataAdapter da = new MySqlDataAdapter(micon);
				DataTable dt = new DataTable();
				da.Fill(dt);
                if(dt.Rows.Count>0) grilla.DataSource = dt;
            }
            catch (MySqlException ex)
            {
                MessageBox.Show(ex.Message, "Error en obtener docs.vta.elect");
                Application.Exit();
                return;
            }
			conn.Close ();
        }

        void miform_FormClosing(object sender, EventArgs e)
        {
            if (temporizador != null)
            {
                temporizador.Stop();
                temporizador.Dispose();
            }
        }
		void Btn_Click(object sender, EventArgs e)  // boton jala datos
		{
            if (tx_ano.Text == "")
            {
                MessageBox.Show("Ingrese el año de los documentos", "Atención verifique");
                tx_ano.Focus();
                return;
            }
            if (tx_mes.Text == "")
            {
                MessageBox.Show("Ingrese el mes de los documentos", "Atención verifique");
                tx_mes.Focus();
                return;
            }
            if (tx_dia.Text == "")
            {
                MessageBox.Show("Ingrese el dia de los documentos", "Atención verifique");
                tx_dia.Focus();
                return;
            }
            trabaja();
            // llamamos al temporizador
            if (temporizador != null)
            {
                temporizador.Stop();
                temporizador.Dispose();
            }
            InitTimer();
		}
        void Bt_print_Click(object sender, EventArgs e) // imprime los registros seleccionados
        {
            if (grilla.SelectedRows.Count > 0)
            {
                foreach (DataGridViewRow dgv in grilla.SelectedRows)
                {

                }
            }
            else
            {
                MessageBox.Show("Ud. debe dar click al extremo izquierdo de la" + Environment.NewLine +
                "fila que desea seleccionar, de forma similar" + Environment.NewLine +
                "a una hoja de calculo.", "Atención - Seleccione una fila", MessageBoxButtons.OK, MessageBoxIcon.Information);
                grilla.Focus();
                return;
            }
        }
        void Bt_anu_Click(object sender, EventArgs e)   // anula (da de baja) los registros seleccionados
        {
            if (grilla.SelectedRows.Count > 0)
            {
                foreach (DataGridViewRow dgv in grilla.SelectedRows)
                {

                }
            }
            else
            {
                MessageBox.Show("Ud. debe dar click al extremo izquierdo de la" + Environment.NewLine +
                "fila que desea seleccionar, de forma similar" + Environment.NewLine +
                "a una hoja de calculo.", "Atención - Seleccione una fila", MessageBoxButtons.OK, MessageBoxIcon.Information);
                grilla.Focus();
                return;
            }
        }
        private void trabaja()
        {
            // mensaje de aviso de la carga
            Form f = new Form();
            f.Size = new System.Drawing.Size(400, 50);
            f.Location = new Point(400, 200);
            f.Text = "Obteniendo datos de ventas ....";
            f.MinimizeBox = false;
            f.MaximizeBox = false;
            f.ControlBox = false;
            f.Enabled = false;
            f.StartPosition = FormStartPosition.CenterScreen;
            f.Show();
            // jalamos y mostramos los datos
            jala_datos();
            muestra_datos();
            //cerramos el mensaje de aviso de la carga
            f.Enabled = true;
            f.Close();
            // llamamos al creador de xml y tenemos registros nuevos
            creador();
        }
        private void InitTimer()
        {
            temporizador = new System.Windows.Forms.Timer();
            temporizador.Interval = int.Parse(tx_time.Text) * 1000;    // 10000;        // 10 segundos
            temporizador.Tick += new EventHandler(timer_tick);
            temporizador.Start();
        }
        private void timer_tick(object sender, EventArgs e)
        {
            //Btn_Click(null, null);
            trabaja();
        }
        private void creador()          // creador de xml y zip por cada registro
        {
            // nombres de archivos
            string ruta = Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
            ruta = ruta + "/xml/";
            for (int i = 0; i < grilla.Rows.Count - 1; i++)
            {
                string archi = rucclie + "-";
                if (grilla.Rows[i].Cells["cdr"].Value.ToString() == "9")
                {
                    archi = archi + grilla.Rows[i].Cells["docvta"].Value.ToString() + "-" +
                        grilla.Rows[i].Cells["servta"].Value.ToString() + "-" +
                        grilla.Rows[i].Cells["corvta"].Value.ToString().Trim();
                    crearXML(grilla.Rows[i].Cells["id"].Value.ToString(), ruta + archi);
                    crearZIP(ruta, archi);
                    actualizaCDR(i,grilla.Rows[i].Cells["id"].Value.ToString(), "8"); // actualizamos CDR del registro
                }
            }
        }
        public void crearXML(string idr,String file_path)
        {
            string CONN_CLTE = "server=" + tx_serv.Text.Trim() + ";port=" + tx_port.Text.Trim() + ";uid=" + tx_usua.Text.Trim() + 
                ";pwd=" + tx_pass.Text.Trim() + ";database=" + tx_base.Text.Trim() + ";ConnectionLifeTime=" + ctl + ";";
            // jalamos los datos
            MySqlConnection conc = new MySqlConnection(CONN_CLTE);
            try
            {
                conc.Open();
                string concab = "select * from " + tx_tabl.Text.Trim() + " where id=@idr";
                string condet = "select * from " + tx_tabd.Text.Trim() + " where idc=@idr";
                MySqlCommand micon1 = new MySqlCommand(concab, conc);
                micon1.Parameters.AddWithValue("@idr", idr);
                MySqlDataReader dr = micon1.ExecuteReader();
                if (dr.Read())      // ******************* CABECERA ***********************
                {
                    string _fecemi = "";    // 01 fecha de emision   yyyy-mm-dd                 10
                    // 02 firma digital             3000
                    string Prazsoc = "";    // 03 razon social del emisor                      100
                    string Pnomcom = "";    // 04 nombre comercial del emisor                  100
                    string Pdf_ubi = "";    // 05 DOMICILIO FISCAL - Ubigeo                      6
                    string Pdf_dir = "";    // 05 DOMICILIO FISCAL - direccion                 100
                    string Pdf_urb = "";    // 05 DOMICILIO FISCAL - Urbanizacion               25
                    string Pdf_pro = "";    // 05 DOMICILIO FISCAL - provincia                  30
                    string Pdf_dep = "";    // 05 DOMICILIO FISCAL - departamento               30
                    string Pdf_dis = "";    // 05 DOMICILIO FISCAL - distrito                   30
                    string Pdf_cpa = "";    // 05 DOMICILIO FISCAL - código de país              2
                    string Prucpro = "";    // 06 Ruc del emisor                                11
                    string Pcrupro = "";    // 06 codigo Ruc emisor                              1
                    string _tipdoc = "";    // 07 Tipo de documento de venta                     1
                    string _sercor = "";    // 08 Serie y correlat concatenado F001-00000001    13
                    string Cnumdoc = "";    // 09 numero de doc. del cliente                    15
                    string Ctipdoc = "";    // 09 tipo de doc. del cliente                       1
                    string Cnomcli = "";    // 10 nombre del cliente                           100
                    string Iumeded = "";    // 11 DETALLE - Unidad de medida                     3
                    string Icantid = "";    // 12 DETALLE - Cantidad de items   n(12,3)         16
                    string Idescri = "";    // 13 DETALLE - Nombre o descripcion               250
                    string Ivaluni = "";    // 14 DETALLE - Valor unitario del item n(12,2)     15
                    string Ipreuni = "";    // 15 DETALLE - Precio de venta unitario n(12,2)    15
                    string Icodpre = "";    // 15 DETALLE - Código tipo de precio                2
                    string Iigvite = "";    // 16 DETALLE - monto IGV del item  n(12,2)         15
                    string Isubigv = "";    // 16 DETALLE - sub tot igv n(12,2)                 15
                    string Icatigv = "";    // 16 DETALLE - Afectacion al igv                    2
                    string Icodigv = "";    // 16 DETALLE - Código de tributo                    4
                    string Inomtig = "";    // 16 DETALLE - NOmbre del tributo                   6
                    string Iinttri = "";    // 16 DETALLE - Codigo internacional del tributo     3
                    string Iiscite = "";    // 17 DETALLE - ISC del item n(12,2)                15
                    string Imonisc = "";    // 17 DETALLE - Monto del ISC del item n(12,2)      15
                    string Itipisc = "";    // 17 DETALLE - Tipo de sistema isc                  2
                    string Icodtri = "";    // 17 DETALLE - Codigo del tributo                   4
                    string Inomtis = "";    // 17 DETALLE - Nombre del tributo isc               6
                    string Icoditi = "";    // 17 DETALLE - Código internacional del tributo     3
                    string _codtmo = "";    // 18 Código tipo de monto                           4
                    string _totogr = "";    // 18 Tot valor venta operaciones grabadas n(12,2)  15
                    string _codtvo = "";    // 19 Código total valor operac. inafectas           4
                    string _totvoi = "";    // 19 tot valor venta operaciones inafectas n(12,2) 15
                    string _codvoe = "";    // 20 codigo operaciones exoneradas                  4
                    string _totvoe = "";    // 20 tot valor venta operaciones exoneradas        15
                    string Ivalvta = "";    // 21 valor venta por item                          15
                    string _sumigv = "";    // 22 Sumatoria de igv                              15
                    string _sumig2 = "";    // 22 Sumatoria de igv                              15
                    string _cotrig = "";    // 22 Código de tributo                              4
                    string _notrig = "";    // 22 Nombre del tributo                             6
                    string _coitri = "";    // 22 Código internacional del tributo               3
                    string _sumisc = "";    // 23 Sumatoria de ISC                              15
                    string _sumis2 = "";    // 23 Sumatoria de ISC                              15
                    string _cotisc = "";    // 23 Código del tributo isc                         4
                    string _notisc = "";    // 23 NOmbre del tributo isc                         6
                    string _coiisc = "";    // 23 Código internacional de tributo isc            3
                    string _sumotr = "";    // 24 Sumatoria de otros tributos                   15
                    string _sumot2 = "";    // 24 Sumatoria de otros tributos                   15
                    string _cootri = "";    // 24 Codigo otros tributos                          4
                    string _nootri = "";    // 24 Nombre otros tributos                          6
                    string _coiotr = "";    // 24 codigo internacional del otro tributo          3
                    string _suotca = "";    // 25 Sumatoria de otros cargos                     15
                    string _cotdes = "";    // 26 Código tipo de descuento                       4
                    string _totdes = "";    // 26 Total descuentos                              15
                    string _totven = "";    // 27 Importe total de la venta n(12,2)             15
                    string _moneda = "";    // 28 Moneda del doc. de venta                       3
                    string _scguia = "";    // 29 serie y numero concatenado de la guia         30
                    string _codgui = "";    // 29 Código de la guia de remision                  2
                    string _scotro = "";    // 30 serie y numero concatenado de otro docu       30
                    string _codotr = "";    // 30 Código del otro documento                      2
                    string _codley = "";    // 31 Codigo de la leyenda                           4
                    string _leyend = "";    // 31 descripcion de la leyenda                    100
                    string _codper = "";    // 32 Código de la percpcion                         4
                    string _bpermn = "";    // 32 base imponible percepcion en moneda nac       15
                    string _mpermn = "";    // 32 monto de la percepcion                        15
                    string _mtpemn = "";    // 32 monto total incluido la percepcion            15
                    string Inumord = "";    // 33 numero de orden del item                       3
                    string Icodpro = "";    // 34 codigo del producto                           30
                    string Ivareun = "";    // 35 Valor referencial unitario del item           15
                    string Icvarun = "";    // 35 Código del valor referencial unitario          2
                    string _verubl = "";    // 36 Version del UBL                               10
                    string _verest = "";    // 37 Versión de la estructura del documento        10
                    string _cvrstv = "";    // 38 codigo del valor refe. del serv. de trans      4
                    string _vrstvt = "";    // 38 valor referencial del serv. de transp. terres 15
                    string _codepe = "";    // 39 codigo del asunto de la embarcacion pesq       4
                    string _nymepe = "";    // 39 nomgre y registro de la embarcacion pesquer  100
                    string _ctydcv = "";    // 40 codigo del tipo y especia vendida              4
                    string _dtceve = "";    // 40 descrip.del tipo y cant. de especie vendida  150
                    string _cldesc = "";    // 41 codigo lugar de la descarga                    4
                    string _lugdes = "";    // 41 lugar de la descarga del pescado             100
                    string _ffedes = "";    // 42 codigo fecha de la descarga                    4
                    string _fecdes = "";    // 42 fecha de la descarga     yyyy-mm-dd           10
                    string _cremtc = "";    // 43 codigo registro del mtc                        4
                    string _regmtc = "";    // 43 registro del mtc                              20
                    string _ccoveh = "";    // 44 codigo configuracion vehicular                 4
                    string _conveh = "";    // 44 configuracion vehicular                       20
                    string _cpuori = "";    // 45 codigo punto de origen                         4
                    string _punori = "";    // 45 punto de origen                              100
                    string _cpudes = "";    // 46 codigo punto destino                           4
                    string _pundes = "";    // 46 punto de destino                             100
                    string _cvalpr = "";    // 47 codigo valor preliminar                        4
                    string _destra = "";    // 47 descripcion del tramo o viaje                100
                    string _valpre = "";    // 47 valor preliminar                              15
                    string _cfecon = "";    // 48 codigo fecha de consumo                        4
                    string _feccon = "";    // 48 fecha de consumo                              10
                    string _ctvogr = "";    // 49 codigo total valor venta op.gratuitas          4
                    string _totvog = "";    // 49 total valor venta operac gratuitas (15,2)     18
                    string _desglo = "";    // 50 descuentos globales                           15
                    string _cdeite = "";    // 51 codigo descuentos por item                     5
                    string _desite = "";    // 51 descuentos por item                           15
                }
                dr.Close();

                XmlTextWriter writer;
                file_path = file_path + ".XML";
                writer = new XmlTextWriter(file_path, Encoding.UTF8);
                writer.Formatting = Formatting.Indented;
                writer.WriteStartDocument();
                writer.WriteStartElement("Invoice");
                writer.WriteElementString("nodo1", "texto del nodo1");
                writer.WriteElementString("nodo2", "texto del nodo2");
                writer.WriteEndElement();
                writer.WriteEndDocument();
                writer.Flush();
                writer.Close();
            }
            catch (MySqlException ex)
            {
                MessageBox.Show(ex.Message, "Error en conexión");
                Application.Exit();
                return;
            }
            conc.Close();
        }
        public void crearZIP(string path, string file)
        {
			int p = (int)Environment.OSVersion.Platform;
			if(p == 4 || p == 6 || p == 128)
			{
				// ver que proceso hacemos aca!
			}
			else
			{	// en linux mono esto no esta implementado ... no funca la creacion del zip
				string rp = Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) + "/zip/";
				file = file + ".ZIP";
				ZipFile.CreateFromDirectory(path, rp + file);
				//File.Delete(arbor);
			}
        }
        public void actualizaCDR(int idg,string idr, string estado)
        {
            //DataGridCell mycell = (DataGridCell)grilla.Rows[idg].Cells["cdr"].Value;
            grilla.Rows[idg].Cells["cdr"].Value = "8";
            MySqlConnection conn = new MySqlConnection(DB_CONN_STR);
            try
            {
                conn.Open();
                string actua = "update madocvtas set cdr=@cdr where id=@idr";
                MySqlCommand miact = new MySqlCommand(actua, conn);
                miact.Parameters.AddWithValue("@idr", idr);
                miact.Parameters.AddWithValue("@cdr", "8");
                miact.ExecuteNonQuery();
            }
            catch (MySqlException ex)
            {
                MessageBox.Show(ex.Message, "Error de conexión a la base de datos");
                Application.Exit();
                return;
            }
            conn.Close();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // miform
            // 
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Name = "miform";
            this.Load += new System.EventHandler(this.miform_Load);
            this.ResumeLayout(false);

        }

        private void miform_Load(object sender, EventArgs e)
        {

        }
	}
}

