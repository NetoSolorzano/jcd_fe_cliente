using System;
using System.Windows.Forms;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.Configuration;
using System.Data;
using System.Xml;
using System.Text;
using System.IO;
using System.IO.Compression;
using System.Globalization;
using MySql.Data.MySqlClient;
using Gma.QrCodeNet.Encoding;               // este no se si obligatorio
using Gma.QrCodeNet.Encoding.Windows.Render;       // codigo QR

namespace altiplano_fact_elect_clt
{
	public class miform : Form
	{
        static string nomform = "miform";
        NumLetra mel = new NumLetra();
        public string pubmail = "";     // correo electrónico del cliente que se visualiza y actualiza en el form documento
        public int idxgri = 0;          // index de la fila de la grilla seleccionada para su impresion
        string nom_imp = "";            // noombre de la impresora de tickets
        string gloser = "";             // glosa del servicio de carga que va en el detalle de los docs.vta.
        string leyen1, nuausu, leyen3, desped, despe2, provee, tasaigv, monesim, Cfactura, Cboleta, ubigeoe = "";
        string direcem, distemi, provemi, urbemis, depaemi, iFE, codlocsun = "";
        string ctadetr, glodetr, pordetr, mondetr = "";     // variables para la detracción
        DataTable sunat = new DataTable();                // codigos sunat y otros para la facturacion electronica
        #region conexion a la base de datos
        // own database connection
        static string serv = ConfigurationManager.AppSettings["serv"].ToString();
        static string port = ConfigurationManager.AppSettings["port"].ToString();
        static string usua = ConfigurationManager.AppSettings["user"].ToString();
        static string cont = ConfigurationManager.AppSettings["pass"].ToString();
        static string data = ConfigurationManager.AppSettings["data"].ToString();
        static string ctl = ConfigurationManager.AppSettings["ConnectionLifeTime"].ToString();
        static string sped = ConfigurationManager.AppSettings["spedvta"].ToString();    // series electronicas
        static string operador = ConfigurationManager.AppSettings["operador"].ToString();   // cajero/vendedor/operador
        static string tienda = ConfigurationManager.AppSettings["tienda"].ToString();   // codigo de la tienda/pto.venta
        static string tloa = ConfigurationManager.AppSettings["timeload"].ToString();   // milisegundos para jalar datos
        static int copias = int.Parse(ConfigurationManager.AppSettings["copiasTK"].ToString());    // cantidad de copias del ticket
        string DB_CONN_STR = "server=" + serv + ";port=" + port + ";uid=" + usua + ";pwd=" + cont + ";database=" + data + 
            ";ConnectionLifeTime=" + ctl + ";";
        string CONN_CLTE = "";
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
        private CheckBox chk_impr;  // marca si jala los documentos ya impresos, true = jala, false = no jala
        private CheckBox chk_auto;  // marca si jala los datos automaticamente, true = jala automatico, false = manual
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
        private TextBox tx_loc;     // tienda pto. de venta
        //
        private Panel marco4;           // marco para la grilla
        private DataGridView grilla;    // grilla de los datos actuales
        private DataGridView grillad;   // grilla para el detalle, invisible
        //
        private Panel marco5;           // botones de comando
        private Button bt_print;
        private Button bt_anu;

        #endregion
        #region declaracion de variables
        // main form defaults margins
        int ancho = 1080;
        int largo = 600;
        // temporizador
        private System.Windows.Forms.Timer temporizador;
        // datos del cliente del sistema
        private string nomclie = "";
        private string rucclie = "";
        private string dirclie = "";
        private string rasclie = "";
        private string corremi = "";
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
            this.grilla.CellDoubleClick += new DataGridViewCellEventHandler(grilla_doble_click);
            bt_anu.Visible = false;     // no anulamos nada de momento
            btn_.Focus();
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
        private void init()             // iniciamos valores no configurados por xml
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
            lb_serv = new Label { Text = "Servidor: ", Location = new Point(this.Left + 10, this.Top + 7), Width = 70, Anchor = (AnchorStyles.None) };
            tx_serv = new TextBox { Location = new Point(lb_serv.Width + 10, this.Top + 5), Width = 150, ReadOnly = true, Anchor = AnchorStyles.None };
            lb_base = new Label { Text = "Base datos: ", Location = new Point(300, this.Top + 7), Width = 80, Anchor = (AnchorStyles.None) };
            tx_base = new TextBox { Location = new Point(lb_base.Left + lb_base.Width + 10, this.Top + 5), Width = 100, ReadOnly = true, Anchor = AnchorStyles.None };
            lb_tabl = new Label { Text = "Tabla Vtas.: ", Location = new Point(tx_base.Left + tx_base.Width + 20, this.Top + 7), Width = 80, Anchor = (AnchorStyles.None) };
            tx_tabl = new TextBox { Location = new Point(lb_tabl.Left + lb_tabl.Width + 10, this.Top + 5), Width = 100, ReadOnly = true, Anchor = AnchorStyles.None };
            lb_tabd = new Label { Text = "Detalle Vtas.", Location = new Point(tx_tabl.Left + tx_tabl.Width + 20, this.Top + 7), Width=80, Anchor = AnchorStyles.None };
            tx_tabd = new TextBox { Location = new Point(lb_tabd.Left + lb_tabd.Width + 10, this.Top + 5), Width = 100, ReadOnly = true, Anchor = AnchorStyles.None };
            //
            lb_usua = new Label { Text = "Usuario: ", Location = new Point(this.Left + 10, this.Top + 35), Width = 70, Anchor = (AnchorStyles.None) };
            tx_usua = new TextBox { Location = new Point(lb_usua.Width + 10, this.Top + 33), Width = 150, ReadOnly = true, Anchor = (AnchorStyles.None) };
            lb_pass = new Label { Text = "Contraseña: ", Location = new Point(300, this.Top + 35), Width = 80, Anchor = (AnchorStyles.None) };
            tx_pass = new TextBox { Location = new Point(lb_pass.Left + lb_pass.Width + 10, this.Top + 35), ReadOnly = true, PasswordChar = '*', Width = 100, Anchor = (AnchorStyles.None) };
            lb_port = new Label { Text = "Puerto: ", Location = new Point(tx_pass.Left + tx_pass.Width + 20, this.Top + 35), Width = 80, Anchor = (AnchorStyles.None) };
            tx_port = new TextBox { Location = new Point(lb_port.Left + lb_port.Width + 10, this.Top + 35), Width = 100,ReadOnly = true, Anchor = (AnchorStyles.None) };
            //
            marco2 = new Panel();
            marco2.Location = new Point(this.Left, marco1.Height+5);
            marco2.Size = new Size(this.Width, 30);
            marco2.Anchor = (AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top);
            marco2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            marco2.BackColor = Color.AliceBlue;
            lb_letr = new Label { Text = "Letras: ", Location = new Point(this.Left + 10, this.Top + 5), Width = 70, Anchor = (AnchorStyles.None) };
            tx_letr = new TextBox { Location = new Point(lb_letr.Width + 10, this.Top + 3), Width = 100, ReadOnly = true, Anchor = (AnchorStyles.None) };
            lb_digs = new Label { Text = "#Digs.serie: ", Location = new Point(300, this.Top + 5), Width = 80, Anchor = (AnchorStyles.None) };
            tx_digs = new TextBox { Location = new Point(lb_digs.Left + lb_digs.Width + 10, this.Top + 5), Width = 100, ReadOnly = true, Anchor = (AnchorStyles.None) };
            lb_time = new Label { Text = "Tiempo (seg): ", Location = new Point(tx_digs.Left + tx_digs.Width + 20, this.Top + 5), Width = 80,Anchor = AnchorStyles.None };
            tx_time = new TextBox { Location = new Point(lb_time.Left + lb_time.Width + 10, this.Top + 5), Width = 100, ReadOnly = true, Anchor = AnchorStyles.None };
            chk_impr = new CheckBox { Text = "Jala impresos", Location = new Point(tx_time.Left + tx_time.Width + 100, this.Top + 5), Width = 100, Anchor = AnchorStyles.None };
            chk_auto = new CheckBox { Text = "Jala automático", Location = new Point(this.Width - 130, this.Top + 5), Width = 130, Anchor = AnchorStyles.None };
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
            lb_spe = new Label { Text = "Series Elect.: ", Location = new Point(tx_dia.Left + tx_dia.Width + 20, this.Top + 5), Width = 80 };
            tx_spe = new TextBox { Location = new Point(lb_spe.Left + lb_spe.Width + 10, this.Top + 3), Width = 50, ReadOnly = true, TextAlign = HorizontalAlignment.Center };
            lb_usu = new Label { Text = "Usuario: ", Location = new Point(tx_spe.Left + tx_spe.Width + 20, this.Top + 5), Width = 50 };
            tx_usu = new TextBox { Location = new Point(lb_usu.Left + lb_usu.Width + 10, this.Top + 3), Width = 100, ReadOnly = true, TextAlign = HorizontalAlignment.Center };
            lb_loc = new Label { Text = "Tienda: ", Location = new Point(tx_usu.Left + tx_usu.Width + 20, this.Top + 5), Width = 50 };
            tx_loc = new TextBox { Location = new Point(lb_loc.Left + lb_loc.Width + 10, this.Top + 3), Width = 100, ReadOnly = true, TextAlign = HorizontalAlignment.Center };
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
            marco2.Controls.Add(chk_impr);
            marco2.Controls.Add(chk_auto);
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
            marco3.Controls.Add(lb_usu);
            marco3.Controls.Add(tx_usu);
            marco3.Controls.Add(lb_loc);
            marco3.Controls.Add(tx_loc);
            marco3.Controls.Add(btn_);
            //
            marco4 = new Panel();
            marco4.Location = new Point(this.Left, marco3.Top + marco3.Height); //  + 5
            //marco4.Size = new Size(ancho, largo-10);
            marco4.Size = new Size(this.Width, this.Height - 30);
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
            grillad = new DataGridView();
            grillad.Visible = false;        // probamos con la grilla sin que este asociada a un control
            marco4.Controls.Add(grillad);
            //
            marco5 = new Panel();
            marco5.Location = new Point(this.Left, marco4.Top + marco4.Height);
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
            try
            {
                MySqlConnection conn = new MySqlConnection(DB_CONN_STR);
                conn.Open();
                string consulta = "select * from baseconf limit 1"; //  where referen0=@tda
                MySqlCommand micon = new MySqlCommand(consulta, conn);
                micon.Parameters.AddWithValue("@tda", tienda);
                MySqlDataReader dr = micon.ExecuteReader();
                if (dr.HasRows)
                {
                    if (dr.Read())
                    {
                        nomclie = dr.GetString("Cliente");                      // nombre comercial
                        rucclie = dr.GetString("Ruc");
                        dirclie = dr.GetString("direcc").Trim() + " - " + dr.GetString("distrit").Trim() + " - " + dr.GetString("departamento");
                        rasclie = dr.GetString("rasonsocial");
                        tasaigv = dr.GetString("igv");
                        ubigeoe = dr.GetString("referen1");                     // ubigeo del emisor
                        corremi = dr.GetString("referen3");                     // correo electronico del cliente
                        direcem = dr.GetString("direcc").Trim();
                        distemi = dr.GetString("distrit").Trim();
                        provemi = dr.GetString("provin").Trim();
                        urbemis = dr.GetString("referen2").Trim();              // urbanizacion
                        depaemi = dr.GetString("departamento").Trim();          // departamento
                        ctadetr = dr.GetString("ctadetra").Trim();              // cuenta en el BN para la detracción
                        glodetr = dr.GetString("glosadet").Trim();              // glosa de la detracción
                        pordetr = dr.GetString("detra").Trim();                 // porcentaje detraccion
                        mondetr = dr.GetString("valdetra").Trim();              // monto a partir del cual va la detracción
                        //
                    }
                    dr.Close();
                }
                else
                {
                    dr.Close();
                    conn.Close();
                    MessageBox.Show("No se ubica código de empresa", "Error fatal de config.");
                    Application.Exit();
                    return;
                }
                //
                consulta = "select campo,param,valor from enlaces where formulario=@nofo";
                micon = new MySqlCommand(consulta, conn);
                micon.Parameters.AddWithValue("@nofo", nomform);
                MySqlDataAdapter da = new MySqlDataAdapter(micon);
                DataTable dt = new DataTable();
                da.Fill(dt);
                for (int t = 0; t < dt.Rows.Count; t++)
                {
                    DataRow row = dt.Rows[t];
                    if (row["campo"].ToString() == "leyendas")
                    {
                        if (row["param"].ToString() == "1") leyen1 = row["valor"].ToString();             // leyenda1
                        if (row["param"].ToString() == "2") nuausu = row["valor"].ToString();             // autorizsunat
                        if (row["param"].ToString() == "3") leyen3 = row["valor"].ToString();             // leyenda3
                        if (row["param"].ToString() == "4") desped = row["valor"].ToString();             // despedida
                        if (row["param"].ToString() == "5") despe2 = row["valor"].ToString();             // despedida 2
                        if (row["param"].ToString() == "6") provee = row["valor"].ToString();             // pag. del proveedor
                        if (row["param"].ToString() == "servicio") gloser = row["valor"].ToString();             // glosa del servicio
                    }
                    if (row["campo"].ToString() == "docvta")
                    {
                        if (row["param"].ToString() == "factura") Cfactura = row["valor"].ToString();             // documento factura
                        if (row["param"].ToString() == "boleta") Cboleta = row["valor"].ToString();             // documento boleta
                    }
                    if (row["campo"].ToString() == "identificador")
                    {
                        if (row["param"].ToString() == "identif") iFE = row["valor"].ToString().Trim();             // identif. de fact. electrónica
                    }
                    if (row["campo"].ToString() == "impresora")
                    {
                        if (row["param"].ToString() == "nombre") nom_imp = row["valor"].ToString().Trim();             // nombre de la impresora
                    }
                }
                da.Dispose();
                dt.Dispose();
                consulta = "select idtabella,idcodice,descrizione,descrizionerid,codice2,codsunat from descrittive";
                micon = new MySqlCommand(consulta, conn);
                da = new MySqlDataAdapter(micon);
                da.Fill(sunat);

                da.Dispose();
                conn.Close();
            }
            catch (MySqlException ex)
            {
                MessageBox.Show(ex.Message, "Error de conexión");
                Application.Exit();
                return;
            }
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
                    if (row[0].ToString() == "letras") tx_letr.Text = iFE;  // row[2].ToString();
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
            tx_usu.Text = operador;
            tx_loc.Text = tienda;   // jalamos el local del usuario
        }

        private void jala_datos()       // obtiene los datos del cliente, segun fecha, usuario, local y serie
        {
            // conexion a la base del cliente, solo jala los datos del usuario, de la fecha y de su local
            string CONN_CLTE = "server=" + tx_serv.Text.Trim() + ";port=" + tx_port.Text.Trim() + ";uid=" + tx_usua.Text.Trim() + 
                ";pwd=" + tx_pass.Text.Trim() + ";database=" + tx_base.Text.Trim() + ";ConnectionLifeTime=" + ctl + ";";
            // jalamos los datos
            MySqlConnection conc = new MySqlConnection(CONN_CLTE);
            try
            {
                string ptoorig = "", ptodest = "";
                string parte = "";
                if (tx_dia.Text == "" && tx_mes.Text == "") parte = "";
                if (tx_dia.Text == "" && tx_mes.Text != "") parte = " month(a.fechope)=@mes and";
                if (tx_dia.Text != "" && tx_mes.Text != "") parte = " month(a.fechope)=@mes and day(a.fechope)=@dia and";
                if (tx_spe.Text != "") parte = parte + " a.servta in (" + tx_spe.Text.Trim() + ")";
                parte = parte + " and a.servta in (" + tx_spe.Text.Trim() + ")";
                //parte = parte + " and a.usercaja=@ucaj and a.local=@loca and a.cdr=' '";
                parte = parte + " and a.usercaja=@ucaj and a.numcaja=@loca and a.cdr=' '";
                //if (chk_impr.Checked == true) parte = parte + " and a.impreso='S'";
                //else parte = parte + " and a.impreso='N'";
                conc.Open();
                // obtenemos los datos del cliente, cabecera de ventas ....... ,a.clidv,a.dist,a.pedido,ifnull(b.email,'') as email,c.flagcol as ubigeo,
                string consulta = "select a.iddocvtas,a.usercaja,a.fechope,a.tipcam,a.numcaja as local,a.docvta,a.servta,a.corrvta,a.doccli,a.numdcli," +
                    "a.direc1,a.direc2,c.nomdep,p.nomprov,dt.nomdist,a.grem,a.notent,a.observ,a.observ2,a.moneda,a.subtot,a.igv,a.doctot,a.status,a.dia," +
                    "a.numdcli as codcli,'EFECTIVO' as mpago,b.nomclie as nomcli,b.telefono1,dt.nomdist as distAdq,p.nomprov as provAdq,c.nomdep as dptoAdq," +
                    "ifnull(de.codice2,'PE') as paisAdq,cl.ciudad as PTORIG,dl.ciudad as PTODEST " +
                    "from " + tx_base.Text.Trim() + "." + tx_tabl.Text.Trim() + " a " +
                    "left join peg.maclie b on b.docclie=a.doccli and b.numdoc=a.numdcli " +
                    "left join peg.dptos c on c.ncdep = b.depart " +
                    "left join peg.provin p on p.ncdep=c.ncdep and p.codprov=b.provin " +
                    "left join peg.distrit dt on dt.iddistrit=b.distrit and dt.coddep=b.provin " +
                    "left join altiplano_fe.descrittive de on de.idtabella='PAI' and de.idcodice=b.pais " +
                    "left join peg.manoen bm on concat(bm.sernoe,'-',bm.cornoe)=a.notent " +
                    "left join peg.locales cl on cl.codloc = bm.numcaja " +
                    "left join peg.locales dl on dl.codloc = bm.destino " +
                    "where year(a.fechope)=@yea and " + parte;
                //                     "left join pagos d on d.docvta=concat(a.docvta,'-',a.servta,'-',a.corrvta) and d.pedido=a.pedido " +
                //                     "left join desc_mpa e on e.idcodice=d.mpago " + 
                MySqlCommand micon = new MySqlCommand(consulta, conc);
                micon.Parameters.AddWithValue("@yea", tx_ano.Text);
                micon.Parameters.AddWithValue("@mes", tx_mes.Text);
                micon.Parameters.AddWithValue("@dia", tx_dia.Text);
                micon.Parameters.AddWithValue("@ucaj", tx_usu.Text);
                micon.Parameters.AddWithValue("@loca", tx_loc.Text);
                MySqlDataAdapter da = new MySqlDataAdapter(micon);
                DataTable dt = new DataTable();
                da.Fill(dt);
                // obtenemos los datos del cliente, detalle de ventas
                /*
                select a.*,c.ciudad,d.ciudad 
                from peg.docvtas a left join peg.manoen b on concat(b.sernoe,'-',b.cornoe)=a.notent
                left join peg.locales c on c.codloc=b.numcaja 
                left join peg.locales d on d.codloc=b.destino
                where a.docvta='BV' and a.servta='001' and a.corrvta='0000002';
                
                consulta = "select b.tipdv,b.servta,b.corvta,b.codprd,b.descrip,b.precio,b.cantid,b.total,b.unidad,b.peso," +
                    "a.numcaja as local,a.grem,a.status " +
                    "from " + tx_tabd.Text.Trim() + " b left join " + tx_tabl.Text.Trim() +
                    " a on a.docvta=b.tipdv and a.servta=b.servta and a.corrvta=b.corvta " +
                    "where year(a.fechope)=@yea and " + parte;
                */
                consulta = "select b.tipdv,b.servta,b.corvta,b.codprd,b.descrip,b.precio,b.cantid,b.total,b.unidad,b.peso," +
                    "a.numcaja as local,a.grem,a.status,ifnull(c.docremi,'') as gremit1,ifnull(c.docremi2,'') as gremit2 " +
                    "from detavtas b " +
                    "left join docvtas a on a.docvta=b.tipdv and a.servta=b.servta and a.corrvta=b.corvta " +
                    "left join magrem c on a.grem=concat(c.sergre,'-',c.corgre) " +
                    "where year(a.fechope)=@yea and " + parte;
                micon = new MySqlCommand(consulta, conc);
                micon.Parameters.AddWithValue("@yea", tx_ano.Text);
                micon.Parameters.AddWithValue("@mes", tx_mes.Text);
                micon.Parameters.AddWithValue("@dia", tx_dia.Text);
                micon.Parameters.AddWithValue("@ucaj", tx_usu.Text);
                micon.Parameters.AddWithValue("@loca", tx_loc.Text);
                MySqlDataAdapter dad = new MySqlDataAdapter(micon);
                DataTable dtd = new DataTable();
                dad.Fill(dtd);
                // insertamos la cabecera
                foreach (DataRow row in dt.Rows)
                {
                    //a.notent,
                    string vid = row["iddocvtas"].ToString();             // ID del registro
                    string usca = row["usercaja"].ToString();             // usuario de la caja
                    string vfec = string.Format("{0:yyyy-MM-dd}", row["fechope"]); // fecha del documento
                    string vtca = row["tipcam"].ToString();             // tipo de cambio
                    string vloc = row["local"].ToString();              // local donde se emitio el doc ... ALTIPLANO NUMCAJA
                    string vdvt = row["docvta"].ToString();             // cod.doc.vta
                    string vsvt = vdvt.Substring(0,1) + row["servta"].ToString();       // identif. de fact. elec. + serie
                    string vcvt = "0" + row["corrvta"].ToString();      // correlativo
                    string vdcl = row["doccli"].ToString();             // tip doc cliente
                    string vndc = row["numdcli"].ToString();            // num doc cliente
                    string vdic = row["direc1"].ToString();             // direcc. cliente
                    string vdc2 = row["direc2"].ToString();             // direcc. cliente
                    //string vpcl = row["grem"].ToString();
                    string[] nada = row["grem"].ToString().Split('-');  // guia de remision
                    string vpcl = "";
                    if (nada[0].Trim() != "" && nada[1].Trim() != "")
                    {
                        string xser = ("000" + nada[0]).Substring(("000" + nada[0]).Length - 4, 4);
                        string xcor = ("0000000" + nada[1]).Substring(("0000000" + nada[1]).Length - 8, 8);
                        vpcl = xser + "-" + xcor;     // formato SSSS-CCCCCCCC (4 serie, guion, 8 numero)             
                    }
                    string vdpt = row["nomdep"].ToString();             // nombre departamento
                    string vdis = row["nomdist"].ToString();            // nombre distrito
                    string vpro = row["nomprov"].ToString();            // nombre provincia
                    string vobs = row["observ"].ToString();             // comentarios del doc.vta.
                    string vob2 = row["observ2"].ToString();            // comentarios 2 del doc.vta.
                    string vmon = row["moneda"].ToString();             // moneda del documento
                    string vsub = row["subtot"].ToString();             // sub total del doc
                    string vigv = row["igv"].ToString();                // igv del doc
                    string vtot = row["doctot"].ToString();             // total del doc
                    string vsta = row["status"].ToString();             // situacion de doc
                    string vccl = row["codcli"].ToString();             // codigo del cliente
                    string vdtc = row["dia"].ToString();                // fecha de creacion del doc
                    string mpago = row["mpago"].ToString();             // medio de pago
                    string vcdv = row["nomcli"].ToString();             // nombre del adquiriente
                    string vrsc = row["nomcli"].ToString();             // razon social del adquiriente ... en Altiplano hay un solo nombre del Adq
                    string vtcl = row["telefono1"].ToString();          // telefono del adquiriente
                    string vdsc = row["nomdist"].ToString();            // nombre distrito
                    string vmai = "";                                   // correo electrónico
                    //
                    string ubigeoAdq = "";  // row["ubigeo"].ToString();                    // ubigeo del adquiriente
                    string tipDoc = equivalencias("DVT", vdvt, "sunat");            // tipoDocumento sunat
                    string tipMon = equivalencias("MON", vmon, "sunat");            // tipoMoneda sunat
                    monesim = equivalencias("MON", vmon, "cod2");                   // simbolo de la moneda
                    string numDocEm = rucclie;                                      // numeDocEmi
                    string tipDocEm = equivalencias("DOC", "ruc", "sunat");         // tipoDocEmi
                    string razSocEmi = rasclie;                                     // razonSocEmi
                    string corrEmisor = equivalencias("MAI", vloc, "cod2");       // correoEmisor
                    string codLocAnEm = equivalencias("LOC", vloc, "sunat");      // codiLocAnEmi
                    string tipoDocAdq = equivalencias("DOC", vdcl, "sunat");        // tipoDocAdq
                    //string totIgv = (double.Parse(vtot) * double.Parse(tasaigv) / 100).ToString();            // totalIgv
                    //string totIgv = (double.Parse(vtot) - (double.Parse(vtot) / (double.Parse(tasaigv) / 100 + 1))).ToString();            // totalIgv
                    decimal verdura = Math.Round(decimal.Parse(vtot) - (decimal.Parse(vtot) / (decimal.Parse(tasaigv) / 100 + 1)),2);
                    string totIgv = verdura.ToString();
                    //string totImp = (double.Parse(vtot) * double.Parse(tasaigv) / 100).ToString();             // totalImp
                    decimal verde = Math.Round(decimal.Parse(vtot)) - verdura;
                    vsub = verde.ToString();
                    string totImp = totIgv;             // totalImp
                    //string toValNetGra = (double.Parse(vtot) - (double.Parse(vtot) * double.Parse(tasaigv) / 100)).ToString(); // totValVtaNetGrav
                    string toValNetGra = (double.Parse(vtot) / (double.Parse(tasaigv) / 100 + 1)).ToString(); // totValVtaNetGrav

                    string totVen = vtot;                                          // totalVenta
                    string tipOper = equivalencias("SUN", "TIPFAC", "sunat");          // tipoOperac 
                    string nomComEmi = nomclie;                  // nombreComercialEmisor
                    string ubigeoEmi = ubigeoe;                  // ubigeoEmisor
                    string direcEmis = direcem;                  // direccionEmisor
                    string urbEmisor = urbemis;                  // urbanizacion
                    string provinEmi = provemi;                  // provinciaEmisor
                    string departEmi = depaemi;                  // departamentoEmisor
                    string distriEmi = distemi;                  // distritoEmisor
                    string paisEmisor = "PE";                 // paisEmisor
                    string razSocAdq = vcdv;                  // razonSocialAdquiriente
                    string codLeyen1 = equivalencias("SUN","NUMLET","sunat");                  // codigoLeyenda_1
                    string texLeyen1 = "SON: " + mel.Convertir(vtot, true) + " " +
                    equivalencias("MON", "Soles", "rid");   // 
                    string codAux401 = equivalencias("SUN", "IGV%", "sunat");                  // codigoAuxiliar40_1
                    string texAux401 = equivalencias("SUN", "IGV%", "cod2");                  // textoAuxiliar40_1
                    string horaEmis = "";                   // horaEmision
                    string disAdq = row["distAdq"].ToString();        // distrito adquiriente
                    string provAdq = row["provAdq"].ToString();        // provincia adquiriente
                    string dptoAdq = row["dptoAdq"].ToString();        // departamento adquiriente
                    string paisAdq = row["paisAdq"].ToString();        // pais adquiriente
                    string codTipIm = equivalencias("SUN", "IGV", "sunat");        // codigo tipo de impuesto
                    string nomTipIm = "IGV";        // nombre tipo de impuesto
                    string tipImpue = equivalencias("SUN", "IGV", "cod2");        // tipo de tributo
                    ptoorig = row["PTORIG"].ToString();          // origen del servicio
                    ptodest = row["PTODEST"].ToString();         // destino del servicio
                    codlocsun = equivalencias("LOC", vloc, "cod2");
                    // detracciones
                    string _valdet = "";                // valor calculado de la detraccion
                    string _codser = "";                // codigo del servicio de la detraccion
                    string _pordet = "";                // porcentaje de la detraccion
                    string _nuctbn = "";                // cuenta en el BN para la detraccion
                    if(double.Parse(totVen) >= double.Parse(mondetr))       // SE ASUME QUE TODAS LAS VENTAS SON EN SOLES
                    {
                        _valdet = Math.Round(decimal.Parse(totVen) * decimal.Parse(pordetr) / 100, 2).ToString(); // monto de la detraccion
                        _codser = "037";  // "027"                                                        // codigo del servicio detraccion
                        _pordet = pordetr;                                                       // porcentaje de la detraccion
                        _nuctbn = ctadetr;                                                      // cuenta en el BN para la detraccion
                        tipOper = equivalencias("SUN", "TIPFAC", "cod2");                       // tipoOperac sujeta a detraccion
                    }
                    // insertamos la tabla propia
                    if (valexis("fe_cabecera", vdvt, vsvt, vcvt) == false) // solo los registros que no existan ya en la tabla
                    {							// false = no existe el docvta en la tabla propia
                        string inserta = "insert into fe_cabecera (usercaja,fechope,tipcam,local,docvta,servta,corrvta,doccli,numdcli,direc1,direc2,nomcli," +
                            "telef,codcli,clidv,grem,observ,observ2,moneda,subtot,igv,doctot,status,dia,useranul,fechanul,marca1,dist,cdr,mail," +
                            "tasaigv,simboloMon,autorizsunat,leyenda1,leyenda3,leyenda4,leyenda5,leyenda6,ubigeoAdq," +
                            "tipoDocumento,tipoMoneda,numeDocEmi,tipoDocEmi,razonSocEmi,correoEmisor,codiLocAnEmi,tipoDocAdq," +
                            "totalImp,totValVtaNetGrav,totalIgv,totalVenta,tipoOperac,nomComEmi,ubigeoEmi,direcEmis,urbEmisor,provinEmi," +
                            "departEmi,distriEmi,paisEmisor,razSocAdq,codLeyen1,texLeyen1,codAux401,texAux401,horaEmis,medPagoAdq," +
                            "distAdq,provAdq,dptoAdq,paisAdq,codTipoImp,nomTipoImp,tipoTributo," +
                            "pordetra,valdetra,codserd,ctadetra,codlocsun) " +
                            "values (" +
                            "@usca,@vfec,@vtca,@vloc,@vdvt,@vsvt,@vcvt,@vdcl,@vndc,@vdic,@vdc2,@vrsc," +
                            "@vtcl,@vccl,@vcdv,@vpcl,@vobs,@vob2,@vmon,@vsub,@vigv,@vtot,@vsta,@vdtc,@vuan,@vfan,@vma1,@vdsc,@cdr,@mai," +
                            "@tasaigv,@monesim,@nuausu,@leyen1,@leyen3,@desped,@despe2,@provee,@ubigeoAdq," +
                            "@tipDoc,@tipMon,@numDocEm,@tipDocEm,@razSocEmi,@corrEmisor,@codLocAnEm,@tipoDocAdq,@totImp,@toValNetGra,@totIgv," +
                            "@totVen,@tipOper,@nomComEmi,@ubigeoEmi,@direcEmis,@urbEmisor,@provinEmi," +
                            "@departEmi,@distriEmi,@paisEmisor,@razSocAdq,@codLeyen1,@texLeyen1,@codAux401,@texAux401,@horaEmis,@mpago," +
                            "@disAdq,@provAdq,@dptoAdq,@paisAdq,@codTipIm,@nomTipIm,@tipImpue," +
                            "@_porde,@_valde,@_codse,@_nucta,@codlocsun)";
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
                            mins.Parameters.AddWithValue("@vdc2", vdc2);    // direccion 2
                            mins.Parameters.AddWithValue("@vrsc", vrsc);    // nombre del cliente
                            mins.Parameters.AddWithValue("@vtcl", vtcl);    // telefono
                            mins.Parameters.AddWithValue("@vccl", vccl);    // codigo cliente
                            mins.Parameters.AddWithValue("@vcdv", vcdv);    // cliente doc venta
                            mins.Parameters.AddWithValue("@vpcl", vpcl);    // pedido/contrato con el cliente
                            mins.Parameters.AddWithValue("@vobs", vobs);    // observaciones
                            mins.Parameters.AddWithValue("@vob2", vob2);    // observaciones 2
                            mins.Parameters.AddWithValue("@vmon", vmon);    // moneda
                            mins.Parameters.AddWithValue("@vsub", vsub);    // sub total
                            mins.Parameters.AddWithValue("@vigv", vigv);    // igv
                            mins.Parameters.AddWithValue("@vtot", vtot);    // total    
                            mins.Parameters.AddWithValue("@vsta", vsta);    // estado del doc
                            mins.Parameters.AddWithValue("@vdtc", vdtc);    // fecha creacion del doc
                            mins.Parameters.AddWithValue("@vuan", "");      // vuan  usuario que anulo doc
                            mins.Parameters.AddWithValue("@vfan", DBNull.Value);        // vfan  fecha de anulacion
                            mins.Parameters.AddWithValue("@vma1", "");                  // vma1 marca 1
                            mins.Parameters.AddWithValue("@vdsc", vdsc);                // distrito
                            mins.Parameters.AddWithValue("@cdr", "9");                  // cdr
                            mins.Parameters.AddWithValue("@mai", vmai);                 // correo elect
                            mins.Parameters.AddWithValue("@mpago", mpago);              // medio de pago
                            mins.Parameters.AddWithValue("@tasaigv", tasaigv);          // tasaigv
                            mins.Parameters.AddWithValue("@monesim", monesim);          // monesim      simbolo moneda
                            mins.Parameters.AddWithValue("@leyen1", leyen1);            // leyen1
                            mins.Parameters.AddWithValue("@nuausu", nuausu);            // nuausu
                            mins.Parameters.AddWithValue("@leyen3", leyen3);            // leyen3
                            mins.Parameters.AddWithValue("@desped", desped);            // desped
                            mins.Parameters.AddWithValue("@despe2", despe2);            // despe2
                            mins.Parameters.AddWithValue("@provee", provee);            // provee
                            mins.Parameters.AddWithValue("@ubigeoAdq", ubigeoAdq);      // ubigeo adquiriente
                            mins.Parameters.AddWithValue("@tipDoc", tipDoc);            // tipoDocumento
                            mins.Parameters.AddWithValue("@tipMon", tipMon);            // tipoMoneda
                            mins.Parameters.AddWithValue("@numDocEm", numDocEm);        // numeDocEmi
                            mins.Parameters.AddWithValue("@tipDocEm", tipDocEm);        // tipoDocEmi
                            mins.Parameters.AddWithValue("@razSocEmi", razSocEmi);      // razonSocEmi
                            mins.Parameters.AddWithValue("@corrEmisor", corrEmisor);    // correoEmisor
                            mins.Parameters.AddWithValue("@codLocAnEm", codLocAnEm);    // codiLocAnEmi
                            mins.Parameters.AddWithValue("@tipoDocAdq", tipoDocAdq);    // tipoDocAdq
                            mins.Parameters.AddWithValue("@totImp", totImp);            // totalImp
                            mins.Parameters.AddWithValue("@toValNetGra", toValNetGra);  // totValVtaNetGrav
                            mins.Parameters.AddWithValue("@totIgv", totIgv);            // totalIgv
                            mins.Parameters.AddWithValue("@totVen", totVen);            // totalVenta
                            mins.Parameters.AddWithValue("@tipOper", tipOper);          // tipoOperac 
                            mins.Parameters.AddWithValue("@nomComEmi", nomComEmi);      // nombreComercialEmisor
                            mins.Parameters.AddWithValue("@ubigeoEmi", ubigeoEmi);      // ubigeoEmisor
                            mins.Parameters.AddWithValue("@direcEmis", direcEmis);      // direccionEmisor
                            mins.Parameters.AddWithValue("@urbEmisor", urbEmisor);      // urbanizacion
                            mins.Parameters.AddWithValue("@provinEmi", provinEmi);      // provinciaEmisor
                            mins.Parameters.AddWithValue("@departEmi", departEmi);      // departamentoEmisor
                            mins.Parameters.AddWithValue("@distriEmi", distriEmi);      // distritoEmisor
                            mins.Parameters.AddWithValue("@paisEmisor", paisEmisor);    // paisEmisor
                            mins.Parameters.AddWithValue("@razSocAdq", razSocAdq);      // razonSocialAdquiriente
                            mins.Parameters.AddWithValue("@codLeyen1", codLeyen1);      // codigoLeyenda_1
                            mins.Parameters.AddWithValue("@texLeyen1", texLeyen1);      // textoLeyenda_1
                            mins.Parameters.AddWithValue("@codAux401", codAux401);      // codigoAuxiliar40_1
                            mins.Parameters.AddWithValue("@texAux401", texAux401);      // textoAuxiliar40_1
                            mins.Parameters.AddWithValue("@horaEmis", horaEmis);        // horaEmision
                            mins.Parameters.AddWithValue("@disAdq", disAdq);            // distrito adquiriente
                            mins.Parameters.AddWithValue("@provAdq", provAdq);          // provincia adquiriente
                            mins.Parameters.AddWithValue("@dptoAdq", dptoAdq);          // departamento adquiriente
                            mins.Parameters.AddWithValue("@paisAdq", paisAdq);          // pais adquiriente
                            mins.Parameters.AddWithValue("@codTipIm", codTipIm);        // codigo tipo de impuesto
                            mins.Parameters.AddWithValue("@nomTipIm", nomTipIm);        // nombre tipo de impuesto
                            mins.Parameters.AddWithValue("@tipImpue", tipImpue);        // tipo de tributo
                            mins.Parameters.AddWithValue("@_porde", _pordet);           // porcentaje de detraccion
                            mins.Parameters.AddWithValue("@_valde", _valdet);           // valor de la detraccion a pagar
                            mins.Parameters.AddWithValue("@_codse", _codser);           // codigo de 
                            mins.Parameters.AddWithValue("@_nucta", _nuctbn);           // cuenta para detraccion en BN
                            mins.Parameters.AddWithValue("@codlocsun", codlocsun);      // codigo local anexo sunat  V.3006
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
                        string actua = "update " + tx_tabl.Text.Trim() + " set cdr=@cdr where iddocvtas=@vid";
                        MySqlCommand miact = new MySqlCommand(actua, conc);
                        miact.Parameters.AddWithValue("@cdr", "9");
                        miact.Parameters.AddWithValue("@vid", vid);
                        miact.ExecuteNonQuery();
                    }
                }
                // insertamos el detalle 
                int nuor = 0;
                string docum = "";
                foreach (DataRow rowd in dtd.Rows)
                {
                    if (docum == rowd["tipdv"].ToString() + iFE + rowd["servta"].ToString() + rowd["corvta"].ToString()) nuor = nuor + 1;
                    else 
                    { 
                        nuor = 1;
                        docum = rowd["tipdv"].ToString() + iFE + rowd["servta"].ToString() + rowd["corvta"].ToString();
                    }
                    string vdvtd = rowd["tipdv"].ToString(); 	        // cod.doc.vta
                    string vsvtd = vdvtd.Substring(0,1) + rowd["servta"].ToString(); 	// serie
                    string vcvtd = "0" + rowd["corvta"].ToString();     // correlativo
                    string vcpdd = rowd["codprd"].ToString();           // codigo del producto
                    string descd = rowd["descrip"].ToString();          // descrip
                    string vpred = rowd["total"].ToString();            // precio en caso Altiplano = total      
                    string vcand = rowd["cantid"].ToString();           // cantidad 
                    string vtotd = rowd["total"].ToString();            // total del doc
                    string vtodd = "0.00";                              // total del doc dolares
                    if (string.IsNullOrEmpty(vpred) || string.IsNullOrWhiteSpace(vpred)) vpred = "0";
                    if (string.IsNullOrEmpty(vcand) || string.IsNullOrWhiteSpace(vcand)) vcand = "0";
                    if (string.IsNullOrEmpty(vtotd) || string.IsNullOrWhiteSpace(vtotd)) vtotd = "0";
                    if (string.IsNullOrEmpty(vtodd) || string.IsNullOrWhiteSpace(vtodd)) vpred = "0";
                    string vlocd = rowd["local"].ToString();   // local
                    string vpedd = rowd["grem"].ToString();   // guia de remision
                    string vstad = rowd["status"].ToString();   // situacion de doc
                    string grr1 = rowd["gremit1"].ToString();             // guia remitente 1
                    string grr2 = rowd["gremit2"].ToString();             // guia remitente 2
                    string vmard = "";   // marca
                    string tipoDocEmi = equivalencias("DOC", "ruc", "sunat");            // tipoDocumentoEmisor
                    string numeDocEmi = rucclie;                                         // numeroDocumentoEmisor
                    string tipDocumen = equivalencias("DVT", vdvtd, "sunat");            // tipoDocumento
                    string serieNumer = vsvtd + vcvtd;                                   // serieNumero
                    string vtasaigv = tasaigv;                                           // tasaIGV 
                    string numOrdItem = nuor.ToString();                                 // numeroOrdenItem
                    string codProducto = "";
                    string codProdSUNAT = "";
                    string descripcion = gloser + " " + vcand  + " " + rowd["unidad"].ToString() + "(S) " +
                        ptoorig + " - " + ptodest + " S/GR:" + grr1.Trim() + grr2.Trim() + " " + descd.Trim()  ;                         // descripcion
                    if (descripcion.Trim().Length > 200) descripcion = descripcion.Substring(0, 199);
                    string cantidad = "1";                                            // cantidad (vcand;) TRATANDOSE DE SERVICIOS ES 1
                    string unidadMed = equivalencias("SUN", "CODUME","sunat");             // unidadMedida
                    string impUniConImp = vpred;                                           // importeUnitarioConImpuesto
                    //if (int.Parse(vcand) > 1) impUniConImp = (double.Parse(vpred) / int.Parse(vcand)).ToString();
                    //else impUniConImp = vpred;
                    //string impUniSinImp = (double.Parse(impUniConImp)/((double.Parse(tasaigv)/100)+1)).ToString(); // importeUnitarioSinImpuesto
                    //string impTotSinImp = (double.Parse(impUniSinImp) * int.Parse(vcand)).ToString();              // importeTotalSinImpuesto
                    decimal verdad = Math.Round((decimal.Parse(impUniConImp) / ((decimal.Parse(tasaigv) / 100) + 1)),2);
                    string impUniSinImp = verdad.ToString();
                    string impTotSinImp = impUniSinImp;              // caso altiplano cantidad = 1, total sin impuesto = unitario sin impuesto
                    string codImUnConIm = "";                        // codigoImporteUnitarioConImpues
                    string montoBaseIgv = "";                        // montoBaseIgv
                    //string importeIgv = (double.Parse(vpred) * double.Parse(tasaigv) / 100).ToString();             // importeIgv
                    string importeIgv = (double.Parse(impUniConImp) - double.Parse(impUniSinImp)).ToString();  // importeIgv
                    //string impTotalImp = (double.Parse(vpred) * double.Parse(tasaigv) / 100).ToString();            // importeTotalImpuestos
                    string impTotalImp = importeIgv;  // importeTotalImpuestos
                    string codRazonExo = "";                         // codigoRazonExoneracion
                    string codigoIgv = equivalencias("SUN", "TIPIGV", "sunat");
                    string identifIgv = equivalencias("SUN", "IGV", "sunat");
                    string nombreIgv = "IGV";
                    string tipoTribIgv = equivalencias("SUN", "IGV", "cod2");
                    if (valexis("fe_detalle", vdvtd, vsvtd, vcvtd) == false) // solo los registros que no existan ya en la tabla
                    {
                        // insertamos la tabla propia
                        string insertad = "insert into fe_detalle (" +
                            "tipdv,servta,corvta,codprd,descrip,precio,cantid,total,totdol,local,pedido,status,marca1," +
                            "tipoDocEmi,numeDocEmi,tipoDocumento,serieNumero,tasaigv,numOrdenItem,codigoProducto,codigoProductoSUNAT,descripcion," +
                            "cantidad,unidadMed,impTotSinImp,impUniSinImp,impUniConImp,codImpUniConImp,montoBaseIgv,importeIgv,impTotalImp,codRazonExo," +
                            "codigoIgv,identifIgv,nombreIgv,tipoTribIgv) " +
                            "values (" +
                            "@vdvt,@vsvt,@vcvt,@vcpd,@desc,@vpre,@vcan,@vtot,@vtod,@vloc,@vped,@vsta,@vmar," +
                            "@tipoDocEmi,@numeDocEmi,@tipDocumen,@serieNumer,@vtasaigv,@numOrdItem,@codProducto,@codProdSUNAT,@descripcion," +
                            "@cantidad,@unidadMed,@impTotSinImp,@impUniSinImp,@impUniConImp,@codImUnConIm,@montoBaseIgv,@importeIgv,@impTotalImp,@codRazonExo," +
                            "@codigoIgv,@identifIgv,@nombreIgv,@tipoTribIgv)";
                        MySqlConnection conn = new MySqlConnection(DB_CONN_STR);
                        try
                        {
                            conn.Open();
                            MySqlCommand mins = new MySqlCommand(insertad, conn);
                            mins.Parameters.AddWithValue("@vdvt", vdvtd);    // cod.doc.vta
                            mins.Parameters.AddWithValue("@vsvt", vsvtd);    // serie
                            mins.Parameters.AddWithValue("@vcvt", vcvtd);    // correlativo
                            mins.Parameters.AddWithValue("@vcpd", vcpdd);    // codigo del producto
                            mins.Parameters.AddWithValue("@desc", descd);    // descipcion
                            mins.Parameters.AddWithValue("@vpre", vpred);    // precio 
                            mins.Parameters.AddWithValue("@vcan", vcand);    // cantidad
                            mins.Parameters.AddWithValue("@vtot", vtotd);    // total del doc
                            mins.Parameters.AddWithValue("@vtod", vtodd);    // total del doc dolares
                            mins.Parameters.AddWithValue("@vloc", vlocd);    // local
                            mins.Parameters.AddWithValue("@vped", vpedd);    // pedido/contrato
                            mins.Parameters.AddWithValue("@vsta", vstad);    // situacion de doc
                            mins.Parameters.AddWithValue("@vmar", vmard);    // marca
                            mins.Parameters.AddWithValue("@tipoDocEmi", tipoDocEmi);
                            mins.Parameters.AddWithValue("@numeDocEmi", numeDocEmi);
                            mins.Parameters.AddWithValue("@tipDocumen", tipDocumen);
                            mins.Parameters.AddWithValue("@serieNumer", serieNumer);
                            mins.Parameters.AddWithValue("@vtasaigv", vtasaigv);
                            mins.Parameters.AddWithValue("@numOrdItem", numOrdItem);
                            mins.Parameters.AddWithValue("@codProducto", codProducto);
                            mins.Parameters.AddWithValue("@codProdSUNAT", codProdSUNAT);
                            mins.Parameters.AddWithValue("@descripcion", descripcion);
                            mins.Parameters.AddWithValue("@cantidad", cantidad);
                            mins.Parameters.AddWithValue("@unidadMed", unidadMed);
                            mins.Parameters.AddWithValue("@impTotSinImp", impTotSinImp);
                            mins.Parameters.AddWithValue("@impUniSinImp", impUniSinImp);
                            mins.Parameters.AddWithValue("@impUniConImp", impUniConImp);
                            mins.Parameters.AddWithValue("@codImUnConIm", codImUnConIm);
                            mins.Parameters.AddWithValue("@montoBaseIgv", montoBaseIgv);
                            mins.Parameters.AddWithValue("@importeIgv", importeIgv);
                            mins.Parameters.AddWithValue("@impTotalImp", impTotalImp);
                            mins.Parameters.AddWithValue("@codRazonExo", codRazonExo);
                            mins.Parameters.AddWithValue("@codigoIgv", codigoIgv);
                            mins.Parameters.AddWithValue("@identifIgv", identifIgv);
                            mins.Parameters.AddWithValue("@nombreIgv", nombreIgv);
                            mins.Parameters.AddWithValue("@tipoTribIgv", tipoTribIgv);
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
            if (tx_mes.Text != "00") condicion = " and month(b.fechope)=@mes";
			if (tx_dia.Text != "00") condicion = condicion + " and day(b.fechope)=@dia";
            //if (tx_spe.Text != "") condicion = condicion + " and left(b.servta,1) in ('F','B') and b.servta=@ser";    // " and b.servta in ('" + tx_letr.Text + tx_spe.Text.Trim() + "')";
            //if (tx_spe.Text != "") condicion = condicion + " and right(b.servta,3)=@ser";    // " and b.servta in ('" + tx_letr.Text + tx_spe.Text.Trim() + "')";
            condicion = condicion + " and b.usercaja=@ucaj and b.local=@loca";
            if (chk_impr.Checked == true) condicion = condicion + "";
            else condicion = condicion + " and b.impreso='N'";
            MySqlConnection conn = new MySqlConnection(DB_CONN_STR);
            try
            {
                conn.Open();    //  as numeroDocumentoAdquiriente
                string consulta = "select b.fechope,b.tipcam,b.docvta,b.servta,b.corrvta,b.doccli,b.numdcli,b.direc1,b.direc2,b.nomcli," +
                        "b.telef,b.codcli,b.clidv,b.pedido,b.observ,b.observ2,b.moneda,b.subtot,b.igv,b.doctot,b.status,date_format(b.dia,'%Y-%m-%d')," +
                        "b.useranul,date_format(b.fechanul,'%Y-%m-%d'),b.marca1,b.dist,b.cdr,b.impreso,b.mail as correoAdquiriente,b.local,b.ubigeoAdq," +
                        "b.tasaigv,b.simboloMon,b.autorizsunat,b.leyenda1,b.leyenda3,b.leyenda4,b.leyenda5,b.leyenda6," +
                        "b.tipoDocumento,b.tipoMoneda,b.numeDocEmi,b.tipoDocEmi,b.razonSocEmi,b.correoEmisor,b.codiLocAnEmi,b.tipoDocAdq," +
                        "b.totalImp,b.totValVtaNetGrav,b.totalIgv,b.totalVenta,b.tipoOperac,b.nomComEmi,b.ubigeoEmi,b.direcEmis,b.urbEmisor,b.provinEmi," +
                        "b.departEmi,b.distriEmi,b.paisEmisor,b.razSocAdq,b.codLeyen1,b.texLeyen1,b.codAux401,b.texAux401,b.horaEmis,b.medPagoAdq " +
                        "from fe_cabecera b where year(b.fechope)=@yea" + condicion;
                MySqlCommand micon = new MySqlCommand(consulta,conn);
				micon.Parameters.AddWithValue("@yea", tx_ano.Text);
				micon.Parameters.AddWithValue("@mes", tx_mes.Text);
				micon.Parameters.AddWithValue("@dia", tx_dia.Text);
                micon.Parameters.AddWithValue("@ucaj", tx_usu.Text);
                micon.Parameters.AddWithValue("@loca", tx_loc.Text);
                micon.Parameters.AddWithValue("@ser", tx_serv.Text);
                MySqlDataAdapter da = new MySqlDataAdapter(micon);
				DataTable dt = new DataTable();
				da.Fill(dt);
                if (dt.Rows.Count > 0)
                {
                    grilla.DataSource = dt;
                    for (int x = 30; x < dt.Columns.Count; x++)                     // ocultamos campos equivalentes
                    {
                        grilla.Columns[x].Visible = false;
                    }
                }
                consulta = "select a.local,a.codprd,a.descrip,a.precio,a.cantid,a.total,a.precdol,a.totdol,a.pedido,a.status,a.marca1,a.tipdv,a.servta,a.corvta," +
                    "a.tipoDocEmi,a.numeDocEmi,a.tipoDocumento,a.serieNumero,a.tasaigv,a.numOrdenItem,a.codigoProducto,a.codigoProductoSUNAT,a.descripcion," +
                    "a.cantidad,a.unidadMed,a.impTotSinImp,a.impUniSinImp,a.impUniConImp,a.codImpUniConImp,a.montoBaseIgv,a.importeIgv,a.impTotalImp,a.codRazonExo " +
                    "from fe_detalle a left join fe_cabecera b on concat(a.tipdv,a.servta,a.corvta)=concat(b.docvta,b.servta,b.corrvta) " +
                    "where year(b.fechope)=@yea" + condicion;
                micon = new MySqlCommand(consulta, conn);
                micon.Parameters.AddWithValue("@yea", tx_ano.Text);
                micon.Parameters.AddWithValue("@mes", tx_mes.Text);
                micon.Parameters.AddWithValue("@dia", tx_dia.Text);
                micon.Parameters.AddWithValue("@ucaj", tx_usu.Text);
                micon.Parameters.AddWithValue("@loca", tx_loc.Text);
                micon.Parameters.AddWithValue("@ser", tx_serv.Text);
                da = new MySqlDataAdapter(micon);
                DataTable dtd = new DataTable();
                da.Fill(dtd);
                if (dtd.Rows.Count > 0) grillad.DataSource = dtd;
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
        private void grilla_doble_click(object sender, DataGridViewCellEventArgs e)  // doble click para seleccionar e imprimir
        {
            idxgri = e.RowIndex;
            documento(idxgri);
        }
        private void documento(int fila)                                            // visualiza doc.vta. pide correo e imprime
        {
            Form documento = new Form();
            documento.Name = "documento";
            documento.Text = "Documento de Venta Electrónico";
            documento.Width = 500;
            documento.Height = 600;
            documento.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            documento.MinimizeBox = false;
            documento.MaximizeBox = false;
            documento.AutoSize = false;
            documento.BackColor = Color.WhiteSmoke;
            documento.StartPosition = FormStartPosition.CenterScreen;
            //
            int esver = 30;                 // espacio vertical entre lineas
            string tipdo = "";              // tipo de documento
            string serie = "";              // serie
            string corre = "";              // correlativo
            Label lb_fec = new Label();
            Label tx_fec = new Label();
            lb_fec = new Label { Text = "Fecha: ", Location = new Point(10, 10), Width = 140, Font = new Font("Arial",15.0F), Anchor = (AnchorStyles.None) };
            tx_fec = new Label { Location = new Point(lb_fec.Left + lb_fec.Width + 10, 10), Width = 200, Font = new Font("Arial", 15.0F), Anchor = AnchorStyles.None };
            //
            Label lb_doc = new Label();
            Label tx_doc = new Label();
            lb_doc = new Label { Text = "Documento: ", Location = new Point(10, lb_fec.Top + esver), Width = 140, Font = new Font("Arial", 15.0F), Anchor = (AnchorStyles.None) };
            tx_doc = new Label { Location = new Point(lb_doc.Left + lb_doc.Width + 10, tx_fec.Top + esver), Width = 200, Font = new Font("Arial", 15.0F), Anchor = AnchorStyles.None };
            //
            Label lb_cli = new Label();
            Label tx_cli = new Label();
            lb_cli = new Label { Text = "Cliente: ", Location = new Point(10, lb_doc.Top + esver), Width = 140, Font = new Font("Arial", 15.0F), Anchor = (AnchorStyles.None) };
            tx_cli = new Label { Location = new Point(lb_cli.Left + lb_cli.Width + 10, tx_doc.Top + esver), Width = 400, Font = new Font("Arial", 13.0F), Anchor = AnchorStyles.None };
            //
            Label lb_ruc = new Label();
            Label tx_ruc = new Label();
            lb_ruc = new Label { Text = "Ruc/Dni: ", Location = new Point(10, lb_cli.Top + esver), Width = 140, Font = new Font("Arial", 15.0F), Anchor = (AnchorStyles.None) };
            tx_ruc = new Label { Location = new Point(lb_ruc.Left + lb_ruc.Width + 10, tx_cli.Top + esver), Width = 400, Font = new Font("Arial", 15.0F), Anchor = AnchorStyles.None };
            //
            Label lb_dir = new Label();
            Label tx_dir = new Label();
            lb_dir = new Label { Text = "Dirección: ", Location = new Point(10, lb_ruc.Top + esver), Width = 140, Font = new Font("Arial", 15.0F), Anchor = (AnchorStyles.None) };
            tx_dir = new Label { Location = new Point(lb_dir.Left + lb_dir.Width + 10, tx_ruc.Top + esver), Width = 400, Font = new Font("Arial", 13.0F), Anchor = AnchorStyles.None };
            //
            Label lb_dis = new Label();
            Label tx_dis = new Label();
            lb_dis = new Label { Text = "Distrito: ", Location = new Point(10, lb_dir.Top + esver), Width = 140, Font = new Font("Arial", 15.0F), Anchor = (AnchorStyles.None) };
            tx_dis = new Label { Location = new Point(lb_dis.Left + lb_dis.Width + 10, tx_dir.Top + esver), Width = 400, Font = new Font("Arial", 15.0F), Anchor = AnchorStyles.None };
            //
            Label lb_pro = new Label();
            Label tx_pro = new Label();
            lb_pro = new Label { Text = "Prov./Región: ", Location = new Point(10, lb_dis.Top + esver), Width = 140, Font = new Font("Arial", 15.0F), Anchor = (AnchorStyles.None) };
            tx_pro = new Label { Location = new Point(lb_pro.Left + lb_pro.Width + 10, tx_dis.Top + esver), Width = 400, Font = new Font("Arial", 15.0F), Anchor = AnchorStyles.None };

            // datos de cabecera    docvta,servta,corrvta
            tx_fec.Text = grilla.Rows[fila].Cells["fechope"].Value.ToString().Substring(0, 10);
            tx_doc.Text = grilla.Rows[fila].Cells["docvta"].Value.ToString() + "-" +
                grilla.Rows[fila].Cells["servta"].Value.ToString() + "-" +
                grilla.Rows[fila].Cells["corrvta"].Value.ToString();
            tx_cli.Text = grilla.Rows[fila].Cells["clidv"].Value.ToString();    // antes clidv
            tx_ruc.Text = grilla.Rows[fila].Cells["numdcli"].Value.ToString();
            tx_dir.Text = grilla.Rows[fila].Cells["direc1"].Value.ToString().Trim() + grilla.Rows[fila].Cells["direc2"].Value.ToString().Trim();
            tx_dis.Text = grilla.Rows[fila].Cells["direc2"].Value.ToString();
            tx_pro.Text = grilla.Rows[fila].Cells["dist"].Value.ToString();
            tipdo = grilla.Rows[fila].Cells["docvta"].Value.ToString();
            serie = grilla.Rows[fila].Cells["servta"].Value.ToString();
            corre = grilla.Rows[fila].Cells["corrvta"].Value.ToString();
            // detalle del documento
            DataGridView gd = new DataGridView();
            gd = new DataGridView { Width = this.Width - 4, Height = 100, Location = new Point(0,tx_pro.Top + esver), ColumnCount = 4, ReadOnly = true };
            gd.Columns[0].Width = 100;   // codigo
            gd.Columns[1].Width = 300;   // descrip
            gd.Columns[2].Width = 40;   // cant
            gd.Columns[3].Width = 60;   // precio
            //foreach (DataGridViewRow row in grillad.Rows)
            for ( int i = 0; i < grillad.Rows.Count - 1; i++ )
            {
                DataGridViewRow row = grillad.Rows[i];
                if (row.Cells["tipdv"].Value.ToString() == tipdo && row.Cells["servta"].Value.ToString() == serie && row.Cells["corvta"].Value.ToString() == corre)
                {
                    gd.Rows.Add(row.Cells["codprd"].Value.ToString(), row.Cells["descripcion"].Value.ToString(),
                        row.Cells["cantid"].Value.ToString(), row.Cells["total"].Value.ToString());
                }
            }
            // pie del documento
            Label lb_sub = new Label();
            Label tx_sub = new Label();
            lb_sub = new Label { Text = "Sub Total: ", Location = new Point(200, gd.Bottom + esver), Width = 140, Font = new Font("Arial", 15.0F), Anchor = (AnchorStyles.None) };
            tx_sub = new Label { Location = new Point(lb_sub.Left + lb_sub.Width + 10, gd.Bottom + esver), Width = 100, TextAlign = ContentAlignment.MiddleRight, Font = new Font("Arial", 15.0F), Anchor = AnchorStyles.None };
            //
            Label lb_igv = new Label();
            Label tx_igv = new Label();
            lb_igv = new Label { Text = "Igv: ", Location = new Point(200, lb_sub.Top + esver), Width = 140, Font = new Font("Arial", 15.0F), Anchor = (AnchorStyles.None) };
            tx_igv = new Label { Location = new Point(lb_igv.Left + lb_igv.Width + 10, lb_sub.Top + esver), Width = 100, TextAlign = ContentAlignment.MiddleRight, Font = new Font("Arial", 15.0F), Anchor = AnchorStyles.None };
            //
            Label lb_tot = new Label();
            Label tx_tot = new Label();
            lb_tot = new Label { Text = "Total: ", Location = new Point(200, lb_igv.Top + esver), Width = 140, Font = new Font("Arial", 15.0F), Anchor = (AnchorStyles.None) };
            tx_tot = new Label { Location = new Point(lb_tot.Left + lb_tot.Width + 10, lb_igv.Top + esver), Width = 100, TextAlign = ContentAlignment.MiddleRight, Font = new Font("Arial", 15.0F), Anchor = AnchorStyles.None };
            //
            tx_sub.Text = grilla.Rows[fila].Cells["totValVtaNetGrav"].Value.ToString();   // subtot
            tx_igv.Text = grilla.Rows[fila].Cells["totalIgv"].Value.ToString();     // igv
            tx_tot.Text = grilla.Rows[fila].Cells["totalVenta"].Value.ToString();   // 
            //
            Label lb_mail = new Label();
            TextBox tx_mail = new TextBox();
            lb_mail = new Label { Text = "Correo Electrónico: ", Location = new Point(10, tx_tot.Top + esver + esver), Width = 150, Font = new Font("Sans", 12.0F), Anchor = (AnchorStyles.None) };
            tx_mail = new TextBox { Location = new Point(lb_mail.Left + lb_mail.Width + 10, tx_tot.Top + esver + esver), Width = 300, Font = new Font("Sans", 12.0F), Anchor = (AnchorStyles.None) };
            tx_mail.Leave += tx_mail_leave;

            tx_mail.Text = grilla.Rows[fila].Cells["correoAdquiriente"].Value.ToString();    // mail
            pubmail = tx_mail.Text;
            //
            Button bt_impt = new Button();
            bt_impt = new Button { Text = "Graba/Imprime", Location = new Point(180, tx_mail.Top + esver + 7), Width = 100, Height = 30, BackColor = Color.GhostWhite, Anchor = (AnchorStyles.None) };
            bt_impt.Click += bt_click_doc;
            // ---- boton de impresion ----
            documento.Controls.Add(lb_fec);
            documento.Controls.Add(tx_fec);
            documento.Controls.Add(lb_doc);
            documento.Controls.Add(tx_doc);
            documento.Controls.Add(lb_cli);
            documento.Controls.Add(tx_cli);
            documento.Controls.Add(lb_ruc);
            documento.Controls.Add(tx_ruc);
            documento.Controls.Add(lb_dir);
            documento.Controls.Add(tx_dir);
            documento.Controls.Add(lb_dis);
            documento.Controls.Add(tx_dis);
            documento.Controls.Add(lb_pro);
            documento.Controls.Add(tx_pro);
            documento.Controls.Add(gd);
            documento.Controls.Add(lb_sub);
            documento.Controls.Add(tx_sub);
            documento.Controls.Add(lb_igv);
            documento.Controls.Add(tx_igv);
            documento.Controls.Add(lb_tot);
            documento.Controls.Add(tx_tot);
            documento.Controls.Add(lb_mail);
            documento.Controls.Add(tx_mail);
            documento.Controls.Add(bt_impt);
            //
            documento.ShowDialog();
        }
        public void tx_mail_leave(object sender, EventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            pubmail = textBox.Text.ToString();
        }
        public void bt_click_doc(object sender, EventArgs e)                       // click del botton del documento
        {
            if (string.IsNullOrEmpty(pubmail) || string.IsNullOrWhiteSpace(pubmail))
            {
                MessageBox.Show("Ingrese el correo electrónico del cliente", "Acción Obligatoria", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            else
            {
                string doc = grilla.Rows[idxgri].Cells["local"].Value.ToString() + grilla.Rows[idxgri].Cells["docvta"].Value.ToString() +
                    grilla.Rows[idxgri].Cells["servta"].Value.ToString() + grilla.Rows[idxgri].Cells["corrvta"].Value.ToString();
                // graba el correo en la tabla de doc.vta. de la interfas
                try
                {
                    MySqlConnection conn = new MySqlConnection(DB_CONN_STR);
                    conn.Open();
                    if (conn.State.ToString() == "Open")
                    {
                        string actua = "update fe_cabecera set mail=@mail where concat(local,docvta,servta,corrvta)=@doc";
                        MySqlCommand micon = new MySqlCommand(actua, conn);
                        micon.Parameters.AddWithValue("@mail", pubmail);
                        micon.Parameters.AddWithValue("@doc", doc);
                        micon.ExecuteNonQuery();
                    }
                    else
                    {
                        MessageBox.Show("No se puede actualizar el correo electrónico", "Error en conectarse al servidor");
                    }
                    conn.Close();
                    // graba/actualiza el campo email en la base y tabla del cliente (maestra de clientes)
                    string CONN_CLTE = "server=" + tx_serv.Text.Trim() + ";port=" + tx_port.Text.Trim() + ";uid=" + tx_usua.Text.Trim() +
                        ";pwd=" + tx_pass.Text.Trim() + ";database=" + tx_base.Text.Trim() + ";ConnectionLifeTime=" + ctl + ";";
                    MySqlConnection conc = new MySqlConnection(CONN_CLTE);
                    conc.Open();
                    if (conc.State.ToString() == "Open")
                    {
                        string actua = "update maclie set email=@mail where docclie=@tdc and numdoc=@idc";
                        MySqlCommand micon = new MySqlCommand(actua, conc);
                        micon.Parameters.AddWithValue("@mail", pubmail);
                        micon.Parameters.AddWithValue("@tdc", grilla.Rows[idxgri].Cells["doccli"].Value.ToString());
                        micon.Parameters.AddWithValue("@idc", grilla.Rows[idxgri].Cells["numdcli"].Value.ToString());
                        micon.ExecuteNonQuery();
                    }
                    else
                    {
                        MessageBox.Show("No se puede actualizar el correo electrónico", "Error en conectarse al cliente");
                    }
                    conc.Close();
                    // imprime el doc. de venta
                    if (imprime(idxgri) == true)                                // imprime y graba la marca impreso
                    {
                        if (marca_print(idxgri) == true)                        // marca las banderas "impreso"
                        {
                            FormCollection fc = Application.OpenForms;
                            foreach (Form frm in fc)
                            {
                                if (frm.Name == "documento") frm.Close();
                            }
                        }
                        else
                        {
                            MessageBox.Show("Ocurrio un error en la actualización de datos" + Environment.NewLine +
                                "por favor, vuelta a imprimir el documento", "Error en actualización", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            FormCollection fc = Application.OpenForms;
                            foreach (Form frm in fc)
                            {
                                if (frm.Name == "documento") frm.Close();
                            }
                        }
                        return;
                    }
                }
                catch (MySqlException ex)
                {
                    MessageBox.Show(ex.Message,"Error de conectividad");
                    Application.Exit();
                    return;
                }
            }
        }
		void Btn_Click(object sender, EventArgs e)                                // boton jala datos
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
            // llamamos al temporizador si el check_box auto esta marcado
            if (chk_auto.Checked == true)
            {
                if (temporizador != null)
                {
                    temporizador.Stop();
                    temporizador.Dispose();
                }
                InitTimer();
            }
            else
            {
                if (temporizador != null)
                {
                    temporizador.Stop();
                    temporizador.Dispose();
                }
            }
		}
        void Bt_print_Click(object sender, EventArgs e)                         // muestra los registros seleccionados a modo pre visualizacion
        {
            if (grilla.SelectedRows.Count == 1)
            {
                foreach (DataGridViewRow dgv in grilla.SelectedRows)
                {
                    if (dgv.Cells["impreso"].Value.ToString().Trim() == "" || dgv.Cells["impreso"].Value.ToString().Trim() == "N")
                    {
                        idxgri = dgv.Index;
                        documento(idxgri);
                    }
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
        private bool imprime(int idc)                                           // IMPRIME EN FISICO EL DOC. DE VENTA
		{
  	         bool retorna = false;
            /*
             PageSettings ps = new PageSettings();
             ps.Margins = new Margins(0, 0, 0, 0);                              // izquierda, derecha, arriba, abajo
             PaperSize ticket = new PaperSize("Custom", CentimeterToPixel(8), 850);              // custom, ancho = 200, largo = 112
             // la seleccion del formato de impresion debe ser en la configuracion - TICKET 
             */ 
             PrintDocument printDocument1 = new PrintDocument();
             printDocument1.PrintPage += new PrintPageEventHandler
                   (this.printDocument1_PrintPage);
            /*
             printDocument1.DefaultPageSettings.Margins = ps.Margins;
             printDocument1.DefaultPageSettings.PaperSize = ticket;
             */
            printDocument1.PrinterSettings.PrinterName = nom_imp;
            printDocument1.Print();
             // se envia a la impresora
             retorna = true;
  			return retorna;
  	   }
        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            //printDocument1.PrinterSettings.PrinterName = lnp.enlaces(nomform, "impresora", "default");    // "ticketera"
            Font lt_gra = new Font("Arial", 12);
            Font lt_tit = new Font("Lucida Console", 10);
            Font lt_peq = new Font("Arial", 7);
            Font lt_peb = new Font("Arial", 7, FontStyle.Bold);
            Font lt_min = new Font("Arial", 6);
            Font lt_med = new Font("Arial", 9);
            float ancho = 300.0F;    // ancho de la impresion
            int ctacop = 0;
            for (int i = 1; i <= copias; i++)
            {
                ctacop = ctacop + 1;
                int coli = 15;           // columna inicial
                float posi = 20;         // posicion x,y inicial
                int alfi = 15;          // alto de cada fila
                //
                float lt = (ancho - e.Graphics.MeasureString(nomclie, lt_gra).Width) / 2;
                PointF puntoF = new PointF(lt, posi);                        // serie y correlativo
                e.Graphics.DrawString(nomclie, lt_gra, Brushes.Black, puntoF, StringFormat.GenericTypographic);     // nombre comercial
                posi = posi + alfi + 5;
                lt = (ancho - e.Graphics.MeasureString(rasclie, lt_peq).Width) / 2;
                puntoF = new PointF(lt, posi);
                e.Graphics.DrawString(rasclie, lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);     // razon social
                posi = posi + alfi + 5;
                lt = (ancho - e.Graphics.MeasureString("RUC: " + rucclie, lt_tit).Width) / 2;
                puntoF = new PointF(lt, posi);
                e.Graphics.DrawString("RUC: " + rucclie, lt_tit, Brushes.Black, puntoF, StringFormat.GenericTypographic);     // ruc de emisor
                posi = posi + alfi + 2;
                lt = (ancho - e.Graphics.MeasureString(dirclie, lt_peq).Width) / 2;
                puntoF = new PointF(lt, posi);
                e.Graphics.DrawString(dirclie, lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);     // direccion emisor
                posi = posi + alfi;
                lt = (ancho - e.Graphics.MeasureString("Correo: " + corremi, lt_peq).Width) / 2;
                puntoF = new PointF(lt, posi);
                e.Graphics.DrawString("Correo: " + corremi, lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                string tipdo = grilla.Rows[idxgri].Cells["docvta"].Value.ToString().Trim();
                string serie = grilla.Rows[idxgri].Cells["servta"].Value.ToString();
                string corre = grilla.Rows[idxgri].Cells["corrvta"].Value.ToString();
                string nota = tipdo + "-" + serie + "-" + corre;
                string titdoc = "";
                if (tipdo == Cboleta) titdoc = "Boleta de Venta Electrónica";
                if (tipdo == Cfactura) titdoc = "Factura Electrónica";
                posi = posi + alfi + alfi;
                lt = coli + 10;
                lt = (ancho - e.Graphics.MeasureString(titdoc, lt_gra).Width) / 2;
                puntoF = new PointF(lt, posi);
                e.Graphics.DrawString(titdoc, lt_gra, Brushes.Black, puntoF, StringFormat.GenericTypographic);      // tipo de documento
                posi = posi + alfi + alfi;
                lt = (ancho - e.Graphics.MeasureString(serie + " - " + corre, lt_gra).Width) / 2;
                puntoF = new PointF(lt, posi);
                e.Graphics.DrawString(serie + " - " + corre, lt_gra, Brushes.Black, puntoF, StringFormat.GenericTypographic);   // serie y numero
                //posi = posi + alfi + alfi;
                //string locyus = tienda + " - " + operador;
                //puntoF = new PointF(coli, posi);
                //e.Graphics.DrawString(locyus, lt_med, Brushes.Black, puntoF, StringFormat.GenericTypographic);      // tienda y vendedor
                posi = posi + alfi + alfi;
                puntoF = new PointF(coli, posi);    // DateTime.Now.ToString()
                e.Graphics.DrawString("FECHA: " + grilla.Rows[idxgri].Cells["fechope"].Value.ToString().Substring(0,10) + " " +
                DateTime.Now.ToString("hh:mm tt", CultureInfo.InvariantCulture), lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic); // fecha y .... la hora emision no va
                posi = posi + alfi * 2;
                puntoF = new PointF(coli, posi);
                string glocli = "SEÑOR(ES): " + grilla.Rows[idxgri].Cells["clidv"].Value.ToString();
                e.Graphics.DrawString(glocli, lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);            // DNI/RUC cliente
                posi = posi + alfi + 5;
                puntoF = new PointF(coli, posi);
                e.Graphics.DrawString(grilla.Rows[idxgri].Cells["doccli"].Value.ToString() + ": " + grilla.Rows[idxgri].Cells["numdcli"].Value.ToString(),
                    lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                //puntoF = new PointF(coli + 140, posi);                                                                  // ubigeo
                //e.Graphics.DrawString("UBIGEO: " + grilla.Rows[idxgri].Cells["ubigeoAdq"].Value.ToString(),
                //    lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                posi = posi + alfi + 5;
                puntoF = new PointF(coli, posi);                                                                        // direccion
                //  me quede aca poniendo un rectangulo de 2 linea para la direccion
                SizeF sizDir2 = new SizeF(300, 30);
                RectangleF recdir = new RectangleF(puntoF, sizDir2);
                e.Graphics.DrawString(grilla.Rows[idxgri].Cells["direc1"].Value.ToString(), lt_peq, Brushes.Black, recdir, StringFormat.GenericTypographic);
                posi = posi + alfi + alfi; // *2;
                puntoF = new PointF(coli, posi);
                PointF puntoX = new PointF(coli + ancho - 20, posi);
                Pen raya = new Pen(Color.Black,1);
                e.Graphics.DrawLine(raya, puntoF, puntoX);                  // pinta la raya horizontal de separacion
                posi = posi + 7; // *2;
                // **************** detalle del documento ****************//
                puntoF = new PointF(coli - 2, posi);
                e.Graphics.DrawString("Cant", lt_peb, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                puntoF = new PointF(coli + 25, posi);
                e.Graphics.DrawString("Unidad", lt_peb, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                puntoF = new PointF(coli + 70, posi);
                e.Graphics.DrawString("Descripción", lt_peb, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                puntoF = new PointF(coli + 240, posi);
                e.Graphics.DrawString("Total", lt_peb, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                posi = posi + alfi; // *2;
                puntoF = new PointF(coli, posi);
                puntoX = new PointF(coli + ancho - 20, posi);
                e.Graphics.DrawLine(raya, puntoF, puntoX);                  // pinta la raya horizontal de separacion
                posi = posi + 7; // *2;
                //posi = posi + alfi; // *2;
                for (int f = 0; f < grillad.Rows.Count - 1; f++)
                {
                    DataGridViewRow row = grillad.Rows[f];
                    if (row.Cells["tipdv"].Value.ToString() == tipdo && row.Cells["servta"].Value.ToString() == serie && row.Cells["corvta"].Value.ToString() == corre)
                    {
                        puntoF = new PointF(coli + 10, posi);
                        e.Graphics.DrawString(row.Cells["cantidad"].Value.ToString().Trim(), lt_min, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                        puntoF = new PointF(coli + 30, posi);
                        e.Graphics.DrawString("NIU", lt_min, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                        puntoF = new PointF(coli + 60, posi);
                        SizeF sizD = new SizeF(170, 60);
                        RectangleF recdes = new RectangleF(puntoF, sizD);
                        e.Graphics.DrawString(row.Cells["descripcion"].Value.ToString(), lt_min, Brushes.Black, recdes, StringFormat.GenericTypographic);
                        //e.Graphics.DrawString(row.Cells["codprd"].Value.ToString().Trim(), lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                        puntoF = new PointF(coli + 240, posi);  // antes 255, lo baje para que salga el ".00" 06-03-2019
                        e.Graphics.DrawString(row.Cells["total"].Value.ToString().Trim(), lt_min, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                        //posi = posi + alfi + alfi;
                        //puntoF = new PointF(coli, posi);
                    }
                }
                posi = posi + alfi * 3 + 5;
                puntoF = new PointF(coli, posi);
                puntoX = new PointF(coli + ancho - 20, posi);
                raya.DashStyle = System.Drawing.Drawing2D.DashStyle.Dash;
                e.Graphics.DrawLine(raya, puntoF, puntoX);                  // pinta la raya horizontal de separacion
                // pie del documento ;
                if (tipdo == "BV")
                {
                    SizeF siz = new SizeF(70, 15);
                    StringFormat alder = new StringFormat(StringFormatFlags.DirectionRightToLeft);
                    posi = posi + alfi; // * 2;
                    puntoF = new PointF(coli, posi);
                    e.Graphics.DrawString("SUBTOTAL", lt_med, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                    puntoF = new PointF(coli + 170, posi);
                    e.Graphics.DrawString(grilla.Rows[idxgri].Cells["simboloMon"].Value.ToString(), lt_med, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                    puntoF = new PointF(coli + 190, posi);
                    RectangleF recst = new RectangleF(puntoF, siz);
                    e.Graphics.DrawString(grilla.Rows[idxgri].Cells["totValVtaNetGrav"].Value.ToString(), lt_med, Brushes.Black, recst, alder);
                    posi = posi + alfi + alfi;
                    puntoF = new PointF(coli, posi);    // grilla.Rows[idxgri].Cells["tasaigv"].Value.ToString()
                    e.Graphics.DrawString("IGV " + string.Format("{0:##}", grilla.Rows[idxgri].Cells["tasaigv"].Value) + " % ", lt_med, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                    puntoF = new PointF(coli + 170, posi);
                    e.Graphics.DrawString(grilla.Rows[idxgri].Cells["simboloMon"].Value.ToString(), lt_med, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                    puntoF = new PointF(coli + 190, posi);
                    RectangleF recgv = new RectangleF(puntoF, siz);
                    e.Graphics.DrawString(grilla.Rows[idxgri].Cells["totalIgv"].Value.ToString(), lt_med, Brushes.Black, recgv, alder);
                    posi = posi + alfi + alfi;
                    puntoF = new PointF(coli, posi);
                    e.Graphics.DrawString("IMPORTE TOTAL ", lt_med, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                    puntoF = new PointF(coli + 170, posi);
                    e.Graphics.DrawString(grilla.Rows[idxgri].Cells["simboloMon"].Value.ToString(), lt_med, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                    puntoF = new PointF(coli + 190, posi);
                    RectangleF recto = new RectangleF(puntoF, siz);
                    e.Graphics.DrawString(grilla.Rows[idxgri].Cells["totalVenta"].Value.ToString(), lt_med, Brushes.Black, recto, alder);
                }
                if (tipdo == "FT")
                {
                    SizeF siz = new SizeF(70, 15);
                    StringFormat alder = new StringFormat(StringFormatFlags.DirectionRightToLeft);
                    posi = posi + alfi; // * 2;
                    puntoF = new PointF(coli, posi);
                    e.Graphics.DrawString("SUBTOTAL", lt_med, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                    puntoF = new PointF(coli + 170, posi);
                    e.Graphics.DrawString(grilla.Rows[idxgri].Cells["simboloMon"].Value.ToString(), lt_med, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                    puntoF = new PointF(coli + 190, posi);
                    RectangleF recst = new RectangleF(puntoF, siz);
                    e.Graphics.DrawString(grilla.Rows[idxgri].Cells["totValVtaNetGrav"].Value.ToString(), lt_med, Brushes.Black, recst, alder);
                    posi = posi + alfi + alfi;
                    puntoF = new PointF(coli, posi);
                    e.Graphics.DrawString("IGV " + string.Format("{0:##}", grilla.Rows[idxgri].Cells["tasaigv"].Value) + " % ", lt_med, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                    puntoF = new PointF(coli + 170, posi);
                    e.Graphics.DrawString(grilla.Rows[idxgri].Cells["simboloMon"].Value.ToString(), lt_med, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                    puntoF = new PointF(coli + 190, posi);
                    RectangleF recgv = new RectangleF(puntoF, siz);
                    e.Graphics.DrawString(grilla.Rows[idxgri].Cells["totalIgv"].Value.ToString(), lt_med, Brushes.Black, recgv, alder);
                    posi = posi + alfi + alfi;
                    puntoF = new PointF(coli, posi);
                    e.Graphics.DrawString("IMPORTE TOTAL ", lt_med, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                    puntoF = new PointF(coli + 170, posi);
                    e.Graphics.DrawString(grilla.Rows[idxgri].Cells["simboloMon"].Value.ToString(), lt_med, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                    puntoF = new PointF(coli + 190, posi);
                    RectangleF recto = new RectangleF(puntoF, siz);
                    e.Graphics.DrawString(grilla.Rows[idxgri].Cells["totalVenta"].Value.ToString(), lt_med, Brushes.Black, recto, alder);
                }
                // monto en letras
                posi = posi + alfi * 2;
                puntoF = new PointF(coli, posi);
                e.Graphics.DrawString("SON: " + mel.Convertir(grilla.Rows[idxgri].Cells["totalVenta"].Value.ToString(),true) + " " +
                    equivalencias("MON","Soles","rid"), lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                // detracción 
                if (tipdo == "FT" && 
                    Convert.ToDecimal(grilla.Rows[idxgri].Cells["totalVenta"].Value.ToString()) >= Convert.ToDecimal(mondetr))
                {
                    StringFormat alcen = new StringFormat();
                    alcen.Alignment = StringAlignment.Center;
                    SizeF siz = new SizeF(70, 15);
                    string glosa = glodetr + " " + ctadetr;
                    float anglo = e.Graphics.MeasureString(glosa, lt_med).Width;
                    if (anglo - ancho > 0) siz = new SizeF(ancho, 30);
                    posi = posi + alfi * 2;
                    puntoF = new PointF(coli, posi);
                    RectangleF detra = new RectangleF(puntoF, siz);
                    e.Graphics.DrawString(glosa, lt_med, Brushes.Black, detra, alcen);
                    //pordetr = dr.GetString("detra").Trim();                 // porcentaje detraccion
                }
                string separ = "|";
                string codigo = rucclie + separ + grilla.Rows[idxgri].Cells["tipoDocumento"].Value.ToString() + separ +
                    grilla.Rows[idxgri].Cells["servta"].Value.ToString() + "-" + grilla.Rows[idxgri].Cells["corrvta"].Value.ToString() + separ +
                    grilla.Rows[idxgri].Cells["totalIgv"].Value.ToString() + separ + grilla.Rows[idxgri].Cells["totalVenta"].Value.ToString() + separ +
                    string.Format("{0:yyyy-MM-dd}",grilla.Rows[idxgri].Cells["fechope"].Value) + separ + grilla.Rows[idxgri].Cells["tipoDocAdq"].Value.ToString() + separ +
                    grilla.Rows[idxgri].Cells["numdcli"].Value.ToString();
                //
                var rnd = Path.GetRandomFileName();
                var otro = Path.GetFileNameWithoutExtension(rnd);
                otro = otro + ".png";
                //
                var qrEncoder = new QrEncoder(ErrorCorrectionLevel.H);
                var qrCode = qrEncoder.Encode(codigo);
                var renderer = new GraphicsRenderer(new FixedModuleSize(5, QuietZoneModules.Two), Brushes.Black, Brushes.White);
                using (var stream = new FileStream(otro, FileMode.Create))
                    renderer.WriteToStream(qrCode.Matrix, ImageFormat.Png, stream);    // "qrcode.png"
                Bitmap png = new Bitmap(otro);  // "qrcode.png"
                //var bmp = new Bitmap("qrcode.png");
                //Bitmap png = (Bitmap)bmp.Clone();
                posi = posi + alfi + alfi;
                puntoF = new PointF(coli + 50, posi);
                SizeF cuadro = new SizeF(CentimeterToPixel(4),CentimeterToPixel(4));    // 5x5 cm
                RectangleF rec = new RectangleF(puntoF,cuadro);
                e.Graphics.DrawImage(png, rec);
                posi = posi + alfi + CentimeterToPixel(4);
                /* SEGUN PERUSECURE ESTO YA NO VA ... 22-02-19
                // leyenda 1
                posi = posi + alfi + CentimeterToPixel(4);
                puntoF = new PointF(coli, posi);                // Autorizado mediante resolucion ...
                e.Graphics.DrawString(leyen1, lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                // leyenda 2
                posi = posi + alfi;
                puntoF = new PointF(coli, posi);                // 034-005-0010474/SUNAT
                e.Graphics.DrawString(nuausu, lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                */
                // leyenda 3
                //posi = posi + alfi + alfi;
                puntoF = new PointF(coli, posi);                // "Representación impresa de la boleta  de venta electronica"
                SizeF leyen = new SizeF(ancho - 20, alfi * 3);
                RectangleF recley3 = new RectangleF(puntoF, leyen);
                e.Graphics.DrawString(leyen3, lt_peq, Brushes.Black, recley3, StringFormat.GenericTypographic);
                // despedida
                StringFormat sf = new StringFormat();
                sf.Alignment = StringAlignment.Center;
                posi = posi + alfi * 3;
                //lt = (ancho - e.Graphics.MeasureString(desped, lt_gra).Width) / 2;
                puntoF = new PointF(coli, posi);
                leyen = new SizeF(ancho - 20, alfi * 3);
                RectangleF recdesp = new RectangleF(puntoF, leyen);
                e.Graphics.DrawString(desped, lt_peq, Brushes.Black, recdesp, sf);
                // leyenda 4
                posi = posi + alfi * 2;
                puntoF = new PointF(coli, posi);
                leyen = new SizeF(ancho - 20, alfi * 3);
                RectangleF recdesp2 = new RectangleF(puntoF, leyen);
                e.Graphics.DrawString(despe2, lt_peq, Brushes.Black, recdesp2, sf);
                // leyenda 5
                posi = posi + alfi * 2;
                puntoF = new PointF(coli, posi);
                leyen = new SizeF(ancho - 20, alfi * 3);
                RectangleF recley5 = new RectangleF(puntoF, leyen);
                e.Graphics.DrawString(provee, lt_peq, Brushes.Black, recley5, sf);
                // 
                if (copias == ctacop)
                {
                    e.HasMorePages = false;
                }
                else { e.HasMorePages = true; }
            }
        }
        private bool marca_print(int indice)                                    // actualiza las banderas de "impreso"
        {
            bool retorna = false;
            try
            {
                MySqlConnection conn = new MySqlConnection(DB_CONN_STR);
                conn.Open();
                if (conn.State.ToString() == "Open")            // base de fact. electronica
                {
                    string actua = "update fe_cabecera set impreso='S' where local=@loc and docvta=@doc and servta=@ser and corrvta=@cor";
                    MySqlCommand micon = new MySqlCommand(actua, conn);
                    micon.Parameters.AddWithValue("@loc", grilla.Rows[indice].Cells["local"].Value.ToString());
                    micon.Parameters.AddWithValue("@doc", grilla.Rows[indice].Cells["docvta"].Value.ToString());
                    micon.Parameters.AddWithValue("@ser", grilla.Rows[indice].Cells["servta"].Value.ToString());
                    micon.Parameters.AddWithValue("@cor", grilla.Rows[indice].Cells["corrvta"].Value.ToString());
                    micon.ExecuteNonQuery();
                    retorna = true;
                }
                else
                {
                    MessageBox.Show("Error en abrir la conexión", "Fallo de conexión");
                }
                conn.Close();
                // jalamos los datos
                string CONN_CLTE = "server=" + tx_serv.Text.Trim() + ";port=" + tx_port.Text.Trim() + ";uid=" + tx_usua.Text.Trim() +
                    ";pwd=" + tx_pass.Text.Trim() + ";database=" + tx_base.Text.Trim() + ";ConnectionLifeTime=" + ctl + ";";
                MySqlConnection conc = new MySqlConnection(CONN_CLTE);
                conc.Open();
                if (conc.State.ToString() == "Open")        // base y tabla del cliente, cabecera de ventas
                {
                    string actua = "update docvtas set impreso='S' where numcaja=@loc and docvta=@doc and servta=@ser and corrvta=@cor";
                    MySqlCommand micon = new MySqlCommand(actua, conc);
                    micon.Parameters.AddWithValue("@loc", grilla.Rows[indice].Cells["local"].Value.ToString());
                    micon.Parameters.AddWithValue("@doc", grilla.Rows[indice].Cells["docvta"].Value.ToString());
                    micon.Parameters.AddWithValue("@ser", grilla.Rows[indice].Cells["servta"].Value.ToString());
                    micon.Parameters.AddWithValue("@cor", grilla.Rows[indice].Cells["corrvta"].Value.ToString());
                    micon.ExecuteNonQuery();
                    retorna = true;
                }
                else
                {
                    MessageBox.Show("Error en abrir la conexión", "Fallo de conexión cliente");
                }
                conc.Close();
                // *************** actualizamos la grilla ******************
                //object obj = (object)grilla.Rows[idxgri].DataBoundItem;
                //object obj = (object)grilla.Rows[idxgri].Cells["ds"].ColumnIndex;
                //DataGridCell val = (DataGridCell)grilla.Rows[idxgri].Cells["impreso"].Value;
                grilla.Rows[idxgri].Cells["impreso"].Value = "S";
                
            }
            catch (MySqlException ex)
            {
                MessageBox.Show(ex.Message, "Error en la conexión");
                Application.Exit();
            }
            return retorna;
        }
        int CentimeterToPixel(double Centimeter)                                // utilitario para calcular pixceles desde centimetros
        {
            double pixel = -1;
            using (Graphics g = this.CreateGraphics())
            {
                pixel = Centimeter * g.DpiY / 2.54d;
            }
            return (int)pixel;
        }
        private string equivalencias(string tabella,string codigo,string cual)  // retorna el codigo sunat del codigo del emisor
        {
            string retorna = "";
            // busco en datatable sunat el codigo ingresado y retorno el codigo sunat
            for (int i = 0; i < sunat.Rows.Count; i++)
            {
                DataRow row = sunat.Rows[i];
                if(row["idtabella"].ToString() == tabella && row["idcodice"].ToString().ToUpper() == codigo.ToString().ToUpper())
                {
                    if (cual == "sunat") retorna = row["codsunat"].ToString();
                    if (cual == "cod2") retorna = row["codice2"].ToString();
                    if (cual == "rid") retorna = row["descrizionerid"].ToString().ToUpper();
                    if (cual == "sunat") retorna = row["codsunat"].ToString();
                }
            }
            if (retorna == "")
            {
                MessageBox.Show("No se encuentra el equivalente al tabella [" + tabella + "]" + Environment.NewLine +
                "código [" + codigo + "] en la tabla de descripciones", "Error en configuración", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return retorna;
        }
        void Bt_anu_Click(object sender, EventArgs e)                           // anula (da de baja) los registros seleccionados
        {
            if (grilla.SelectedRows.Count > 0)
            {
                foreach (DataGridViewRow dgv in grilla.SelectedRows)
                {
                        // esto esta por verce en viza
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
            //creador();
        }
        private void InitTimer()
        {
            if (chk_auto.Checked == true)
            {
                temporizador = new System.Windows.Forms.Timer();
                temporizador.Interval = int.Parse(tx_time.Text) * 1000;    // 10000;        // 10 segundos
                temporizador.Tick += new EventHandler(timer_tick);
                temporizador.Start();
            }
        }
        private void timer_tick(object sender, EventArgs e)
        {
            //Btn_Click(null, null);
            trabaja();
        }
        private bool valexis(string caode,string tipo, string serie, string corre)           // valida existencia del doc de venta en la tabla propia
        {                   // caode = ca o de, cabeza o detalle
            bool retorna = false;
            try
            {
                MySqlConnection conn = new MySqlConnection(DB_CONN_STR);
                conn.Open();
                string consulta = "";
                if (caode == "fe_cabecera")
                {
                    consulta = "select count(*) from fe_cabecera where docvta=@tipo and servta=@serie and corrvta=@corre";
                }
                if (caode == "fe_detalle")
                {
                    consulta = "select count(*) from fe_detalle where tipdv=@tipo and servta=@serie and corvta=@corre";
                }
                MySqlCommand micon = new MySqlCommand(consulta, conn);
                micon.Parameters.AddWithValue("@tipo", tipo);
                micon.Parameters.AddWithValue("@serie", serie);
                micon.Parameters.AddWithValue("@corre", corre);
                MySqlDataReader dr = micon.ExecuteReader();
                if (dr.Read())
                {
                    if ( dr.GetString(0) == "1") retorna = true;
                    else retorna = false;
                    //dr.Close();
                }
                dr.Close();
                conn.Close();
            }
            catch (MySqlException ex)
            {
                MessageBox.Show(ex.Message, "error de conectividad");
                Application.Exit();
            }
            return retorna;
        }
        #region xmlzipcdr
        private void creador()                                          // creador de xml y zip por cada registro
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
            //string CONN_CLTE = "server=" + tx_serv.Text.Trim() + ";port=" + tx_port.Text.Trim() + ";uid=" + tx_usua.Text.Trim() + 
            //    ";pwd=" + tx_pass.Text.Trim() + ";database=" + tx_base.Text.Trim() + ";ConnectionLifeTime=" + ctl + ";";
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
        #endregion
    }
}

