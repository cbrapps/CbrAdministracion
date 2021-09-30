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
using System.Diagnostics;
using System.Drawing.Text;
using System.Security.Permissions;
using WECPOFLogic;
using System.Data.SqlClient;
using CBR_ADMIN.Properties;
using System.Net.Mail;
using System.Net.Mime;
using aejw.Network;
using System.Threading;
using CBR_ADMIN.Sistema;
using CBR_ADMIN.IT;

namespace CBR_ADMIN
{
    public partial class windowsNavegador : UserControl
    {
        private static windowsNavegador _instance;
        public static windowsNavegador Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new windowsNavegador();
                return _instance;
            }
        }
        private bool isFile = false;
        private string filePath = "G:/";
        private string seleccionNombreArchvio = "";
        private object fbd;
        public string usuario,apellido,emailusuario,departamento,formato;
       public  string valor="", valor2 = "";
        public string contadorcar = "";
        public string Nombre1 = "";
        Process p = new Process();
        public string anorastreo, proyectorastreo,carpetarastreo,codifica,nomproyec1;
        string carptape,proyectoras;
        String Email1, Email2, Email3, Email4, Email5, Email6, Email7, Email8, Email9, Email10;
        String Departamento;
        int numerador; string Nombre; int suma, k;
        int indicador = 0;
       
        string Año2="" ;
        string tipo3="";
        string Nombre2="";
   
        public windowsNavegador()
        {
            InitializeComponent();
        }


        public class TableLayoutPanelNoFlicker : TableLayoutPanel
        {
            public TableLayoutPanelNoFlicker()
            {
               this.DoubleBuffered = true;
                SetStyle(ControlStyles.OptimizedDoubleBuffer |
         ControlStyles.AllPaintingInWmPaint, true);

            }
        }
        private void windowsNavegador_Load(object sender, EventArgs e)
        {
            string url = "//192.168.1.101/Aplicativos/CBR Sistema global/network.jpg";
            this.DoubleBuffered = true;

            webBrowser1.Url = new Uri(url);
            navegador.Text = url;
            this.carpetasTableAdapter.Fill(carpetas._Carpetas);
            this.folio_ProyectosTableAdapter.Fill(folio_Proyectos._Folio_Proyectos);
            this.folio_ProspectosTableAdapter.Fill(cBR_Prospectos.Folio_Prospectos);
            this.login_DepartamentosTableAdapter.Fill(cBR_Login_Dep.Login_Departamentos);
            this.term_inox_adminTableAdapter.Fill(cb_term_inox.Term_inox_admin);
            this.term_af_com_adminTableAdapter.Fill(term_af._Term_af_com_admin);
          
            indus_po.Visible = false;
            Cb_af_com.Visible = false;
            cb_inox.Visible = false;
            Atencionaclientes.Visible = false;
            Calidad.Visible = false;
            Compras.Visible = false;
            OperacionyManteniento.Visible = false;
            Proyectos.Visible = false;
            Ventas.Visible = false;
            button2.Visible = false;
            button3.Visible = false;
            button4.Visible = false;
            button5.Visible = false;
            button6.Visible = false;
            button7.Visible = false;
            button8.Visible = false;
            button9.Visible = false;
            button10.Visible = false;
            button11.Visible = false;
            button12.Visible = false;
            button13.Visible = false;
            button14.Visible = false;
            button15.Visible = false;
            button16.Visible = false;
            button17.Visible = false;
            button18.Visible = false;

            SqlConnection conexion = new SqlConnection(ObtenerCadena());
    

            conexion.Open();
            ////////////////////////////////////////////////////////////////////////////////////






            ////////////////////////////////////SE EJECUTA EL COMANDO HACIA LA VISTA Y SE HACE LA VALIDACION//////////////////////////////////////
            SqlCommand cmd = new SqlCommand(
                                            "select " +
                                            "Count('Consecutivo') " +



                                            "from [Folio_Proyectos] "

                                            , conexion);


            SqlDataAdapter sda = new SqlDataAdapter(cmd);
            sda.SelectCommand.CommandTimeout = 36000;
            DataTable dt = new DataTable();
            sda.Fill(dt);



            conexion.Dispose();
            if (dt.Rows.Count > 0)
            {
                Nombre = dt.Rows[0][0].ToString();

              //  Extractor();
            }

            else { }

            if (departamento == "Ventas")
            {
                Ventas.Visible = true;
                button8.Visible = true;
            }
            else if (departamento == "Atencion a Clientes") { Atencionaclientes.Visible = true;
           
                Calidad.Visible = false;
                Compras.Visible = false;
                OperacionyManteniento.Visible = false;
                Proyectos.Visible = false;
                Ventas.Visible = false;
                button9.Visible = true;

            }
            else if (departamento == "Infraestructura")
            {
                button10.Visible = true;

            }
            else if (departamento == "Produccion")
            {
                button16.Visible = true;

            }
            else if (departamento == "Construccion e Instalaciones")
            {
                button15.Visible = true;
                button14.Visible = true;

            }
            else if (departamento == "SuperAdmin" || departamento == "Direccion" || departamento == "Asistente de direccion")
            {
                button2.Visible = true;
                button3.Visible = true;
                button4.Visible = true;
                button5.Visible = true;
                button6.Visible = true;
                button7.Visible = true;
                button8.Visible = true;
                button9.Visible = true;
                button10.Visible = true;
                button11.Visible = true;
                button12.Visible = true;
                button13.Visible = true;
                button14.Visible = true;
                button15.Visible = true;
                button16.Visible = true;
                button17.Visible = true;
                button18.Visible = true;

            }

            else if (departamento == "Compras") { Compras.Visible = true;
                Atencionaclientes.Visible = false;
                Calidad.Visible = false;
                button18.Visible = true;

                OperacionyManteniento.Visible = false;
                Proyectos.Visible = false;
                Ventas.Visible = false;
            }
            else if (departamento == "Proyectos") { Proyectos.Visible = true;
                Atencionaclientes.Visible = false;
                Calidad.Visible = false;
                Compras.Visible = false;
                OperacionyManteniento.Visible = false;
                button17.Visible = true;
                Ventas.Visible = false;
            }
            else if (departamento == "Operacion y Mantenimiento") { 
                OperacionyManteniento.Visible = true;
                Atencionaclientes.Visible = false;
                Calidad.Visible = false;
                Compras.Visible = false;
                button13.Visible = true;
                Proyectos.Visible = false;
                Ventas.Visible = false;
            
            }
            else if (departamento == "Proyectos" || departamento == "Produccion") { Calidad.Visible = true;
                Atencionaclientes.Visible = false;
             
                Compras.Visible = false;
                OperacionyManteniento.Visible = false;
                Proyectos.Visible = false;
                Ventas.Visible = false;
            }

            if (departamento == "Ventas" || departamento == "Potabiliza" || departamento == "SuperAdmin")
            {
                label16.Visible = true;
                ruta.Visible = true;
                pictureBox9.Visible = true;
            }
            else
            {
                label16.Visible = false;
                ruta.Visible = false;
                pictureBox9.Visible = false;
            }

            string strCmdText;
            //strCmdText = "/C attrib +h +s G:\\SGC-PROYECTOS-CBR ";
            //Process cop = System.Diagnostics.Process.Start("CMD.exe", strCmdText);
            //cop.WaitForExit();
            //strCmdText = "/C attrib +h +s G:\\SGC";
            //Process cop1 = System.Diagnostics.Process.Start("CMD.exe", strCmdText);
            //cop1.WaitForExit();
            //strCmdText = "/C attrib +h +s G:\\Sistema ";
            //Process cop2 = System.Diagnostics.Process.Start("CMD.exe", strCmdText);
            //cop2.WaitForExit();
            Screen screen = Screen.PrimaryScreen;

            int Height = screen.Bounds.Width;

            int Width = screen.Bounds.Height;
            string het = Convert.ToString(Height);
            string wit = Convert.ToString(Width);
            //label17.Text = het + "  " + wit;
            if (Height < 1536)
            {
                webBrowser1.Size = new Size(1400, 302);
                this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));

            }
            else { }


        }
        private void MostrarCuadroInicio()
        {
            Loading Cargando = new Loading();

            Cargando.ShowDialog();

        }
        public void Extractor()
        {

            SqlConnection conexion = new SqlConnection(ObtenerCadena());

            numerador = Int32.Parse(Nombre);

            suma = 0;
            conexion.Dispose();
            for (k = 1; k <= numerador; k++)
            {
                suma = suma + k;
                Consulta();

                // richTextBox1.Text = suma.ToString();








            }

            ////////////////////////////////////SE EJECUTA EL COMANDO HACIA LA VISTA Y SE HACE LA VALIDACION//////////////////////////////////////

        }
        public void Consulta()
        {
            SqlConnection conexion = new SqlConnection(ObtenerCadena());
            conexion.Open();
            SqlCommand cmd = new SqlCommand(
                                   "select top " + "(" + suma + ")" +
                                   "[Ano], " +
                                   "[tipo2], " +
                                   "[Nombre] " +

                                   "from [Folio_Proyectos] "

                                   , conexion);
            SqlDataAdapter sda = new SqlDataAdapter(cmd);
            sda.SelectCommand.CommandTimeout = 36000;
            DataTable dt = new DataTable();
            sda.Fill(dt);
            if (dt.Rows.Count > 0)
            {
                string Año = dt.Rows[k - 1][0].ToString();
                string tipo = dt.Rows[k - 1][1].ToString();
                string Nombre = dt.Rows[k - 1][2].ToString();
                Thread.Sleep(50);
                // richTextBox1.Text = richTextBox1.Text + "\n" + "G:\\SGC-PROYECTOS-CBR\\SGC\\" + Año + "\\" + tipo + "\\" + Nombre;
                string strCmdText;
                strCmdText = "/C attrib -h -s "+"\""+"G:\\SGC-PROYECTOS-CBR\\SGC\\" + Año + "\\" + tipo + "\\" + Nombre+"\\"+"*" +"\"" + "  /S"+"  /D";
                Process cop = System.Diagnostics.Process.Start("CMD.exe", strCmdText);
                cop.WaitForExit();

                conexion.Dispose();
            }
            else { }
        }
        private void pictureBox3_Click(object sender, EventArgs e)
        {
            webBrowser1.Refresh();




        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            rastreadorproyecto();
            string url1;
            url1 = "file:///G:/SGC-PROYECTOS-CBR/SGC/" + Año2 + "/" + tipo3 + "/" + valor;
            string url2 = webBrowser1.Url.ToString();
            if (Proyecto.Text == "")
            {
                if (webBrowser1.CanGoBack)
                    webBrowser1.GoBack();

                navegador.Text = webBrowser1.Url.ToString();
            }
          
            else if (url1 != url2)
            {
                if (webBrowser1.CanGoBack)
                    webBrowser1.GoBack();

                navegador.Text = webBrowser1.Url.ToString();
            }
            else
            {

              
                DialogResult Resultado;
                Resultado = MessageBox.Show("Está seguro que desea terminar la sesion en el proyecto: " + Proyecto.Text, "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    Loading apertura = new Loading();
                    apertura.Show();
                    apertura.BringToFront();
                    apertura.WindowState = FormWindowState.Normal;
                    rastreadorproyecto();
                    this.registro_AccesosTableAdapter.elimina(Proyecto.Text);
                 

                 
                    webBrowser1.GoBack();
                  



                    string strCmdText;
                     strCmdText = "/C attrib +h +s  " +"\""+"G:\\SGC-PROYECTOS-CBR\\SGC\\" + Año2 + "\\" + tipo3 + "\\" +valor + "\\*"+"\" "+" /S /D";
                    Nombre1 = "";
                    valor = "";
                    Process cm = new Process();
                    cm.StartInfo.FileName = "cmd.exe";
                    cm.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    cm.StartInfo.Arguments = strCmdText;
                    cm.Start();
                    cm.WaitForExit();
                    apertura.Hide();
                    MessageBox.Show("Sesion cerrada correctamente");
                  
                    this.folio_ProyectosTableAdapter.Consulta(folio_Proyectos._Folio_Proyectos,"");
                    if (webBrowser1.CanGoBack)
                        webBrowser1.GoBack();

                    navegador.Text = webBrowser1.Url.ToString();
                  
                }
                else if (Resultado == DialogResult.No)
                {


                    MessageBox.Show("Operacion Cancelada");
                }
            }
            
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            if (webBrowser1.CanGoForward)
                webBrowser1.GoForward();


            navegador.Text = webBrowser1.Url.ToString();
        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
          

        }
                                                     
        private void webBrowser1_Navigating(object sender, WebBrowserNavigatingEventArgs e)                                                                                                                                                                                                                                                                        
                  {
        }

        private void webBrowser1_Navigated(object sender, WebBrowserNavigatedEventArgs e)
        {          

               string tipo1="";                                  
                                                                                                                    
            navegador.Text = webBrowser1.Url.ToString();                            
            char [] nave = navegador.Text.ToCharArray();
                 String proyecnave=navegador.Text.ToString();
             
                int f = nave.Count(); //6
            string str = new string(nave);
            if (f >= 52 && f < 53)/// inox
            { indicador = 1; valor2 = "inox"; }
            if (f >= 56 && f <= 57)/// compacta
            { indicador = 3; valor2 = "compacta"; }

            if (f >= 58 && f <= 59)/// altoflujo
            { indicador = 4; valor2 = "altoflujo"; }

            if (f >= 69 && f <= 70)/// indus pota
            { indicador = 2; valor2 = "Industrial"; }
            if (f >= 79 && f <= 80 && valor2== "Industrial")/// indus pota propuesta
            { indicador = 2; valor2 = "Industrial"; }

            ////////////////////////////////////////////////////inox////////////////////////
            if (indicador==1 && f >= 60) {
                f = f - 53;
                valor = proyecnave.Substring(53, f);
                indicador = 0;
            
            }
        

            ////////////////////////////////////////////////////indus pota propuesta////////////////////////
            if (indicador == 2 && f > 69)
            {
                f = f - 70;
                valor = proyecnave.Substring(70, f);
                indicador = 0;
         
            }
       

            ////////////////////////////////////////////////////inox////////////////////////
            if (indicador == 3 && f >= 60)
            {
                f = f - 57;
                valor = proyecnave.Substring(57, f);
                indicador = 0;
               
            }
                                                                                                                                                                     
            ////////////////////////////////////////////////////inox////////////////////////
            if (indicador == 4 && f >= 60)
            {
                f = f - 59;
                valor = proyecnave.Substring(59, f);
                indicador = 0;
         
            }

           // this.folio_ProyectosTableAdapter.(folio_Proyectos._Folio_Proyectos);
            SqlConnection conexion1 = new SqlConnection(ObtenerCadena());
            conexion1.Open();
            SqlCommand cmd1 = new SqlCommand(
                                   "select [Nombre2]" +
                                

                                   "from [Folio_Proyectos]  where Nombre=@nombre"

                                   , conexion1);
            cmd1.Parameters.AddWithValue("nombre", valor);
            SqlDataAdapter sda1 = new SqlDataAdapter(cmd1);
            sda1.SelectCommand.CommandTimeout = 36000;
            DataTable dt1 = new DataTable();
            sda1.Fill(dt1);
            if (dt1.Rows.Count > 0)
            {
             
                 Nombre1 = dt1.Rows[0][0].ToString();
             
            }
            else
            { }


            SqlConnection conexion = new SqlConnection(ObtenerCadena());
            conexion.Open();
            SqlCommand cmd = new SqlCommand(
                                   "select [Proyecto]" +
                                   "[ano], " +
                                   "[tipo], " +
                                   "[Usuario] " +

                                   "from [Registro_Accesos]  where Proyecto=@proyecto"

                                   , conexion);
            cmd.Parameters.AddWithValue("proyecto", Nombre1);
           
            SqlDataAdapter sda = new SqlDataAdapter(cmd);
            sda.SelectCommand.CommandTimeout = 36000;
            DataTable dt = new DataTable();
            sda.Fill(dt);
           
            if (dt.Rows.Count > 0 && Nombre1 !="")
            {
          
                string Año = dt.Rows[0][0].ToString();
                string tipo = dt.Rows[0][1].ToString();
                string Nombre = dt.Rows[0][2].ToString();
                if (Nombre != usuario + " " + apellido)
                {
                    DialogResult Resultado;
                    Resultado = MessageBox.Show("El proyecto esta abierto por el usuario:" + Nombre + "\nDesea forzar el ingreso al proyecto?   ", "Confirmación", MessageBoxButtons.YesNo);
                    if (Resultado == DialogResult.Yes)
                    {
                        this.registro_AccesosTableAdapter.elimina(Proyecto.Text);
                        Nombre1 = Proyecto.Text;
                        this.folio_ProyectosTableAdapter.Consulta(folio_Proyectos._Folio_Proyectos, Nombre1);
                        valor2 = tipo2.Text;

                        string url = "G:/SGC-PROYECTOS-CBR/SGC/" + Ano.Text + "/" + tipo2.Text + "/" + Proyecto.SelectedValue.ToString();
                        webBrowser1.Url = new Uri(url);
                        navegador.Text = url;
                        valor = Proyecto.SelectedValue.ToString();
                        contadortop();

                        this.registro_Accesos_monitorTableAdapter.ingresamonitor(Proyecto.Text, Ano.Text, tipo2.Text, usuario + " " + apellido);
                        webBrowser1.Refresh();

                    }
                    else if (Resultado == DialogResult.No)
                    {
                        webBrowser1.GoBack();

                        MessageBox.Show("Operacion Cancelada");
                    }


                }
                else { this.folio_ProyectosTableAdapter.Consulta(folio_Proyectos._Folio_Proyectos,Nombre1); }
            }
            else
            {
               
                if (f == 65 || f == 32 && valor2 == "" || f == 37 && valor2 == "" || f == 58 || f == 56 && valor2 == "" || f == 69 || f == 52 ) { }
                else
                {
                    this.folio_ProyectosTableAdapter.Consulta(folio_Proyectos._Folio_Proyectos, Nombre1);
                    if (valor == "") {  }
                    else
                    {
                       
                        contadortop();
                        this.registro_AccesosTableAdapter.Visualizar(Proyecto.Text, Ano.Text, tipo2.Text, usuario + " " + apellido);
                        this.registro_Accesos_monitorTableAdapter.ingresamonitor(Proyecto.Text, Ano.Text, tipo2.Text, usuario + " " + apellido);
                  
                        // this.folio_ProyectosTableAdapter.Fill(folio_Proyectos._Folio_Proyectos);
                        webBrowser1.Refresh(); 
                        MessageBox.Show("Inicio de Sesion correcto"); 
                    }

                }
                conexion.Dispose();
            }


            //  webBrowser1.GoBack();
           
        }

     
   
        private void pictureBox7_Click(object sender, EventArgs e)
        {
            if (Depa.Text != "" & carpeta.Text != "" & proye_solicita.Text != "")
            {

                DialogResult Resultado;
                Resultado = MessageBox.Show("Desea solicitar actualizacion de estado :\nCarpeta:" + carpeta.Text + "\nProyecto: " + "" + proye_solicita.Text, "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    if (Depa.Text == "") { MessageBox.Show("No ha ingresado el departamento a notificar"); }
                    if (Depa.Text != "")
                    {
                        enviamasivo2(Depa.Text);
                    }

                }
                else if (Resultado == DialogResult.No)
                {


                    MessageBox.Show("Operacion Cancelada");
                }
            }
            else if (Depa.Text == "" & carpeta.Text == "" & proye_solicita.Text == "")
            {
                MessageBox.Show("Faltan campos por llenar");
            }
            else if (Depa.Text == "" )
            {
                MessageBox.Show("Faltan campos por llenar");
            }

            else if ( carpeta.Text == "" )
            {
                MessageBox.Show("Faltan campos por llenar");
            }
            else if ( proye_solicita.Text == "")
            {
                MessageBox.Show("Faltan campos por llenar");
            }
        }

        private void pictureBox3_Click_1(object sender, EventArgs e)
        {

            SqlConnection conexion = new SqlConnection(ObtenerCadena());
            conexion.Open();
            SqlCommand cmd = new SqlCommand(
                                   "select [Proyecto]" +
                                   "[ano], " +
                                   "[tipo], " +
                                   "[Usuario] " +

                                   "from [Registro_Accesos]  where Proyecto=@proyecto"

                                   , conexion);
            cmd.Parameters.AddWithValue("proyecto", Proyecto.Text);
            SqlDataAdapter sda = new SqlDataAdapter(cmd);
            sda.SelectCommand.CommandTimeout = 36000;
            DataTable dt = new DataTable();
            sda.Fill(dt);
            if (dt.Rows.Count > 0)
            {
                string Año = dt.Rows[0][0].ToString();
                string tipo = dt.Rows[0][1].ToString();
                string Nombre = dt.Rows[0][2].ToString();
                if (Nombre != usuario + " " + apellido)
                {

                    DialogResult Resultado;
                    Resultado = MessageBox.Show("El proyecto esta abierto por el usuario:" + Nombre + "\nDesea forzar el ingreso al proyecto?   ", "Confirmación", MessageBoxButtons.YesNo);
                    if (Resultado == DialogResult.Yes)
                    {
                        this.registro_AccesosTableAdapter.elimina(Proyecto.Text);
                        Nombre1 = Proyecto.Text;
                        this.folio_ProyectosTableAdapter.Consulta(folio_Proyectos._Folio_Proyectos, Nombre1);
                        valor2 = tipo2.Text;

                        string url = "G:/SGC-PROYECTOS-CBR/SGC/" + Ano.Text + "/" + tipo2.Text + "/" + Proyecto.SelectedValue.ToString();
                        webBrowser1.Url = new Uri(url);
                        navegador.Text = url;
                        valor = Proyecto.SelectedValue.ToString();
                        contadortop();

                        this.registro_Accesos_monitorTableAdapter.ingresamonitor(Proyecto.Text, Ano.Text, tipo2.Text, usuario + " " + apellido);
                        webBrowser1.Refresh();

                    }
                    else if (Resultado == DialogResult.No)
                    {
                        webBrowser1.GoBack();

                        MessageBox.Show("Operacion Cancelada");
                    }
                }
                
                else {
                    Nombre1 = Proyecto.Text;
                    this.folio_ProyectosTableAdapter.Consulta(folio_Proyectos._Folio_Proyectos, Nombre1);
                    valor2 = tipo2.Text;
                 
                    string url = "G:/SGC-PROYECTOS-CBR/SGC/" + Ano.Text + "/" + tipo2.Text + "/" + Proyecto.SelectedValue.ToString();
                    webBrowser1.Url = new Uri(url);
                    navegador.Text = url;
                    valor = Proyecto.SelectedValue.ToString();
                    contadortop();
               
                    this.registro_Accesos_monitorTableAdapter.ingresamonitor(Proyecto.Text, Ano.Text, tipo2.Text, usuario + " " + apellido);
                    webBrowser1.Refresh();
                
                }
            }
            else
            {
                if (Ano.Text.Trim() != "" && Proyecto.Text.Trim() != "" && tipo2.Text.Trim() != "")
                {
                    string url = "G:/SGC-PROYECTOS-CBR/SGC/" + Ano.Text + "/" + tipo2.Text + "/" + Proyecto.SelectedValue.ToString();

                    consultacarpeta(Proyecto.Text);
                  
                    navegador.Text = url;
                    valor2 = tipo2.Text;
                    valor = Proyecto.SelectedValue.ToString();
                    contadortop();
                    this.registro_AccesosTableAdapter.Visualizar(Proyecto.Text, Ano.Text, tipo2.Text, usuario + " " + apellido);
                    this.registro_Accesos_monitorTableAdapter.ingresamonitor(Proyecto.Text, Ano.Text, tipo2.Text, usuario + " " + apellido);
                  
                    webBrowser1.Refresh();
                    webBrowser1.Url = new Uri(url); 
                    this.folio_ProyectosTableAdapter.Consulta(folio_Proyectos._Folio_Proyectos, nomproyec1);
                    

                }
                else if (Proyecto.Text.Trim() == "" & Ano.Text.Trim() == "" & tipo2.Text.Trim() == "")
                {
                    MessageBox.Show("No ha ingresado ningun dato a buscar");
                }
                else if (Ano.Text.Trim() != "" & tipo2.Text.Trim() == "" & Proyecto.Text.Trim() == "")
                {
                    string url = "G:/SGC-PROYECTOS-CBR/SGC/" + Ano.Text;
                    webBrowser1.Url = new Uri(url);
                    navegador.Text = url;
                    valor = Proyecto.SelectedValue.ToString();
                    contadortop();
                    this.registro_AccesosTableAdapter.Visualizar(Proyecto.Text, Ano.Text, tipo2.Text, usuario + " " + apellido);
                    this.registro_Accesos_monitorTableAdapter.ingresamonitor(Proyecto.Text, Ano.Text, tipo2.Text, usuario + " " + apellido);
                    this.folio_ProyectosTableAdapter.Consulta(folio_Proyectos._Folio_Proyectos, Nombre1);
                    webBrowser1.Refresh();
                 
                }
                else if (Ano.Text.Trim() != "" & tipo2.Text.Trim() != "" & Proyecto.Text.Trim() == "")
                {
                    string url = "G:/SGC-PROYECTOS-CBR/SGC/" + Ano.Text + "/" + tipo2.Text;
                    webBrowser1.Url = new Uri(url);
                    navegador.Text = url;
                    valor = Proyecto.SelectedValue.ToString();
                    contadortop();
                    this.registro_AccesosTableAdapter.Visualizar(Proyecto.Text, Ano.Text, tipo2.Text, usuario + " " + apellido);
                    this.registro_Accesos_monitorTableAdapter.ingresamonitor(Proyecto.Text, Ano.Text, tipo2.Text, usuario + " " + apellido);
                    this.folio_ProyectosTableAdapter.Consulta(folio_Proyectos._Folio_Proyectos, Nombre1);
                    webBrowser1.Refresh();
                   
                }
                else if (Ano.Text.Trim() == "" & tipo2.Text.Trim() != "" & Proyecto.Text.Trim() == "")
                {
                    MessageBox.Show("Dato invalido ingrese almenos el año a buscar");
                }
                else if (Ano.Text.Trim() == "" & tipo2.Text.Trim() == "" & Proyecto.Text.Trim() != "")
                {
                    MessageBox.Show("Dato invalido ingrese  el tipo de proyecto y el año");
                }
                else if (Ano.Text.Trim() == "" & tipo2.Text.Trim() != "" & Proyecto.Text.Trim() != "")
                {
                    MessageBox.Show("Dato invalido ingrese el año");
                }
                else if (Ano.Text.Trim() != "" & tipo2.Text.Trim() == "" & Proyecto.Text.Trim() != "")
                {
                    MessageBox.Show("Dato invalido el tipo");
                }
                this.folio_ProyectosTableAdapter.Consulta(folio_Proyectos._Folio_Proyectos, nomproyec1);
                Nombre1 = nomproyec1;
                this.folio_ProyectosTableAdapter.Consulta(folio_Proyectos._Folio_Proyectos, Nombre1);
          

            }

        


        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            Process.Start(@"G:\SGC\Formatos\FVT.17 Alta de cliente COCOA_1.xlsx");
        }

        private void button8_Click(object sender, EventArgs e)
        {
            Process.Start(@"G:\SGC\Formatos\FCP.07 Lista de Materiales y Requisición_4.xlsx");
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Process.Start(@"G:\SGC\Formatos\FCP.07 Lista de Materiales y Requisición_4.xlsx");
        }

        //private void pictureBox8_Click(object sender, EventArgs e)
        //{
        //    if (Tipo.Text == "Alto flujo")
        //    {

        //        carptape = AF.Text;
        //    }
        //    if (Tipo.Text == "Compacta")
        //    {

        //        carptape = CP.Text;
        //    }
        //    if (Tipo.Text == "Industrial Potabiliza")
        //    {

        //        carptape = IND.Text;
        //    }
        //    if (Tipo.Text == "Inox")
        //    {

        //        carptape = INOX.Text;
        //    }
        //    if (Tipo.Text == "")
        //    {

        //        MessageBox.Show("No ha seleccionado el tipo de proyecto");
        //    }
        //    if (Tipo.Text != "")
        //    {

        //        if (carptape != "")
        //        {
        //            DialogResult Resultado;
        //            Resultado = MessageBox.Show("Desea reasignar permisos a la carpeta:" + carptape + "del proyecto: " + "" + proye_solicita.Text, "Confirmación", MessageBoxButtons.YesNo);
        //            if (Resultado == DialogResult.Yes)
        //            {
        //                consultacarpeta(proye_solicita.Text);

        //                //////////////////////// Crear carpeta Temporal ///////////////////////////////////////////////////
        //                string temporal;
        //                temporal = "/C MD G:\\Sistema\\Temporales\\Temporal\\";
        //                string temporalbackup;
        //                string carpetaf = "" + proye_solicita.SelectedValue.ToString() + "-" + carptape + "";
        //                temporalbackup = "/C MD G:\\Sistema\\Temporales-Backup\\" + carpetaf;
        //                string carpetaf2 = proye_solicita.Text + "-" + carptape;
        //                //////////////////////////////////////////////////////////////////////////////////////////////////

        //                ////////////////////////// Copia archivos a Temporal ///////////////////////////////////////////////////
        //                string copiatemporal = Environment.CurrentDirectory;
        //                copiatemporal = @"/C robocopy" + " " + " \"G:\\SGC-PROYECTOS-CBR\\SGC\\" + anorastreo + "\\" + carpetarastreo + "\\" + proyectorastreo + "\\" + carptape + "\" " + " G:\\Sistema\\Temporales\\Temporal";

        //                ////////////////////////// Copia archivos a Temporal backup///////////////////////////////////////////////////
        //                string copiatemporalbackup = Environment.CurrentDirectory;
        //                copiatemporalbackup = @"/C robocopy" + " " + " \"G:\\SGC-PROYECTOS-CBR\\SGC\\" + anorastreo + "\\" + carpetarastreo + "\\" + proyectorastreo + "\\" + carptape + "\" " + "\"" + "G:\\Sistema\\Temporales-Backup\\" + carpetaf2 + "\"";

        //                //////////////////////////////////////////////////////////////////////////////////////////////////////

        //                //////////////////////////// elimina carpeta con permiso previo ///////////////////////////////////////////////////
        //                string elimina;
        //                elimina = "/C rd" + " " + " \"G:\\SGC-PROYECTOS-CBR\\SGC\\" + anorastreo + "\\" + carpetarastreo + "\\" + proyectorastreo + "\\" + carptape + "\" /S /Q ";

        //                ////////////////////////////////////////////////////////////////////////////////////////////////////

        //                /////////////////////////crea proceso de rastreo de proyecto //////////////////////////////////////
        //                string nuevoperm = "";
        //                if (carpetarastreo == "Proyectos-Compacta") { nuevoperm = "CodificadoresLideres-Af-Comp-pos"; }
        //                if (carpetarastreo == "Proyectos-Alto-Flujo") { nuevoperm = "CodificadoresLideres-Af-Comp-pos"; }
        //                if (carpetarastreo == "Proyectos-Industrial-Potabiliza") { nuevoperm = "CodificadoresLideres-Indus-Pot-pos"; }
        //                if (carpetarastreo == "Proyectos-Inox") { nuevoperm = "Proyecto-INOX-pos"; codifica = ""; }

        //                //////////////////////////// obtiene carpeta con nuevos permisos ///////////////////////////////////////////////////
        //                string recrea;
        //                recrea = @"/C  robocopy" + " " + "G:\\SGC\\" + anorastreo + "\\" + nuevoperm + "\\" + codifica + " " + "G:\\SGC-PROYECTOS-CBR\\SGC\\" + anorastreo + "\\" + carpetarastreo + "\\" + proyectorastreo + "" + " " + "/E /SEC /R:0";
        //                // recrea = @"/C robocopy" + " " + "B:\\SGC\\" + anorastreo + "\\" + nuevoperm + " " + " B:\\Proyectos-CBR\\SGC\\" + anorastreo + "\\" + carpetarastreo + "\\" + proye.Text + "" + " " + "/E /SEC";

        //                //////////////////////////////////////////////////////////////////////////////////////////////////
        //                ////////////////////////// obtiene archivos a nueva carpeta con permisos ///////////////////////////////////////////////////
        //                string copia;

        //                //recrea = @"/C robocopy" + " " + "B:\\SGC\\" + anorastreo + "\\" + nuevoperm  + " " + "B:\\Proyectos-CBR\\SGC\\"+anorastreo+"\\"+carpetarastreo+"\\"+proye.Text+"\\  /E / SEC";
        //                copia = @"/C xcopy " + " " + "\"" + "G:\\Sistema\\Temporales-Backup\\"+ carpetaf + "\""+"  " + "\"" + "G:\\SGC-PROYECTOS-CBR\\SGC\\" + anorastreo + "\\" + carpetarastreo + "\\" + proyectorastreo + "\\" + carptape + "\"" + "" + " " + "/S /Q  /i";
                   
        //                //////////////////////////////////////////////////////////////////////////////////////////////////
        //                //////////////////////////// elimina carpeta temporal ///////////////////////////////////////////////////
        //                string eliminatem;
        //                eliminatem = "/C rd" + " " + " \"G:\\Sistema\\Temporales\\Temporal" + "\" /S /Q";




        //                this.migra_permisosTableAdapter.codificador(usuario, apellido, carpetarastreo, carptape, proyectorastreo, elimina, elimina, elimina, elimina, elimina, recrea, copia, eliminatem);
        //                //Process crea1 = System.Diagnostics.Process.Start("CMD.exe", temporal);
        //                //crea1.WaitForExit();

        //                //Thread.Sleep(1000);
        //                //Process crea = System.Diagnostics.Process.Start("CMD.exe", temporalbackup);
        //                //crea.WaitForExit();
        //                //Thread.Sleep(1000);
        //                //Process copia1 = System.Diagnostics.Process.Start("CMD.exe", copiatemporal);
        //                //copia1.WaitForExit();
        //                //Thread.Sleep(1000);
        //                //Process copia1back = System.Diagnostics.Process.Start("CMD.exe", copiatemporalbackup);
        //                //copia1back.WaitForExit();
        //                //Thread.Sleep(1000);
        //                //Process elimin = System.Diagnostics.Process.Start("CMD.exe", elimina);
        //                //elimin.WaitForExit();
        //                //Thread.Sleep(2000);
        //                //Process recrear = System.Diagnostics.Process.Start("CMD.exe", recrea);
        //                //recrear.WaitForExit();
        //                //Thread.Sleep(1000);
        //                //Process cop = System.Diagnostics.Process.Start("CMD.exe", copia);
        //                //cop.WaitForExit();
        //                //Thread.Sleep(1000);
        //                //Process elim = System.Diagnostics.Process.Start("CMD.exe", eliminatem);
        //                //elim.WaitForExit();
        //                ////Process netStat = new Process();
        //                ////Process elim = System.Diagnostics.Process.Start(@"C:\Windows\system32\cmd.exe", @"/c psexec \\192.168.1.203 -u 192.168.1.48203\Administrador -p asdasd -s cmd.exe");

        //                //elim.WaitForExit();
        //                MessageBox.Show("Actualizacion subida espere correos de confirmacion");





        //            }
        //            else if (Resultado == DialogResult.No)
        //            {


        //            }
        //        }
        //        else if (carptape == "")
        //        {
        //            MessageBox.Show("No ha seleccionado la carpeta a actualizar");
        //        }


        //    }
        //}

        private void INOX_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void proye_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void pictureBox9_Click(object sender, EventArgs e)
        {

            if (PedidoA.Text == "")
            {

                MessageBox.Show("No ha ingresado el nombre del proyecto");
            }
            else
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Desea solicitar autorizacion del pedido con nombre de proyecto:" + PedidoA.Text, "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    string _sender = "support@cbr-ingenieria.com.mx";
                    string _password = "Cbrsoporte2020.";
                    SmtpClient client = new SmtpClient("smtp.office365.com");
                    client.Port = 587;
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.UseDefaultCredentials = false;
                    System.Net.NetworkCredential credentials =
                     new System.Net.NetworkCredential(_sender, _password);
                    client.EnableSsl = true;
                    client.Credentials = credentials;
                    MailMessage message = new MailMessage(_sender, "cecilia@cbr-ingenieria.com.mx");
                    message.CC.Add(new MailAddress("alejandro@cbr-ingenieria.com.mx"));
                    message.CC.Add(new MailAddress(emailusuario));
                    message.Subject = "Solicitud de actorizacion de Pedido";
                    message.Priority = MailPriority.High;
                    message.IsBodyHtml = true;
                    message.AlternateViews.Add(getEmbeddedImage3("//192.168.1.101/Aplicativos/CBR Sistema global/logo_cbr.png"));
                    client.Send(message);
                }
                else if (Resultado == DialogResult.No)
                {


                }
            }

        }

        private void panel5_Paint(object sender, PaintEventArgs e)
        {

        }

        private void pictureBox10_Click(object sender, EventArgs e)
        {
            Consulta_Proyectos crear = new Consulta_Proyectos();

            crear.Show();
        }

        private void tipo2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tipo2.Text == "Proyectos-Alto-Flujo")
            {

                indus_po.Visible = false;
                Cb_af_com.Visible = true;
                cb_inox.Visible = false;

                if (Ano.Text == "2017" && departamento != "SuperAdmin" && Ano.Text == "2017" && departamento != "Direccion")
                {
                    Cb_af_com.Text = departamento;
                }
                else if (Ano.Text == "2018" && departamento != "SuperAdmin" && Ano.Text == "2018" && departamento != "Direccion")
                {
                    Cb_af_com.Text = departamento;
                }
                else if (Ano.Text == "2019" && departamento != "SuperAdmin" && Ano.Text == "2019" && departamento != "Direccion")
                {
                    Cb_af_com.Text = departamento;
                }


                else { }

            }
            if (tipo2.Text == "Proyectos-Compacta")
            {
                indus_po.Visible = false;
                Cb_af_com.Visible = true;
                cb_inox.Visible = false;
                if (Ano.Text == "2017" && departamento != "SuperAdmin" && Ano.Text == "2017" && departamento != "Direccion")
                {
                    Cb_af_com.Text = departamento;
                }
                else if (Ano.Text == "2018" && departamento != "SuperAdmin" && Ano.Text == "2018" && departamento != "Direccion")
                {
                    Cb_af_com.Text = departamento;
                }
                else if (Ano.Text == "2019" && departamento != "SuperAdmin" && Ano.Text == "2019" && departamento != "Direccion")
                {
                    Cb_af_com.Text = departamento;
                }


                else { }
            }
            if (tipo2.Text == "Proyectos-Industrial-Potabiliza")
            {

                indus_po.Visible = true;
                Cb_af_com.Visible = false;
                cb_inox.Visible = false;
                if (Ano.Text == "2017" && departamento != "SuperAdmin" && Ano.Text == "2017" && departamento != "Direccion")
                {
                    indus_po.Text = departamento;
                }
                else if (Ano.Text == "2018" && departamento != "SuperAdmin" && Ano.Text == "2018" && departamento != "Direccion")
                {
                    indus_po.Text = departamento;
                }
                else if (Ano.Text == "2019" && departamento != "SuperAdmin" && Ano.Text == "2019" && departamento != "Direccion")
                {
                    indus_po.Text = departamento;
                }


                else { }
            }
            if (tipo2.Text == "Proyectos-Inox")
            {

                indus_po.Visible = false;
                Cb_af_com.Visible = false;
                cb_inox.Visible = true;
                if (Ano.Text == "2017" && departamento != "SuperAdmin" && Ano.Text == "2017" && departamento != "Direccion")
                {
                    cb_inox.Text = departamento;
                }
                else if (Ano.Text == "2018" && departamento != "SuperAdmin" && Ano.Text == "2018" &&  departamento != "Direccion")
                {
                    cb_inox.Text = departamento;
                }
               else  if (Ano.Text == "2019" && departamento != "SuperAdmin" && Ano.Text == "2019" &&  departamento != "Direccion")
                {
                    cb_inox.Text = departamento;
                }


                else if (Ano.Text == "2017"|| Ano.Text == "2018" || Ano.Text == "2019")
                {

                
                    
                   

                }
                else { }

            }
            if (tipo2.Text == "")
            {

                indus_po.Visible = false;
                Cb_af_com.Visible = false;
                cb_inox.Visible = false;
            }
        }

        private void Proyecto_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            this.panel3.Size = new System.Drawing.Size(502, 337);
        }

        private void pictureBox13_Click(object sender, EventArgs e)
        {


            DialogResult Resultado;
            Resultado = MessageBox.Show("Está seguro que desea terminar la sesion en el proyecto: "+Proyecto.Text, "Confirmación", MessageBoxButtons.YesNo);
            if (Resultado == DialogResult.Yes)
            { rastreadorproyecto();
                this.registro_AccesosTableAdapter.elimina(Proyecto.Text);

                MessageBox.Show("Session cerrada correctamente");
                webBrowser1.GoBack();
                Nombre1 = "";
                valor = "";

               
          
                string strCmdText;
                strCmdText = "/C attrib +h +s G:\\SGC-PROYECTOS-CBR\\SGC\\"+Año2+"\\"+tipo3+"\\"+Nombre2+ "\\* /S /D";
                Process cop = System.Diagnostics.Process.Start("CMD.exe", strCmdText);
                cop.WaitForExit();

            }
            else if (Resultado == DialogResult.No)
            {
              

                MessageBox.Show("Operacion Cancelada");
            }
          
        }

        private void Proyecto_Click(object sender, EventArgs e)
        {
            this.folio_ProyectosTableAdapter.Fill(folio_Proyectos._Folio_Proyectos);
        }
        public void rastreadorproyecto()
        {
            SqlConnection conexion = new SqlConnection(ObtenerCadena());
            conexion.Open();
            SqlCommand cmd = new SqlCommand(
                                   "select [Proyecto]," +
                                   "[ano], " +
                                   "[tipo], " +
                                   "[Usuario] " +

                                   "from [Registro_Accesos]  where Proyecto=@proyecto"

                                   , conexion);
            cmd.Parameters.AddWithValue("proyecto", Proyecto.Text);
            SqlDataAdapter sda = new SqlDataAdapter(cmd);
            sda.SelectCommand.CommandTimeout = 36000;
            DataTable dt = new DataTable();
            sda.Fill(dt);
            if (dt.Rows.Count > 0)
            {
                Nombre2 = dt.Rows[0][0].ToString();
                Año2 = dt.Rows[0][1].ToString();
                tipo3 = dt.Rows[0][2].ToString();
                
                conexion.Dispose();
            }
            else
            {
              
                conexion.Dispose();
            }

        }


        public void contadortop()
        {
            SqlConnection conexion = new SqlConnection(ObtenerCadena());

            conexion.Open();
            ////////////////////////////////////////////////////////////////////////////////////
            if (valor2 == "Proyectos-Inox") { valor2 = "inox"; }
            else if (valor2 == "Proyectos-Compacta") { valor2 = "compacta"; }
            else if (valor2 == "Proyectos-Alto-Flujo") { valor2 = "altoflujo"; }
            else if (valor2 == "Proyectos-Industrial-Potabiliza") { valor2 = "Industrial"; }
            else { }


            SqlCommand cmd = new SqlCommand();
            if (departamento == "SuperAdmin" && valor2 == "inox")
            {
                cmd = new SqlCommand(
                "select " +
                "Count('Nombre') " +
                "from [Term_inox_admin] "
                , conexion);  }
          else  if (departamento == "Direccion" && valor2 == "inox")
            {
                cmd = new SqlCommand(
                "select " +
                "Count('Nombre') " +
                "from [Term_inox_admin] "
                , conexion);
            }

            else if (departamento == "Asistente de direccion" && valor2 == "inox")
            {
                cmd = new SqlCommand(
                "select " +
                "Count('Nombre') " +
                "from [Term_inox_asistedirec] "
                , conexion);
            }
            else if (departamento == "Infraestructura" && valor2 == "inox")
            {
                cmd = new SqlCommand(
                "select " +
                "Count('Nombre') " +
                "from [Term_inox_infraestructura] "
                , conexion);
            }
            else if (departamento == "Potabiliza" && valor2 == "inox")
            {
                cmd = new SqlCommand(
                "select " +
                "Count('Nombre') " +
                "from [Term_inox_potabiliza] "
                , conexion);
            }
            else if (departamento == "Ventas" && valor2 == "inox")
            {
                cmd = new SqlCommand(
                "select " +
                "Count('Nombre') " +
                "from [Term_inox_ventas] "
                , conexion);
            }
            else if (departamento == "Proyectos" && valor2 == "inox")
            {
                cmd = new SqlCommand(
                "select " +
                "Count('Nombre') " +
                "from [Term_inox_proyectos] "
                , conexion);
            }
            else if (departamento == "Atencion a Clientes" && valor2 == "inox")
            {
                cmd = new SqlCommand(
                "select " +
                "Count('Nombre') " +
                "from [Term_inox_atnclientes] "
                , conexion);
            }
            else if (departamento == "Produccion" && valor2 == "inox")
            {
                cmd = new SqlCommand(
                "select " +
                "Count('Nombre') " +
                "from [Term_inox_produccion] "
                , conexion);
            }

            else if (departamento == "Construccion e Instalaciones" && valor2 == "inox")
            {
                cmd = new SqlCommand(
                "select " +
                "Count('Nombre') " +
                "from [Term_inox_consinst] "
                , conexion);
            }

            else if (departamento == "Operacion y Mantenimiento" && valor2 == "inox")
            {
                cmd = new SqlCommand(
                "select " +
                "Count('Nombre') " +
                "from [Term_inox_opyman] "
                , conexion);
            }
            else if (departamento == "Compras" && valor2 == "inox")
            {
                cmd = new SqlCommand(
                "select " +
                "Count('Nombre') " +
                "from [Term_inox_compras] "
                , conexion);
            }


            if (departamento == "SuperAdmin" && valor2 == "altoflujo"|| departamento == "SuperAdmin" && valor2 == "compacta")
            {
                cmd = new SqlCommand(
                "select " +
                "Count('Nombre') " +
                "from [Term_af-com_admin] "
                , conexion);
            }
            else if (departamento == "Direccion" && valor2 == "altoflujo"|| departamento == "Direccion" && valor2 == "compacta")
            {
                cmd = new SqlCommand(
                "select " +
                "Count('Nombre') " +
                "from [Term_af-com_admin] "
                , conexion);
            }
            if (departamento == "SuperAdmin" && valor2 == "altoflujo" || departamento == "SuperAdmin" && valor2 == "Industrial")
            {
                cmd = new SqlCommand(
                "select " +
                "Count('Nombre') " +
                "from [Term_Indus-Pot_admin] "
                , conexion);
            }
            else if (departamento == "Direccion" && valor2 == "altoflujo" || departamento == "Direccion" && valor2 == "Industrial")
            {
                cmd = new SqlCommand(
                "select " +
                "Count('Nombre') " +
                "from [Term_Indus-Pot_admin] "
                , conexion);
            }

            else if (departamento == "Asistente de direccion" && valor2 == "compacta" || departamento == "Asistente de direccion" && valor2 == "altoflujo")
            {
                cmd = new SqlCommand(
                "select " +
                "Count('Nombre') " +
                "from [Term_af-com_asistedirec] "
                , conexion);
            }
            else if (departamento == "Infraestructura" && valor2 == "compacta"|| departamento == "Infraestructura" && valor2 == "altoflujo")
            {
                cmd = new SqlCommand(
                "select " +
                "Count('Nombre') " +
                "from [Term_af-com_infraestructura] "
                , conexion);
            }
            else if (departamento == "Potabiliza" && valor2 == "compacta"|| departamento == "Potabiliza" && valor2 == "altoflujo")
            {
                cmd = new SqlCommand(
                "select " +
                "Count('Nombre') " +
                "from [Term_af-com_potabiliza] "
                , conexion);
            }
            else if (departamento == "Ventas" && valor2 == "compacta"|| departamento == "Ventas" && valor2 == "altoflujo")
            {
                cmd = new SqlCommand(
                "select " +
                "Count('Nombre') " +
                "from [Term_af-com_ventas] "
                , conexion);
            }
            else if (departamento == "Proyectos" && valor2 == "compacta"|| departamento == "Proyectos" && valor2 == "altoflujo")
            {
                cmd = new SqlCommand(
                "select " +
                "Count('Nombre') " +
                "from [Term_af-com_proyectos] "
                , conexion);
            }
            else if (departamento == "Atencion a Clientes" && valor2 == "compacta"|| departamento == "Atencion a Clientes" && valor2 == "altoflujo")
            {
                cmd = new SqlCommand(
                "select " +
                "Count('Nombre') " +
                "from [Term_af-com_atnclientes] "
                , conexion);
            }
            else if (departamento == "Produccion" && valor2 == "compacta"|| departamento == "Produccion" && valor2 == "altoflujo")
            {
                cmd = new SqlCommand(
                "select " +
                "Count('Nombre') " +
                "from [Term_af-com_produccion] "
                , conexion);
            }
            else if (departamento == "Construccion e Instalaciones" && valor2 == "compacta"|| departamento == "Construccion e Instalaciones" && valor2 == "altoflujo")
            {
                cmd = new SqlCommand(
                "select " +
                "Count('Nombre') " +
                "from [Term_af-com_consinst] "
                , conexion);
            }
            else if (departamento == "Operacion y Mantenimiento" && valor2 == "compacta"|| departamento == "Operacion y Mantenimiento" && valor2 == "altoflujo")
            {
                cmd = new SqlCommand(
                "select " +
                "Count('Nombre') " +
                "from [Term_af-com_opyman] "
                , conexion);
            }
            else if (departamento == "Compras" && valor2 == "compacta"|| departamento == "Compras" && valor2 == "altoflujo")
            {
                cmd = new SqlCommand(
                "select " +
                "Count('Nombre') " +
                "from [Term_af-com_compras] "
                , conexion);
            }


            else if (departamento == "Asistente de direccion" && valor2 == "Industrial" || departamento == "Asistente de direccion" && valor2 == "Potabiliza")
            {
                cmd = new SqlCommand(
                "select " +
                "Count('Nombre') " +
                "from [Term_Indus-Pot_asistedirec] "
                , conexion);
            }
            else if (departamento == "Infraestructura" && valor2 == "Industrial" || departamento == "Infraestructura" && valor2 == "Potabiliza")
            {
                cmd = new SqlCommand(
                "select " +
                "Count('Nombre') " +
                "from [Term_Indus-Pot_infraestructura] "
                , conexion);
            }
            else if (departamento == "Potabiliza" && valor2 == "Industrial" || departamento == "Potabiliza" && valor2 == "Potabiliza")
            {
                cmd = new SqlCommand(
                "select " +
                "Count('Nombre') " +
                "from [Term_Indus-Pot_potabiliza] "
                , conexion);
            }
            else if (departamento == "Ventas" && valor2 == "Industrial" || departamento == "Ventas" && valor2 == "Potabiliza")
            {
                cmd = new SqlCommand(
                "select " +
                "Count('Nombre') " +
                "from [Term_Indus-Pot_ventas] "
                , conexion);
            }
            else if (departamento == "Proyectos" && valor2 == "Industrial" || departamento == "Proyectos" && valor2 == "Potabiliza")
            {
                cmd = new SqlCommand(
                "select " +
                "Count('Nombre') " +
                "from [Term_Indus-Pot_proyectos] "
                , conexion);
            }
            else if (departamento == "Atencion a Clientes" && valor2 == "Industrial" || departamento == "Atencion a Clientes" && valor2 == "Potabiliza")
            {
                cmd = new SqlCommand(
                "select " +
                "Count('Nombre') " +
                "from [Term_Indus-Pot_atnclientes] "
                , conexion);
            }
            else if (departamento == "Produccion" && valor2 == "Industrial" || departamento == "Produccion" && valor2 == "Potabiliza")
            {
                cmd = new SqlCommand(
                "select " +
                "Count('Nombre') " +
                "from [Term_Indus-Pot_produccion] "
                , conexion);
            }
            else if (departamento == "Construccion e Instalaciones" && valor2 == "Industrial" || departamento == "Construccion e Instalaciones" && valor2 == "Potabiliza")
            {
                cmd = new SqlCommand(
                "select " +
                "Count('Nombre') " +
                "from [Term_Indus-Pot_consinst] "
                , conexion);
            }
            else if (departamento == "Operacion y Mantenimiento" && valor2 == "Industrial" || departamento == "Operacion y Mantenimiento" && valor2 == "Potabiliza")
            {
                cmd = new SqlCommand(
                "select " +
                "Count('Nombre') " +
                "from [Term_Indus-Pot_opyman] "
                , conexion);
            }
            else if (departamento == "Compras" && valor2 == "Industrial" || departamento == "Compras" && valor2 == "Potabiliza")
            {
                cmd = new SqlCommand(
                "select " +
                "Count('Nombre') " +
                "from [Term_Indus-Pot_compras] "
                , conexion);
            }



            SqlDataAdapter sda = new SqlDataAdapter(cmd);
            sda.SelectCommand.CommandTimeout = 36000;
            DataTable dt = new DataTable();
            sda.Fill(dt);



            conexion.Dispose();
            if (dt.Rows.Count > 0)
            {
                Nombre = dt.Rows[0][0].ToString();
          
                    contador();
              
            }
            else { }
       
        }
        public void Mostrador()
        {
           
          

            SqlConnection conexion = new SqlConnection(ObtenerCadena());
           
            SqlCommand cmd = new SqlCommand();
            if (departamento == "SuperAdmin" && valor2 == "inox")
            {
                cmd = new SqlCommand(
                                   "select top " + "(" + suma + ")" +
                                "[Nombre] " +

                                   "from [Term_inox_admin] "

                                   , conexion);
            }

            else if (departamento == "Direccion" && valor2 == "inox")
            {
                cmd = new SqlCommand(
                    "select top " + "(" + suma + ")" +
                                "[Nombre] " +

                "from [Term_inox_admin] "
                , conexion);
            }
            else if (departamento == "Asistente de direccion" && valor2 == "inox")
            {
                cmd = new SqlCommand(
                "select top " + "(" + suma + ")" +
                                "[Nombre] " +

                "from [Term_inox_asistedirec] "
                , conexion);
            }
            else if (departamento == "Infraestructura" && valor2 == "inox")
            {
                cmd = new SqlCommand(
                     "select top " + "(" + suma + ")" +
                                "[Nombre] " +

                "from [Term_inox_infraestructura] "
                , conexion);
            }
            else if (departamento == "Potabiliza" && valor2 == "inox")
            {
                cmd = new SqlCommand(
                     "select top " + "(" + suma + ")" +
                                "[Nombre] " +

                "from [Term_inox_potabiliza] "
                , conexion);
            }
            else if (departamento == "Ventas" && valor2 == "inox")
            {
                cmd = new SqlCommand(
             "select top " + "(" + suma + ")" +
                                "[Nombre] " +

                "from [Term_inox_ventas] "
                , conexion);
            }
            else if (departamento == "Proyectos" && valor2 == "inox")
            {
                cmd = new SqlCommand(
                   "select top " + "(" + suma + ")" +
                                "[Nombre] " +

                "from [Term_inox_proyectos] "
                , conexion);
            }
            else if (departamento == "Atencion a Clientes" && valor2 == "inox")
            {
                cmd = new SqlCommand(
                      "select top " + "(" + suma + ")" +
                                "[Nombre] " +

                "from [Term_inox_atnclientes] "
                , conexion);
            }
            else if (departamento == "Produccion" && valor2 == "inox")
            {
                cmd = new SqlCommand(
           "select top " + "(" + suma + ")" +
                                "[Nombre] " +

                "from [Term_inox_produccion] "
                , conexion);
            }
            else if (departamento == "Construccion e Instalaciones" && valor2 == "inox")
            {
                cmd = new SqlCommand(
                    "select top " + "(" + suma + ")" +
                                "[Nombre] " +

                "from [Term_inox_consinst] "
                , conexion);
            }
            else if (departamento == "Operacion y Mantenimiento" && valor2 == "inox")
            {
                cmd = new SqlCommand(
                    "select top " + "(" + suma + ")" +
                                "[Nombre] " +

                "from [Term_inox_opyman] "
                , conexion);
            }
            else if (departamento == "Compras" && valor2 == "inox")
            {
                cmd = new SqlCommand(
                    "select top " + "(" + suma + ")" +
                                "[Nombre] " +

                "from [Term_inox_compras] "
                , conexion);
            }

            else if (departamento == "SuperAdmin" && valor2 == "altoflujo" || departamento == "SuperAdmin" && valor2 == "compacta")
            {
                cmd = new SqlCommand(
                    "select top " + "(" + suma + ")" +
                                "[Nombre] " +

                "from [Term_af-com_admin] "
                , conexion);
            }
            else if (departamento == "Direccion" && valor2 == "altoflujo" || departamento == "Direccion" && valor2 == "compacta")
            {
                cmd = new SqlCommand(
                "select top " + "(" + suma + ")" +
                                "[Nombre] " +

                "from [Term_af-com_admin] "
                , conexion);
            }
            else if (departamento == "Asistente de direccion" && valor2 == "altoflujo"|| departamento == "Asistente de direccion" && valor2 == "compacta")
            {
                cmd = new SqlCommand(
                "select top " + "(" + suma + ")" +
                                "[Nombre] " +

                "from [Term_af-com_asistedirec] "
                , conexion);
            }
            else if (departamento == "Infraestructura" && valor2 == "compacta" || departamento == "Infraestructura" && valor2 == "altoflujo")
            {
                cmd = new SqlCommand(
                     "select top " + "(" + suma + ")" +
                                "[Nombre] " +

                "from [Term_af-com_infraestructura] "
                , conexion);
            }
            else if (departamento == "Potabiliza" && valor2 == "compacta" || departamento == "Potabiliza" && valor2 == "altoflujo")
            {
                cmd = new SqlCommand(
                     "select top " + "(" + suma + ")" +
                                "[Nombre] " +

                "from [Term_af-com_potabiliza] "
                , conexion);
            }
            else if (departamento == "Ventas" && valor2 == "compacta" || departamento == "Ventas" && valor2 == "altoflujo")
            {
                cmd = new SqlCommand(
             "select top " + "(" + suma + ")" +
                                "[Nombre] " +

                "from [Term_af-com_ventas] "
                , conexion);
            }
            else if (departamento == "Proyectos" && valor2 == "compacta" || departamento == "Proyectos" && valor2 == "altoflujo")
            {
                cmd = new SqlCommand(
                   "select top " + "(" + suma + ")" +
                                "[Nombre] " +

                "from [Term_af-com_proyectos] "
                , conexion);
            }
            else if (departamento == "Atencion a Clientes" && valor2 == "compacta" || departamento == "Atencion a Clientes" && valor2 == "altoflujo")
            {
                cmd = new SqlCommand(
                      "select top " + "(" + suma + ")" +
                                "[Nombre] " +

                "from [Term_af-com_atnclientes] "
                , conexion);
            }
            else if (departamento == "Produccion" && valor2 == "compacta" || departamento == "Produccion" && valor2 == "altoflujo")
            {
                cmd = new SqlCommand(
           "select top " + "(" + suma + ")" +
                                "[Nombre] " +

                "from [Term_af-com_produccion] "
                , conexion);
            }
            else if (departamento == "Construccion e Instalaciones" && valor2 == "compacta" || departamento == "Construccion e Instalaciones" && valor2 == "altoflujo")
            {
                cmd = new SqlCommand(
                    "select top " + "(" + suma + ")" +
                                "[Nombre] " +

                "from [Term_af-com_consinst] "
                , conexion);
            }
            else if (departamento == "Operacion y Mantenimiento" && valor2 == "compacta" || departamento == "Operacion y Mantenimiento" && valor2 == "altoflujo")
            {
                cmd = new SqlCommand(
                    "select top " + "(" + suma + ")" +
                                "[Nombre] " +

                "from [Term_af-com_opyman] "
                , conexion);
            }
            else if (departamento == "Compras" && valor2 == "compacta"|| departamento == "Compras" && valor2 == "altoflujo")
            {
                cmd = new SqlCommand(
                    "select top " + "(" + suma + ")" +
                                "[Nombre] " +

                "from [Term_af-com_compras] "
                , conexion);
            }


            else if (departamento == "SuperAdmin" && valor2 == "Industrial" || departamento == "SuperAdmin" && valor2 == "Potabiliza")
            {
                cmd = new SqlCommand(
                    "select top " + "(" + suma + ")" +
                                "[Nombre] " +

                "from [Term_Indus-Pot_admin] "
                , conexion);
            }
            else if (departamento == "Direccion" && valor2 == "Industrial" || departamento == "Direccion" && valor2 == "Potabiliza")
            {
                cmd = new SqlCommand(
                "select top " + "(" + suma + ")" +
                                "[Nombre] " +

                "from [Term_Indus-Pot_admin] "
                , conexion);
            }
            else if (departamento == "Asistente de direccion" && valor2 == "Industrial" || departamento == "Asistente de direccion" && valor2 == "Potabiliza")
            {
                cmd = new SqlCommand(
                "select top " + "(" + suma + ")" +
                                "[Nombre] " +

                "from [Term_Indus-Pot_asistedirec] "
                , conexion);
            }
            else if (departamento == "Infraestructura" && valor2 == "Industrial" || departamento == "Infraestructura" && valor2 == "Potabiliza")
            {
                cmd = new SqlCommand(
                     "select top " + "(" + suma + ")" +
                                "[Nombre] " +

                "from [Term_Indus-Pot_infraestructura] "
                , conexion);
            }
            else if (departamento == "Potabiliza" && valor2 == "Industrial" || departamento == "Potabiliza" && valor2 == "Potabiliza")
            {
                cmd = new SqlCommand(
                     "select top " + "(" + suma + ")" +
                                "[Nombre] " +

                "from [Term_Indus-Pot_potabiliza] "
                , conexion);
            }
            else if (departamento == "Ventas" && valor2 == "Industrial" || departamento == "Ventas" && valor2 == "Potabiliza")
            {
                cmd = new SqlCommand(
             "select top " + "(" + suma + ")" +
                                "[Nombre] " +

                "from [Term_Indus-Pot_ventas] "
                , conexion);
            }
            else if (departamento == "Proyectos" && valor2 == "Industrial" || departamento == "Proyectos" && valor2 == "Potabiliza")
            {
                cmd = new SqlCommand(
                   "select top " + "(" + suma + ")" +
                                "[Nombre] " +

                "from [Term_Indus-Pot_proyectos] "
                , conexion);
            }
            else if (departamento == "Atencion a Clientes" && valor2 == "Industrial" || departamento == "Atencion a Clientes" && valor2 == "Potabiliza")
            {
                cmd = new SqlCommand(
                      "select top " + "(" + suma + ")" +
                                "[Nombre] " +

                "from [Term_Indus-Pot_atnclientes] "
                , conexion);
            }
            else if (departamento == "Produccion" && valor2 == "Industrial" || departamento == "Produccion" && valor2 == "Potabiliza")
            {
                cmd = new SqlCommand(
           "select top " + "(" + suma + ")" +
                                "[Nombre] " +

                "from [Term_Indus-Pot_produccion] "
                , conexion);
            }
            else if (departamento == "Construccion e Instalaciones" && valor2 == "Industrial" || departamento == "Construccion e Instalaciones" && valor2 == "Potabiliza")
            {
                cmd = new SqlCommand(
                    "select top " + "(" + suma + ")" +
                                "[Nombre] " +

                "from [Term_Indus-Pot_consinst] "
                , conexion);
            }
            else if (departamento == "Operacion y Mantenimiento" && valor2 == "Industrial" || departamento == "Operacion y Mantenimiento" && valor2 == "Potabiliza")
            {
                cmd = new SqlCommand(
                    "select top " + "(" + suma + ")" +
                                "[Nombre] " +

                "from [Term_Indus-Pot_opyman] "
                , conexion);
            }
            else if (departamento == "Compras" && valor2 == "Industrial" || departamento == "Compras" && valor2 == "Potabiliza")
            {
                cmd = new SqlCommand(
                    "select top " + "(" + suma + ")" +
                                "[Nombre] " +

                "from [Term_Indus-Pot_compras] "
                , conexion);
            }


            SqlDataAdapter sda = new SqlDataAdapter(cmd);
            sda.SelectCommand.CommandTimeout = 36000;
            DataTable dt = new DataTable();
            sda.Fill(dt);
            if (dt.Rows.Count > 0)
            {
                
                string Nombre = dt.Rows[k - 1][0].ToString();
                Thread.Sleep(10);
                string strCmdText;
                if (departamento == "Construccion e Instalaciones" && Ano.Text == "2019") { departamento = "Instalaciones y construccion"; }
                else if (departamento == "Construccion e Instalaciones" && Ano.Text == "2018") { departamento = "Instalaciones y construccion"; }
                else if (departamento == "Construccion e Instalaciones" && Ano.Text == "2017") { departamento = "Instalaciones y construccion"; }
                else if (departamento == "Construccion e Instalaciones" &&  Ano.Text != "2017" && departamento == "Construccion e Instalaciones" && Ano.Text != "2018" && departamento == "Construccion e Instalaciones" && Ano.Text != "2019") { departamento = "Construccion e Instalaciones"; }
                else { }

                // richTextBox1.Text = richTextBox1.Text+"\n"+ "G:\\SGC-PROYECTOS-CBR\\SGC\\"+Año +"\\"+ tipo +"\\" + Nombre ;
                if (Ano.Text == "2019" && departamento != "SuperAdmin")
                {

                    strCmdText = "/C attrib -h -s " + " \"" + "G:\\SGC-PROYECTOS-CBR\\SGC\\" + Ano.Text + "\\" + tipo2.Text + "\\" + valor + "\\" + departamento + "\\*" + "\"" + "  /S" + "  /D";
                    Process cm = new Process();
                    cm.StartInfo.FileName = "cmd.exe";
                    cm.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    cm.StartInfo.Arguments = strCmdText;
                    cm.Start();
                    cm.WaitForExit();
                }
              else  if (Ano.Text == "2019" && departamento != "Direccion")
                {
                    strCmdText = "/C attrib -h -s " + " \"" + "G:\\SGC-PROYECTOS-CBR\\SGC\\" + Ano.Text + "\\" + tipo2.Text + "\\" + valor + "\\" + departamento + "\\*" + "\""  + "  /S" + "  /D";
                    Process cm = new Process();
                    cm.StartInfo.FileName = "cmd.exe";
                    cm.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    cm.StartInfo.Arguments = strCmdText;
                    cm.Start();
                    cm.WaitForExit();
                }

                else if (Ano.Text == "2019" && departamento == "Direccion" || Ano.Text == "2019" && departamento == "SuperAdmin")
                {
                    strCmdText = "/C attrib -h -s " + " \"" + "G:\\SGC-PROYECTOS-CBR\\SGC\\" + Ano.Text + "\\" + tipo2.Text + "\\" + valor + "\\*" +" "+ "\" "  + "  /S" + "  /D";
                    Process cm = new Process();
                    cm.StartInfo.FileName = "cmd.exe";
                    cm.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    cm.StartInfo.Arguments = strCmdText;
                    cm.Start();
                    cm.WaitForExit();
                }

                else if (Ano.Text == "2018" && departamento == "Direccion" || Ano.Text == "2018" && departamento == "SuperAdmin")
                {
                    strCmdText = "/C attrib -h -s " + " \"" + "G:\\SGC-PROYECTOS-CBR\\SGC\\" + Ano.Text + "\\" + tipo2.Text + "\\" + valor + "\\*" + " " + "\"" + "  /S" + "  /D";
                    Process cm = new Process();
                    cm.StartInfo.FileName = "cmd.exe";
                    cm.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    cm.StartInfo.Arguments = strCmdText;
                    cm.Start();
                    cm.WaitForExit();
                }
                else if (Ano.Text == "2017" && departamento == "Direccion" || Ano.Text == "2017" && departamento == "SuperAdmin")
                {
                    strCmdText = "/C attrib -h -s " + " \"" + "G:\\SGC-PROYECTOS-CBR\\SGC\\" + Ano.Text + "\\" + tipo2.Text + "\\" + valor + "\\*" + " " + "\""  + "  /S" + "  /D";
                    Process cm = new Process();
                    cm.StartInfo.FileName = "cmd.exe";
                    cm.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    cm.StartInfo.Arguments = strCmdText;
                    cm.Start();
                    cm.WaitForExit();
                }
                else if(departamento == "Direccion" && Ano.Text != "2019" || departamento == "SuperAdmin" && Ano.Text != "2019" || departamento == "Direccion" && Ano.Text != "2018" || departamento == "SuperAdmin" && Ano.Text != "2018" || departamento == "Direccion" && Ano.Text != "2017" || departamento == "SuperAdmin" && Ano.Text != "2017")
                {
                    strCmdText = "/C attrib -h -s " + " \"" + "G:\\SGC-PROYECTOS-CBR\\SGC\\" + Ano.Text + "\\" + tipo2.Text + "\\" + valor + "\\*" + "\"" + "  /S" + "  /D";
                    Process cm = new Process();
                    cm.StartInfo.FileName = "cmd.exe";
                    cm.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    cm.StartInfo.Arguments = strCmdText;
                    cm.Start();
                    cm.WaitForExit();
                }
                else if (departamento != "Direccion" && Ano.Text == "2019" || departamento != "SuperAdmin" && Ano.Text == "2019" || departamento != "Direccion" && Ano.Text == "2018" || departamento != "SuperAdmin" && Ano.Text == "2018" || departamento != "Direccion" && Ano.Text == "2017" || departamento == "SuperAdmin" && Ano.Text == "2017")

                {
                    strCmdText = "/C attrib -h -s " + " \"" + "G:\\SGC-PROYECTOS-CBR\\SGC\\" + Ano.Text + "\\" + tipo2.Text + "\\" + valor + "\\" + departamento +"\\*" + " \"" + "  /S" + "  /D";
                    Process cm = new Process();
                    cm.StartInfo.FileName = "cmd.exe";
                    cm.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    cm.StartInfo.Arguments = strCmdText;
                    cm.Start();
                    cm.WaitForExit();
                    strCmdText = "/C attrib -h -s " + " \"" + "G:\\SGC-PROYECTOS-CBR\\SGC\\" + Ano.Text + "\\" + tipo2.Text + "\\" + valor + "\\" + departamento  + " \"" + "  /S" + "  /D";
                
                    cm.StartInfo.FileName = "cmd.exe";
                    cm.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    cm.StartInfo.Arguments = strCmdText;
                    cm.Start();
                    cm.WaitForExit();

                }
                else 

                {
                    strCmdText = "/C attrib -h -s " + " \"" + "G:\\SGC-PROYECTOS-CBR\\SGC\\" + Ano.Text + "\\" + tipo2.Text + "\\" + valor + "\\"+Nombre+"\\*"+ "\" " +"/S" + " /D";
                    Process cm = new Process();
                    cm.StartInfo.FileName = "cmd.exe";
                    cm.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    cm.StartInfo.Arguments = strCmdText;
                    cm.Start();
                    cm.WaitForExit();
                    strCmdText = "/C attrib -h -s " + " \"" + "G:\\SGC-PROYECTOS-CBR\\SGC\\" + Ano.Text + "\\" + tipo2.Text + "\\" + valor + "\\" + Nombre + "\" " + "/S" + " /D";
                 
                    cm.StartInfo.FileName = "cmd.exe";
                    cm.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    cm.StartInfo.Arguments = strCmdText;
                    cm.Start();
                    cm.WaitForExit();

                }

                conexion.Dispose();
            }
            else { }
        }

        public void contador()
        {
           
           

            numerador = Int32.Parse(Nombre);

            suma = 0;
         
          
            rastreadorproyecto();

            for (k = 1; k <= numerador; k++)
            {
                suma = suma + k;
                if (Ano.Text == "2017" || Ano.Text == "2018" || Ano.Text == "2019") { k = numerador;  suma = numerador; }
                else { }
                Procesocarp genera = new Procesocarp();
      
                genera.total = numerador.ToString();
                genera.contador = k.ToString();
                genera.Show();
                genera.BringToFront();
                genera.WindowState = FormWindowState.Normal;
                Mostrador();
                genera.Close();
             
                // richTextBox1.Text = suma.ToString();


            }
        }

        private void pictureBox11_Click(object sender, EventArgs e)
        {

            string strCmdText;
            consultacarpeta(Proyecto.Text);
            if (tipo2.Text == "Proyectos-Alto-Flujo")
            {

                strCmdText = "G:\\SGC-PROYECTOS-CBR\\SGC\\" + Ano.Text + "\\" + tipo2.Text + "\\" + proyectorastreo + "\\" + Cb_af_com.Text;
               

             
                string strCmdText1;
                strCmdText1 = "/C attrib +h +s  " + "\"" + "G:\\SGC-PROYECTOS-CBR\\SGC\\" + Año2 + "\\" + tipo2.Text + "\\" + proyectorastreo +"\\"+ Cb_af_com.Text+ ""  + "\" " + " /S /D";
                Process cm = new Process();
                cm.StartInfo.FileName = "cmd.exe";
                cm.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                cm.StartInfo.Arguments = strCmdText1;
                cm.Start();
                cm.WaitForExit();
                webBrowser1.Navigate(strCmdText);
            }
            if (tipo2.Text == "Proyectos-Compacta")
            {

                strCmdText = "G:\\SGC-PROYECTOS-CBR\\SGC\\" + Ano.Text + "\\" + tipo2.Text + "\\" + proyectorastreo + "\\" + Cb_af_com.Text;
             
                string strCmdText1;
                strCmdText1 = "/C attrib +h +s  " + "\"" + "G:\\SGC-PROYECTOS-CBR\\SGC\\" + Año2 + "\\" + tipo2.Text + "\\" + proyectorastreo + "\\" + Cb_af_com.Text + "" + "\" " + " /S /D";
                Process cm = new Process();
                cm.StartInfo.FileName = "cmd.exe";
                cm.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                cm.StartInfo.Arguments = strCmdText1;
                cm.Start();
                cm.WaitForExit();
                webBrowser1.Navigate(strCmdText);

            }
            if (tipo2.Text == "Proyectos-Industrial-Potabiliza")
            {

                strCmdText = "G:\\SGC-PROYECTOS-CBR\\SGC\\" + Ano.Text + "\\" + tipo2.Text + "\\" + proyectorastreo + "\\" + indus_po.Text;
               
                string strCmdText1;
                strCmdText1 = "/C attrib +h +s  " + "\"" + "G:\\SGC-PROYECTOS-CBR\\SGC\\" + Año2 + "\\" + tipo2.Text + "\\" + proyectorastreo + "\\" + indus_po.Text  + "\" " + " /S /D";
                Process cm = new Process();
                cm.StartInfo.FileName = "cmd.exe";
                cm.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                cm.StartInfo.Arguments = strCmdText1;
                cm.Start();
                cm.WaitForExit();
                webBrowser1.Navigate(strCmdText);
            }
            if (tipo2.Text == "Proyectos-Inox")
            {

                strCmdText = "G:\\SGC-PROYECTOS-CBR\\SGC\\" + Ano.Text + "\\" + tipo2.Text + "\\" + proyectorastreo + "\\"+cb_inox.Text ;
                string strCmdText1;
                strCmdText1 = "/C attrib +h +s  " + "\"" + "G:\\SGC-PROYECTOS-CBR\\SGC\\" + Año2 + "\\" + tipo2.Text + "\\" + proyectorastreo + "\\" + cb_inox.Text + "\" " + " /S /D";
                Process cm = new Process();
                cm.StartInfo.FileName = "cmd.exe";
                cm.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                cm.StartInfo.Arguments = strCmdText1;
                cm.Start();
                cm.WaitForExit();
                webBrowser1.Navigate(strCmdText);
            }
         
          

        }
        public void Ejecutar(string texto)
        {
            this.folio_ProyectosTableAdapter.ConsultaQ(texto);
        }

        private void pictureBox5_Click_1(object sender, EventArgs e)
        {
           
            Consulta_Proyectos_General hijo = new Consulta_Proyectos_General();
           
            hijo.Show();


        }

        private void pictureBox14_Click(object sender, EventArgs e)
        {
                    this.folio_ProyectosTableAdapter.Consulta(folio_Proyectos._Folio_Proyectos,Nombre1);
        }

        private void label19_Click(object sender, EventArgs e)
        {

        }

        private void Atencionaclientes_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button2_Click_2(object sender, EventArgs e)
        {
            Process.Start(@"G:\SGC\Formatos\DIR_Planeacion Medicion y Revision");
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            Process.Start(@"G:\SGC\Formatos\SGC_Gestion del Conocimiento");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Process.Start(@"G:\SGC\Formatos\SGC_Comunicacion");
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            Process.Start(@"G:\SGC\Formatos\SGC_Auditoria Interna");
        }

        private void button6_Click_1(object sender, EventArgs e)
        {
            Process.Start(@"G:\SGC\Formatos\SGC_Mejora Continua");
        }

        private void button7_Click_1(object sender, EventArgs e)
        {
            Process.Start(@"G:\SGC\Formatos\RIE_Gestion de Riesgos");
        }

        private void button8_Click_1(object sender, EventArgs e)
        {
            Process.Start(@"G:\SGC\Formatos\VTA_Ventas");
        }

        private void button9_Click(object sender, EventArgs e)
        {
            Process.Start(@"G:\SGC\Formatos\CLI_Servicio a clientes");
        }

        private void button17_Click(object sender, EventArgs e)
        {
            Process.Start(@"G:\SGC\Formatos\ING_Ingenieria");
        }

        private void button16_Click(object sender, EventArgs e)
        {
            Process.Start(@"G:\SGC\Formatos\FAB_Fabricacion");
        }

        private void button15_Click(object sender, EventArgs e)
        {
            Process.Start(@"G:\SGC\Formatos\OCI_Contruccion");
        }

        private void button14_Click(object sender, EventArgs e)
        {
            Process.Start(@"G:\SGC\Formatos\EQI_Equipamiento e Instalaciones");
        }

        private void button13_Click(object sender, EventArgs e)
        {
            Process.Start(@"G:\SGC\Formatos\OYM_Operacion y Mantenimiento");
        }

        private void button12_Click(object sender, EventArgs e)
        {
            Process.Start(@"G:\SGC\Formatos\RHU_Recursos Humanos");
        }

        private void button11_Click(object sender, EventArgs e)
        {
            Process.Start(@"G:\SGC\Formatos\FIN_Admintracion Financiera");
        }

        private void button10_Click(object sender, EventArgs e)
        {
            Process.Start(@"G:\SGC\Formatos\INF_Infraestructura");

        }

        private void button18_Click(object sender, EventArgs e)
        {
            Process.Start(@"G:\SGC\Formatos\COM_Compras");
        }

        private void windowsNavegador_MouseMove(object sender, MouseEventArgs e)
        {
      
        }

        private void windowsNavegador_MouseCaptureChanged(object sender, EventArgs e)
        {
            this.folio_ProyectosTableAdapter.Fill(folio_Proyectos._Folio_Proyectos);
        }

        private void panel7_Paint(object sender, PaintEventArgs e)
        {

        }

        private void windowsNavegador_MouseEnter(object sender, EventArgs e)
        {
            this.folio_ProyectosTableAdapter.Fill(folio_Proyectos._Folio_Proyectos);
        }

        private void tableLayoutPanel1_MouseMove(object sender, MouseEventArgs e)
        {///////////////// aca si///
         //
            string valor = DatosGenerales.Name;
            if (valor == null || valor=="") { }
            else
            {

                this.folio_ProyectosTableAdapter.FillBy(folio_Proyectos._Folio_Proyectos, valor);
            }
        }

        private void indus_po_SelectedIndexChanged(object sender, EventArgs e)
        {

        }



        private void label17_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (departamento == "Calidad")
            {
                formato = Calidad.Text;
            }
           else if (departamento == "Compras")
            {
                formato = Compras.Text;
            }
           else if (departamento == "Ventas")
            {
                formato = Ventas.Text;
            }
          else  if (departamento == "Operacion y Mantenimiento")
            {
                formato = OperacionyManteniento.Text;
            }
          else  if (departamento == "Proyectos") 
            {
                formato = Proyecto.Text;
            }
          else  if (departamento == "Atencion a Clientes")
            {
                formato = Atencionaclientes.Text;
            }

            Process.Start(@"G:\SGC\Formatos\"+departamento+"\\"+formato);
        }

        private void pictureBox9_Click_1(object sender, EventArgs e)
        {
            string url = ruta.SelectedValue.ToString(); ;
            webBrowser1.Url = new Uri(url);
            navegador.Text = url;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Process.Start(@"G:\SGC\Formatos\FVT.01 Pedido Interno_5.xlsx");
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {


            rastreadorproyecto();
            string url1;
            url1 = "file:///G:/SGC-PROYECTOS-CBR/SGC/" + Año2 + "/" + tipo3 + "/" + valor;
            string url2 = webBrowser1.Url.ToString();
            if (Proyecto.Text == "")
            {
                string url = "G:/SGC-PROYECTOS-CBR/SGC/";
                webBrowser1.Url = new Uri(url);
                navegador.Text = url;
            }

            else if (url1 != url2)
            {
                string url = "G:/SGC-PROYECTOS-CBR/SGC/";
                webBrowser1.Url = new Uri(url);
                navegador.Text = url;
            }
            else
            {


                DialogResult Resultado;
                Resultado = MessageBox.Show("Está seguro que desea terminar la sesion en el proyecto: " + Proyecto.Text, "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    Loading apertura = new Loading();
                    apertura.Show();
                    apertura.BringToFront();
                    apertura.WindowState = FormWindowState.Normal;
                    rastreadorproyecto();
                    this.registro_AccesosTableAdapter.elimina(Proyecto.Text);



                    webBrowser1.GoBack();




                    string strCmdText;
                    strCmdText = "/C attrib +h +s  " + "\"" + "G:\\SGC-PROYECTOS-CBR\\SGC\\" + Año2 + "\\" + tipo3 + "\\" + valor + "\\*" + "\" " + " /S /D";
                    Nombre1 = "";
                    valor = "";
                    Process cm = new Process();
                    cm.StartInfo.FileName = "cmd.exe";
                    cm.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    cm.StartInfo.Arguments = strCmdText;
                    cm.Start();
                    cm.WaitForExit();
                    apertura.Hide();
                    MessageBox.Show("Sesion cerrada correctamente");

                    this.folio_ProyectosTableAdapter.Consulta(folio_Proyectos._Folio_Proyectos, "");
                    if (webBrowser1.CanGoBack)
                        webBrowser1.GoBack();

                    navegador.Text = webBrowser1.Url.ToString();

                }
                else if (Resultado == DialogResult.No)
                {


                    MessageBox.Show("Operacion Cancelada");
                }
            }


          
        }

        private void button2_Click(object sender, EventArgs e)
        {
           
        }

        private void button3_Click(object sender, EventArgs e)
        {

            MessageBoxManager.OK = "Compacta";
            MessageBoxManager.Cancel = "Alto Flujo";
            MessageBoxManager.Register();


            DialogResult Resultado;
            Resultado = MessageBox.Show("Seleccione el tipo de Cotizacion", "Confirmación", MessageBoxButtons.OKCancel);
            if (Resultado == DialogResult.OK)
            {
                Process.Start(@"G:\SGC\Formatos\FVT.07 Cotización PTAR COMPACTA WEA®_1.docx");


            }
            else if(Resultado == DialogResult.Cancel)
            {
                Process.Start(@"G:\SGC\Formatos\FVT.08 Cotización PTAR ALTO FLUJO WEA®_1.docx");
            }
        }

        private void pictureBox6_Click(object sender, EventArgs e)
        {

            carptape = textBox1.Text;


            if (carptape != null )
            {

                DialogResult Resultado;
                Resultado = MessageBox.Show("Desea Notificar terminacion de lleno:\nCarpeta:" + carptape + "\nProyecto: " + "" + proye_solicita.Text, "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {

               
                            enviamasivo3(Departame.Text);
           
                    //if (depar.Text == "") { MessageBox.Show("No ha ingresado el departamento a notificar"); }
                    //if (depar.Text != "")
                    //{
                    //    enviamasivo(depar.Text);
                    //}
                    MessageBox.Show("Notificacion enviada correctamente");
                }
                else if (Resultado == DialogResult.No)
                {


                    MessageBox.Show("Operacion Cancelada");
                }
            }
            else if (carptape == null)
            {
                MessageBox.Show("No ha seleccinado la carpeta");
            }

            }


        public void enviamasivo3(string Departamento)
        {
            try
            {
                ////////////////////////////////////SE ABRE LA CONEXION/////////////////////////////////////////////////////////////////////////////////
                SqlConnection conexion = new SqlConnection(ObtenerCadena());
                conexion.Open();
                ////////////////////////////////////SE EJECUTA EL COMANDO HACIA LA VISTA Y SE HACE LA VALIDACION//////////////////////////////////////
                SqlCommand cmd = new SqlCommand(
                                                "select " +
                                                "[Nombre], " +
                                                "[Apellido], " +
                                                "[Privilegios], " +
                                                "[Usuario], " +
                                                "[Contraseña], " +
                                                 "[Puesto], " +
                                                 "[Departamento], " +
                                                  "[Email] " +
                                                "from [Login_Departamentos] " +
                                                "where [Departamento]=@Departamento "

                                                , conexion);
                cmd.Parameters.AddWithValue("Departamento", Departamento);



                SqlDataAdapter sda = new SqlDataAdapter(cmd);
                sda.SelectCommand.CommandTimeout = 36000;
                DataTable dt = new DataTable();
                sda.Fill(dt);

                conexion.Dispose();

                if (dt.Rows.Count == 1)
                {
                    Email1 = dt.Rows[0][7].ToString();
                    string _sender = "support@cbr-ingenieria.com.mx";
                    string _password = "Cbrsoporte2020.";
                    SmtpClient client = new SmtpClient("smtp.office365.com");
                    client.Port = 587;
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.UseDefaultCredentials = false;
                    System.Net.NetworkCredential credentials =
                     new System.Net.NetworkCredential(_sender, _password);
                    client.EnableSsl = true;
                    client.Credentials = credentials;
                    MailMessage message = new MailMessage(_sender, Email1);
                 
                    message.Subject = "Actualizacion de Carpeta";
                    message.IsBodyHtml = true;
                    message.AlternateViews.Add(getEmbeddedImage2("//192.168.1.101/Aplicativos/CBR Sistema global/logo_cbr.png"));
                    client.Send(message);
                }

                else if (dt.Rows.Count == 2)
                {
                    Email1 = dt.Rows[0][7].ToString();
                    Email2 = dt.Rows[1][7].ToString();
                    ///////////////////////////////
                    string _sender = "support@cbr-ingenieria.com.mx";
                    string _password = "Cbrsoporte2020.";
                    SmtpClient client = new SmtpClient("smtp.office365.com");
                    client.Port = 587;
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.UseDefaultCredentials = false;
                    System.Net.NetworkCredential credentials =
                     new System.Net.NetworkCredential(_sender, _password);
                    client.EnableSsl = true;
                    client.Credentials = credentials;
                    MailMessage message = new MailMessage(_sender, Email1);
                    message.CC.Add(new MailAddress(Email2));
                   
                    message.Subject = "Actualizacion de  Carpeta";
                    message.IsBodyHtml = true;
                    message.AlternateViews.Add(getEmbeddedImage2("//192.168.1.101/Aplicativos/CBR Sistema global/logo_cbr.png"));
                    client.Send(message);
                }
                else if (dt.Rows.Count == 3)
                {
                    Email1 = dt.Rows[0][7].ToString();
                    Email2 = dt.Rows[1][7].ToString();
                    Email3 = dt.Rows[2][7].ToString();
                    /////////////////////////////////////
                    string _sender = "support@cbr-ingenieria.com.mx";
                    string _password = "Cbrsoporte2020.";
                    SmtpClient client = new SmtpClient("smtp.office365.com");
                    client.Port = 587;
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.UseDefaultCredentials = false;
                    System.Net.NetworkCredential credentials =
                     new System.Net.NetworkCredential(_sender, _password);
                    client.EnableSsl = true;
                    client.Credentials = credentials;
                    MailMessage message = new MailMessage(_sender, Email1);
                    message.CC.Add(new MailAddress(Email2));
                    message.CC.Add(new MailAddress(Email3));
                   
                    message.Subject = "Actualizacion de Carpeta";
                    message.IsBodyHtml = true;
                    message.AlternateViews.Add(getEmbeddedImage2("//192.168.1.101/Aplicativos/CBR Sistema global/logo_cbr.png"));
                    client.Send(message);
                }
                else if (dt.Rows.Count == 4)
                {
                    Email1 = dt.Rows[0][7].ToString();
                    Email2 = dt.Rows[1][7].ToString();
                    Email3 = dt.Rows[2][7].ToString();
                    Email4 = dt.Rows[3][7].ToString();
                    /////////////////////////////////
                    string _sender = "support@cbr-ingenieria.com.mx";
                    string _password = "Cbrsoporte2020.";
                    SmtpClient client = new SmtpClient("smtp.office365.com");
                    client.Port = 587;
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.UseDefaultCredentials = false;
                    System.Net.NetworkCredential credentials =
                     new System.Net.NetworkCredential(_sender, _password);
                    client.EnableSsl = true;
                    client.Credentials = credentials;
                    MailMessage message = new MailMessage(_sender, Email1);
                    message.CC.Add(new MailAddress(Email2));
                    message.CC.Add(new MailAddress(Email3));
                    message.CC.Add(new MailAddress(Email4));
                   
                    message.Subject = "Actualizacion de Carpeta";
                    message.IsBodyHtml = true;
                    message.AlternateViews.Add(getEmbeddedImage2("//192.168.1.101/Aplicativos/CBR Sistema global/logo_cbr.png"));
                    client.Send(message);
                }
                else if (dt.Rows.Count == 5)
                {
                    Email1 = dt.Rows[0][7].ToString();
                    Email2 = dt.Rows[1][7].ToString();
                    Email3 = dt.Rows[2][7].ToString();
                    Email4 = dt.Rows[3][7].ToString();
                    Email5 = dt.Rows[4][7].ToString();
                    string _sender = "support@cbr-ingenieria.com.mx";
                    string _password = "Cbrsoporte2020.";
                    SmtpClient client = new SmtpClient("smtp.office365.com");
                    client.Port = 587;
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.UseDefaultCredentials = false;
                    System.Net.NetworkCredential credentials =
                     new System.Net.NetworkCredential(_sender, _password);
                    client.EnableSsl = true;
                    client.Credentials = credentials;
                    MailMessage message = new MailMessage(_sender, Email1);
                    message.CC.Add(new MailAddress(Email2));
                    message.CC.Add(new MailAddress(Email3));
                    message.CC.Add(new MailAddress(Email4));
                    message.CC.Add(new MailAddress(Email5));
                   
                    message.Subject = "Actualizacion de Carpeta";
                    message.IsBodyHtml = true;
                    message.AlternateViews.Add(getEmbeddedImage2("//192.168.1.101/Aplicativos/CBR Sistema global/logo_cbr.png"));
                    client.Send(message);

                }
                else if (dt.Rows.Count == 6)
                {
                    Email1 = dt.Rows[0][7].ToString();
                    Email2 = dt.Rows[1][7].ToString();
                    Email3 = dt.Rows[2][7].ToString();
                    Email4 = dt.Rows[3][7].ToString();
                    Email5 = dt.Rows[4][7].ToString();
                    Email6 = dt.Rows[5][7].ToString();
                    /////////////////////////////////////
                    string _sender = "support@cbr-ingenieria.com.mx";
                    string _password = "Cbrsoporte2020.";
                    SmtpClient client = new SmtpClient("smtp.office365.com");
                    client.Port = 587;
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.UseDefaultCredentials = false;
                    System.Net.NetworkCredential credentials =
                     new System.Net.NetworkCredential(_sender, _password);
                    client.EnableSsl = true;
                    client.Credentials = credentials;
                    MailMessage message = new MailMessage(_sender, Email1);
                    message.CC.Add(new MailAddress(Email2));
                    message.CC.Add(new MailAddress(Email3));
                    message.CC.Add(new MailAddress(Email4));
                    message.CC.Add(new MailAddress(Email5));
                    message.CC.Add(new MailAddress(Email6));
                    message.Subject = "Actualizacion de Carpeta";
                    message.IsBodyHtml = true;
                    message.AlternateViews.Add(getEmbeddedImage2("//192.168.1.101/Aplicativos/CBR Sistema global/logo_cbr.png"));
                    client.Send(message);
                }
                else if (dt.Rows.Count == 7)
                {
                    Email1 = dt.Rows[0][7].ToString();
                    Email2 = dt.Rows[1][7].ToString();
                    Email3 = dt.Rows[2][7].ToString();
                    Email4 = dt.Rows[3][7].ToString();
                    Email5 = dt.Rows[4][7].ToString();
                    Email6 = dt.Rows[5][7].ToString();
                    Email7 = dt.Rows[6][7].ToString();
                    //////////////////////////////////////
                    string _sender = "support@cbr-ingenieria.com.mx";
                    string _password = "Cbrsoporte2020.";
                    SmtpClient client = new SmtpClient("smtp.office365.com");
                    client.Port = 587;
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.UseDefaultCredentials = false;
                    System.Net.NetworkCredential credentials =
                     new System.Net.NetworkCredential(_sender, _password);
                    client.EnableSsl = true;
                    client.Credentials = credentials;
                    MailMessage message = new MailMessage(_sender, Email1);
                    message.CC.Add(new MailAddress(Email2));
                    message.CC.Add(new MailAddress(Email3));
                    message.CC.Add(new MailAddress(Email4));
                    message.CC.Add(new MailAddress(Email5));
                    message.CC.Add(new MailAddress(Email6));
                 
                    message.Subject = "Actualizacion de Carpeta";
                    message.IsBodyHtml = true;
                    message.AlternateViews.Add(getEmbeddedImage2("//192.168.1.101/Aplicativos/CBR Sistema global/logo_cbr.png"));
                    client.Send(message);
                }
                else if (dt.Rows.Count == 8)
                {
                    Email1 = dt.Rows[0][7].ToString();
                    Email2 = dt.Rows[1][7].ToString();
                    Email3 = dt.Rows[2][7].ToString();
                    Email4 = dt.Rows[3][7].ToString();
                    Email5 = dt.Rows[4][7].ToString();
                    Email6 = dt.Rows[5][7].ToString();
                    Email7 = dt.Rows[6][7].ToString();
                    Email8 = dt.Rows[7][7].ToString();
                    ///////////////////////////////////////////////////////
                    string _sender = "support@cbr-ingenieria.com.mx";
                    string _password = "Cbrsoporte2020.";
                    SmtpClient client = new SmtpClient("smtp.office365.com");
                    client.Port = 587;
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.UseDefaultCredentials = false;
                    System.Net.NetworkCredential credentials =
                     new System.Net.NetworkCredential(_sender, _password);
                    client.EnableSsl = true;
                    client.Credentials = credentials;
                    MailMessage message = new MailMessage(_sender, Email1);
                    message.CC.Add(new MailAddress(Email2));
                    message.CC.Add(new MailAddress(Email3));
                    message.CC.Add(new MailAddress(Email4));
                    message.CC.Add(new MailAddress(Email5));
                    message.CC.Add(new MailAddress(Email6));
                    message.CC.Add(new MailAddress(Email7));
                    message.CC.Add(new MailAddress(Email8));
                    message.Subject = "Actualizacion de Carpeta";
                    message.IsBodyHtml = true;
                    message.AlternateViews.Add(getEmbeddedImage2("//192.168.1.101/Aplicativos/CBR Sistema global/logo_cbr.png"));
                    client.Send(message);
                }
                else if (dt.Rows.Count == 9)
                {
                    Email1 = dt.Rows[0][7].ToString();
                    Email2 = dt.Rows[1][7].ToString();
                    Email3 = dt.Rows[2][7].ToString();
                    Email4 = dt.Rows[3][7].ToString();
                    Email5 = dt.Rows[4][7].ToString();
                    Email6 = dt.Rows[5][7].ToString();
                    Email7 = dt.Rows[6][7].ToString();
                    Email8 = dt.Rows[7][7].ToString();
                    Email9 = dt.Rows[8][7].ToString();
                    ///////////////////////////////////////////
                    string _sender = "support@cbr-ingenieria.com.mx";
                    string _password = "Cbrsoporte2020.";
                    SmtpClient client = new SmtpClient("smtp.office365.com");
                    client.Port = 587;
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.UseDefaultCredentials = false;
                    System.Net.NetworkCredential credentials =
                     new System.Net.NetworkCredential(_sender, _password);
                    client.EnableSsl = true;
                    client.Credentials = credentials;
                    MailMessage message = new MailMessage(_sender, Email1);
                    message.CC.Add(new MailAddress(Email2));
                    message.CC.Add(new MailAddress(Email3));
                    message.CC.Add(new MailAddress(Email4));
                    message.CC.Add(new MailAddress(Email5));
                    message.CC.Add(new MailAddress(Email6));
                    message.CC.Add(new MailAddress(Email7));
                    message.CC.Add(new MailAddress(Email8));
                    message.CC.Add(new MailAddress(Email9));
                 
                    message.Subject = "Actualizacion de Carpeta";
                    message.IsBodyHtml = true;
                    message.AlternateViews.Add(getEmbeddedImage2("//192.168.1.101/Aplicativos/CBR Sistema global/logo_cbr.png"));
                    client.Send(message);
                }
                else if (dt.Rows.Count == 10)
                {
                    Email1 = dt.Rows[0][7].ToString();
                    Email2 = dt.Rows[1][7].ToString();
                    Email3 = dt.Rows[2][7].ToString();
                    Email4 = dt.Rows[3][7].ToString();
                    Email5 = dt.Rows[4][7].ToString();
                    Email6 = dt.Rows[5][7].ToString();
                    Email7 = dt.Rows[6][7].ToString();
                    Email8 = dt.Rows[7][7].ToString();
                    Email9 = dt.Rows[8][7].ToString();
                    Email10 = dt.Rows[9][7].ToString();
                    ////////////////////////////////////////////
                    string _sender = "support@cbr-ingenieria.com.mx";
                    string _password = "Cbrsoporte2020.";
                    SmtpClient client = new SmtpClient("smtp.office365.com");
                    client.Port = 587;
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.UseDefaultCredentials = false;
                    System.Net.NetworkCredential credentials =
                     new System.Net.NetworkCredential(_sender, _password);
                    client.EnableSsl = true;
                    client.Credentials = credentials;
                    MailMessage message = new MailMessage(_sender, Email1);
                    message.CC.Add(new MailAddress(Email2));
                    message.CC.Add(new MailAddress(Email3));
                    message.CC.Add(new MailAddress(Email4));
                    message.CC.Add(new MailAddress(Email5));
                    message.CC.Add(new MailAddress(Email6));
                    message.CC.Add(new MailAddress(Email7));
                    message.CC.Add(new MailAddress(Email8));
                    message.CC.Add(new MailAddress(Email9));
                    message.CC.Add(new MailAddress(Email10));
                  
                    message.Subject = "Actualizacion de Carpeta";
                    message.IsBodyHtml = true;
                    message.AlternateViews.Add(getEmbeddedImage2("//192.168.1.101/Aplicativos/CBR Sistema global/logo_cbr.png"));
                    client.Send(message);

                }
                else if (dt.Rows.Count > 10)
                {
                    MessageBox.Show("Numero de destinatarios maximo alcanzado por departamento contacte con soporte");
                }


                ////////////////////////////////////SI EL USUARIO, CONTRASENA O PLANTA ESTA MAL//////////////////////////////////////////////////////////////////////////
                else
                {
                    MessageBox.Show("Sin resultado");
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            finally
            {

            }
        }
       
        public void enviamasivo2(string Departamento)
        {
            try
            {
                ////////////////////////////////////SE ABRE LA CONEXION/////////////////////////////////////////////////////////////////////////////////
                SqlConnection conexion = new SqlConnection(ObtenerCadena());
                conexion.Open();
                ////////////////////////////////////SE EJECUTA EL COMANDO HACIA LA VISTA Y SE HACE LA VALIDACION//////////////////////////////////////
                SqlCommand cmd = new SqlCommand(
                                                "select " +
                                                "[Nombre], " +
                                                "[Apellido], " +
                                                "[Privilegios], " +
                                                "[Usuario], " +
                                                "[Contraseña], " +
                                                 "[Puesto], " +
                                                 "[Departamento], " +
                                                  "[Email] " +
                                                "from [Login_Departamentos] " +
                                                "where [Departamento]=@Departamento "

                                                , conexion);
                cmd.Parameters.AddWithValue("Departamento", Depa.Text);



                SqlDataAdapter sda = new SqlDataAdapter(cmd);
                sda.SelectCommand.CommandTimeout = 36000;
                DataTable dt = new DataTable();
                sda.Fill(dt);

                conexion.Dispose();

                if (dt.Rows.Count == 1)
                {
                    Email1 = dt.Rows[0][7].ToString();
                    string _sender = "support@cbr-ingenieria.com.mx";
                    string _password = "Cbrsoporte2020.";
                    SmtpClient client = new SmtpClient("smtp.office365.com");
                    client.Port = 587;
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.UseDefaultCredentials = false;
                    System.Net.NetworkCredential credentials =
                     new System.Net.NetworkCredential(_sender, _password);
                    client.EnableSsl = true;
                    client.Credentials = credentials;
                    MailMessage message = new MailMessage(_sender, Email1);
                    message.CC.Add(new MailAddress("sandra.cabrera@cbr-ingenieria.com.mx"));
                    message.Subject = "Solicitud de actualizacion de Carpeta";
                    message.IsBodyHtml = true;
                    message.AlternateViews.Add(getEmbeddedImage("//192.168.1.101/Aplicativos/CBR Sistema global/logo_cbr.png"));
                    client.Send(message);
                }

                else if (dt.Rows.Count == 2)
                {
                    Email1 = dt.Rows[0][7].ToString();
                    Email2 = dt.Rows[1][7].ToString();
                    ///////////////////////////////
                    string _sender = "support@cbr-ingenieria.com.mx";
                    string _password = "Cbrsoporte2020.";
                    SmtpClient client = new SmtpClient("smtp.office365.com");
                    client.Port = 587;
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.UseDefaultCredentials = false;
                    System.Net.NetworkCredential credentials =
                     new System.Net.NetworkCredential(_sender, _password);
                    client.EnableSsl = true;
                    client.Credentials = credentials;
                    MailMessage message = new MailMessage(_sender, Email1);
                    message.CC.Add(new MailAddress(Email2));
                    message.CC.Add(new MailAddress("sandra.cabrera@cbr-ingenieria.com.mx"));
                    message.Subject = "Solicitud de actualizacion de Carpeta";
                    message.IsBodyHtml = true;
                    message.AlternateViews.Add(getEmbeddedImage("//192.168.1.101/Aplicativos/CBR Sistema global/logo_cbr.png"));
                    client.Send(message);
                }
                else if (dt.Rows.Count == 3)
                {
                    Email1 = dt.Rows[0][7].ToString();
                    Email2 = dt.Rows[1][7].ToString();
                    Email3 = dt.Rows[2][7].ToString();
                    /////////////////////////////////////
                    string _sender = "support@cbr-ingenieria.com.mx";
                    string _password = "Cbrsoporte2020.";
                    SmtpClient client = new SmtpClient("smtp.office365.com");
                    client.Port = 587;
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.UseDefaultCredentials = false;
                    System.Net.NetworkCredential credentials =
                     new System.Net.NetworkCredential(_sender, _password);
                    client.EnableSsl = true;
                    client.Credentials = credentials;
                    MailMessage message = new MailMessage(_sender, Email1);
                    message.CC.Add(new MailAddress(Email2));
                    message.CC.Add(new MailAddress(Email3));
                    message.CC.Add(new MailAddress("sandra.cabrera@cbr-ingenieria.com.mx"));
                    message.Subject = "Solicitud de actualizacion de Carpeta";
                    message.IsBodyHtml = true;
                    message.AlternateViews.Add(getEmbeddedImage("//192.168.1.101/Aplicativos/CBR Sistema global/logo_cbr.png"));
                    client.Send(message);
                }
                else if (dt.Rows.Count == 4)
                {
                    Email1 = dt.Rows[0][7].ToString();
                    Email2 = dt.Rows[1][7].ToString();
                    Email3 = dt.Rows[2][7].ToString();
                    Email4 = dt.Rows[3][7].ToString();
                    /////////////////////////////////
                    string _sender = "support@cbr-ingenieria.com.mx";
                    string _password = "Cbrsoporte2020.";
                    SmtpClient client = new SmtpClient("smtp.office365.com");
                    client.Port = 587;
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.UseDefaultCredentials = false;
                    System.Net.NetworkCredential credentials =
                     new System.Net.NetworkCredential(_sender, _password);
                    client.EnableSsl = true;
                    client.Credentials = credentials;
                    MailMessage message = new MailMessage(_sender, Email1);
                    message.CC.Add(new MailAddress(Email2));
                    message.CC.Add(new MailAddress(Email3));
                    message.CC.Add(new MailAddress(Email4));
                    message.CC.Add(new MailAddress("sandra.cabrera@cbr-ingenieria.com.mx"));
                    message.Subject = "Solicitud de actualizacion de Carpeta";
                    message.IsBodyHtml = true;
                    message.AlternateViews.Add(getEmbeddedImage("//192.168.1.101/Aplicativos/CBR Sistema global/logo_cbr.png"));
                    client.Send(message);
                }
                else if (dt.Rows.Count == 5)
                {
                    Email1 = dt.Rows[0][7].ToString();
                    Email2 = dt.Rows[1][7].ToString();
                    Email3 = dt.Rows[2][7].ToString();
                    Email4 = dt.Rows[3][7].ToString();
                    Email5 = dt.Rows[4][7].ToString();
                    string _sender = "support@cbr-ingenieria.com.mx";
                    string _password = "Cbrsoporte2020.";
                    SmtpClient client = new SmtpClient("smtp.office365.com");
                    client.Port = 587;
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.UseDefaultCredentials = false;
                    System.Net.NetworkCredential credentials =
                     new System.Net.NetworkCredential(_sender, _password);
                    client.EnableSsl = true;
                    client.Credentials = credentials;
                    MailMessage message = new MailMessage(_sender, Email1);
                    message.CC.Add(new MailAddress(Email2));
                    message.CC.Add(new MailAddress(Email3));
                    message.CC.Add(new MailAddress(Email4));
                    message.CC.Add(new MailAddress(Email5));
                    message.CC.Add(new MailAddress("sandra.cabrera@cbr-ingenieria.com.mx"));
                    message.Subject = "Solicitud de actualizacion de Carpeta";
                    message.IsBodyHtml = true;
                    message.AlternateViews.Add(getEmbeddedImage("//192.168.1.101/Aplicativos/CBR Sistema global/logo_cbr.png"));
                    client.Send(message);

                }
                else if (dt.Rows.Count == 6)
                {
                    Email1 = dt.Rows[0][7].ToString();
                    Email2 = dt.Rows[1][7].ToString();
                    Email3 = dt.Rows[2][7].ToString();
                    Email4 = dt.Rows[3][7].ToString();
                    Email5 = dt.Rows[4][7].ToString();
                    Email6 = dt.Rows[5][7].ToString();
                    /////////////////////////////////////
                    string _sender = "support@cbr-ingenieria.com.mx";
                    string _password = "Cbrsoporte2020.";
                    SmtpClient client = new SmtpClient("smtp.office365.com");
                    client.Port = 587;
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.UseDefaultCredentials = false;
                    System.Net.NetworkCredential credentials =
                     new System.Net.NetworkCredential(_sender, _password);
                    client.EnableSsl = true;
                    client.Credentials = credentials;
                    MailMessage message = new MailMessage(_sender, Email1);
                    message.CC.Add(new MailAddress(Email2));
                    message.CC.Add(new MailAddress(Email3));
                    message.CC.Add(new MailAddress(Email4));
                    message.CC.Add(new MailAddress(Email5));
                    message.CC.Add(new MailAddress(Email6));
                    message.Subject = "Solicitud de actualizacion de Carpeta";
                    message.IsBodyHtml = true;
                    message.AlternateViews.Add(getEmbeddedImage("//192.168.1.101/Aplicativos/CBR Sistema global/logo_cbr.png"));
                    client.Send(message);
                }
                else if (dt.Rows.Count == 7)
                {
                    Email1 = dt.Rows[0][7].ToString();
                    Email2 = dt.Rows[1][7].ToString();
                    Email3 = dt.Rows[2][7].ToString();
                    Email4 = dt.Rows[3][7].ToString();
                    Email5 = dt.Rows[4][7].ToString();
                    Email6 = dt.Rows[5][7].ToString();
                    Email7 = dt.Rows[6][7].ToString();
                    //////////////////////////////////////
                    string _sender = "support@cbr-ingenieria.com.mx";
                    string _password = "Cbrsoporte2020.";
                    SmtpClient client = new SmtpClient("smtp.office365.com");
                    client.Port = 587;
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.UseDefaultCredentials = false;
                    System.Net.NetworkCredential credentials =
                     new System.Net.NetworkCredential(_sender, _password);
                    client.EnableSsl = true;
                    client.Credentials = credentials;
                    MailMessage message = new MailMessage(_sender, Email1);
                    message.CC.Add(new MailAddress(Email2));
                    message.CC.Add(new MailAddress(Email3));
                    message.CC.Add(new MailAddress(Email4));
                    message.CC.Add(new MailAddress(Email5));
                    message.CC.Add(new MailAddress(Email6));
                    message.CC.Add(new MailAddress("sandra.cabrera@cbr-ingenieria.com.mx"));
                    message.Subject = "Solicitud de actualizacion de Carpeta";
                    message.IsBodyHtml = true;
                    message.AlternateViews.Add(getEmbeddedImage("//192.168.1.101/Aplicativos/CBR Sistema global/logo_cbr.png"));
                    client.Send(message);
                }
                else if (dt.Rows.Count == 8)
                {
                    Email1 = dt.Rows[0][7].ToString();
                    Email2 = dt.Rows[1][7].ToString();
                    Email3 = dt.Rows[2][7].ToString();
                    Email4 = dt.Rows[3][7].ToString();
                    Email5 = dt.Rows[4][7].ToString();
                    Email6 = dt.Rows[5][7].ToString();
                    Email7 = dt.Rows[6][7].ToString();
                    Email8 = dt.Rows[7][7].ToString();
                    ///////////////////////////////////////////////////////
                    string _sender = "support@cbr-ingenieria.com.mx";
                    string _password = "Cbrsoporte2020.";
                    SmtpClient client = new SmtpClient("smtp.office365.com");
                    client.Port = 587;
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.UseDefaultCredentials = false;
                    System.Net.NetworkCredential credentials =
                     new System.Net.NetworkCredential(_sender, _password);
                    client.EnableSsl = true;
                    client.Credentials = credentials;
                    MailMessage message = new MailMessage(_sender, Email1);
                    message.CC.Add(new MailAddress(Email2));
                    message.CC.Add(new MailAddress(Email3));
                    message.CC.Add(new MailAddress(Email4));
                    message.CC.Add(new MailAddress(Email5));
                    message.CC.Add(new MailAddress(Email6));
                    message.CC.Add(new MailAddress(Email7));
                    message.CC.Add(new MailAddress(Email8));
                    message.Subject = "Solicitud de actualizacion de Carpeta";
                    message.IsBodyHtml = true;
                    message.AlternateViews.Add(getEmbeddedImage("//192.168.1.101/Aplicativos/CBR Sistema global/logo_cbr.png"));
                    client.Send(message);
                }
                else if (dt.Rows.Count == 9)
                {
                    Email1 = dt.Rows[0][7].ToString();
                    Email2 = dt.Rows[1][7].ToString();
                    Email3 = dt.Rows[2][7].ToString();
                    Email4 = dt.Rows[3][7].ToString();
                    Email5 = dt.Rows[4][7].ToString();
                    Email6 = dt.Rows[5][7].ToString();
                    Email7 = dt.Rows[6][7].ToString();
                    Email8 = dt.Rows[7][7].ToString();
                    Email9 = dt.Rows[8][7].ToString();
                    ///////////////////////////////////////////
                    string _sender = "support@cbr-ingenieria.com.mx";
                    string _password = "Cbrsoporte2020.";
                    SmtpClient client = new SmtpClient("smtp.office365.com");
                    client.Port = 587;
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.UseDefaultCredentials = false;
                    System.Net.NetworkCredential credentials =
                     new System.Net.NetworkCredential(_sender, _password);
                    client.EnableSsl = true;
                    client.Credentials = credentials;
                    MailMessage message = new MailMessage(_sender, Email1);
                    message.CC.Add(new MailAddress(Email2));
                    message.CC.Add(new MailAddress(Email3));
                    message.CC.Add(new MailAddress(Email4));
                    message.CC.Add(new MailAddress(Email5));
                    message.CC.Add(new MailAddress(Email6));
                    message.CC.Add(new MailAddress(Email7));
                    message.CC.Add(new MailAddress(Email8));
                    message.CC.Add(new MailAddress(Email9));
                    message.CC.Add(new MailAddress("sandra.cabrera@cbr-ingenieria.com.mx"));
                    message.Subject = "Solicitud de actualizacion de Carpeta";
                    message.IsBodyHtml = true;
                    message.AlternateViews.Add(getEmbeddedImage("//192.168.1.101/Aplicativos/CBR Sistema global/logo_cbr.png"));
                    client.Send(message);
                }
                else if (dt.Rows.Count == 10)
                {
                    Email1 = dt.Rows[0][7].ToString();
                    Email2 = dt.Rows[1][7].ToString();
                    Email3 = dt.Rows[2][7].ToString();
                    Email4 = dt.Rows[3][7].ToString();
                    Email5 = dt.Rows[4][7].ToString();
                    Email6 = dt.Rows[5][7].ToString();
                    Email7 = dt.Rows[6][7].ToString();
                    Email8 = dt.Rows[7][7].ToString();
                    Email9 = dt.Rows[8][7].ToString();
                    Email10 = dt.Rows[9][7].ToString();
                    ////////////////////////////////////////////
                    string _sender = "support@cbr-ingenieria.com.mx";
                    string _password = "Cbrsoporte2020.";
                    SmtpClient client = new SmtpClient("smtp.office365.com");
                    client.Port = 587;
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.UseDefaultCredentials = false;
                    System.Net.NetworkCredential credentials =
                     new System.Net.NetworkCredential(_sender, _password);
                    client.EnableSsl = true;
                    client.Credentials = credentials;
                    MailMessage message = new MailMessage(_sender, Email1);
                    message.CC.Add(new MailAddress(Email2));
                    message.CC.Add(new MailAddress(Email3));
                    message.CC.Add(new MailAddress(Email4));
                    message.CC.Add(new MailAddress(Email5));
                    message.CC.Add(new MailAddress(Email6));
                    message.CC.Add(new MailAddress(Email7));
                    message.CC.Add(new MailAddress(Email8));
                    message.CC.Add(new MailAddress(Email9));
                    message.CC.Add(new MailAddress(Email10));
                    message.CC.Add(new MailAddress("sandra.cabrera@cbr-ingenieria.com.mx"));
                    message.Subject = "Solicitud de actualizacion de Carpeta";
                    message.IsBodyHtml = true;
                    message.AlternateViews.Add(getEmbeddedImage("//192.168.1.101/Aplicativos/CBR Sistema global/logo_cbr.png"));
                    client.Send(message);

                }
                else if (dt.Rows.Count > 10)
                {
                    MessageBox.Show("Numero de destinatarios maximo alcanzado por departamento contacte con soporte");
                }


                ////////////////////////////////////SI EL USUARIO, CONTRASENA O PLANTA ESTA MAL//////////////////////////////////////////////////////////////////////////
                else
                {
                    MessageBox.Show("Sin resultado");
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            finally
            {

            }
        }
        private AlternateView getEmbeddedImage(String filePath)
        {



            // below line was corrected to include the mediatype so it displays in all 
            // mail clients. previous solution only displays in Gmail the inline images 
            LinkedResource res = new LinkedResource(filePath, MediaTypeNames.Image.Jpeg);
            res.ContentId = Guid.NewGuid().ToString();
            String Body = "<DIV><H5>Solicitud de actualizado la carpeta</HEAD></DIV></H>"+carpeta.Text +" "+ "Del proyecto:"+ proye_solicita.Text;
            Body += "<DIV>Ingrese ala plataforma para actualizar estado de carpeta indicada</DIV>";




            string htmlBody = "" + Body + "" + @"<img src='cid:" + res.ContentId + @"'/> ";
            AlternateView alternateView = AlternateView.CreateAlternateViewFromString(htmlBody,
             null, MediaTypeNames.Text.Html);
            alternateView.LinkedResources.Add(res);
            return alternateView;
        }


        private AlternateView getEmbeddedImage2(String filePath)
        {



            // below line was corrected to include the mediatype so it displays in all 
            // mail clients. previous solution only displays in Gmail the inline images 
            LinkedResource res = new LinkedResource(filePath, MediaTypeNames.Image.Jpeg);
            res.ContentId = Guid.NewGuid().ToString();
            String Body = "<DIV><H5>Se ha concluido el llenado de  la carpeta:</HEAD></DIV></H>" + carptape+" "+"Del proyecto:"+ proye_solicita.Text;
            Body += "<DIV>Ingrese ala plataforma para confirmar sus accesos</DIV>";




            string htmlBody = "" + Body + "" + @"<img src='cid:" + res.ContentId + @"'/> ";
            AlternateView alternateView = AlternateView.CreateAlternateViewFromString(htmlBody,
             null, MediaTypeNames.Text.Html);
            alternateView.LinkedResources.Add(res);
            return alternateView;
        }

        private AlternateView getEmbeddedImage3(String filePath)
        {



            // below line was corrected to include the mediatype so it displays in all 
            // mail clients. previous solution only displays in Gmail the inline images 
            LinkedResource res = new LinkedResource(filePath, MediaTypeNames.Image.Jpeg);
            res.ContentId = Guid.NewGuid().ToString();
            String Body = "<DIV><H5>Solicitud de Autorizacion de Pedido:</DIV></H>";
            Body += "<DIV>Se ha solicitado la autorización del pedido:  </DIV>"+PedidoA.Text +"<DIV> Ingrese ala plataforma para revisar el pedido y proceder con la autorizacion</DIV>";




            string htmlBody = "" + Body + "" + @"<img src='cid:" + res.ContentId + @"'/> ";
            AlternateView alternateView = AlternateView.CreateAlternateViewFromString(htmlBody,
             null, MediaTypeNames.Text.Html);
            alternateView.LinkedResources.Add(res);
            return alternateView;
        }


     //private void pictureBox5_Click(object sender, EventArgs e)
     //   {

     //       if (Tipo.Text == "Alto flujo")
     //       {

     //           carptape = AF.Text;
     //       }
     //       if (Tipo.Text == "Compacta")
     //       {

     //           carptape = CP.Text;
     //       }
     //       if (Tipo.Text == "Industrial Potabiliza")
     //       {

     //           carptape = IND.Text;
     //       }
     //       if (Tipo.Text == "Inox")
     //       {

     //           carptape = INOX.Text;
     //       }
     //       if (Tipo.Text == "")
     //       {

     //           MessageBox.Show("No ha seleccionado el tipo de proyecto");
     //       }
     //       if (Tipo.Text != "")
     //       {

     //           if (carptape != "")
     //           {
     //               DialogResult Resultado;
     //               Resultado = MessageBox.Show("Desea terminar el  lleno de la carpeta:" + carptape + "del proyecto: " + "" + proye_solicita.Text, "Confirmación", MessageBoxButtons.YesNo);
     //               if (Resultado == DialogResult.Yes)
     //               {
     //                   consultacarpeta(proye_solicita.Text);

     //                   //////////////////////// Crear carpeta Temporal ///////////////////////////////////////////////////
     //                   string temporal;
     //                   temporal = "/C MD G:\\Sistema\\Temporales\\Temporal\\";
     //                   string temporalbackup;
     //                   string carpetaf = "" + proye_solicita.SelectedValue.ToString() + "-" + carptape + "";
     //                   temporalbackup = "/C MD "+"\""+"G:\\Sistema\\Temporales-Backup\\" + carpetaf+"\"";
     //                   string carpetaf2 = proye_solicita.Text + "-" + carptape;
     //                   //////////////////////////////////////////////////////////////////////////////////////////////////

     //                   ////////////////////////// Copia archivos a Temporal ///////////////////////////////////////////////////
     //                   string copiatemporal = Environment.CurrentDirectory;
     //                   copiatemporal = @"/C robocopy" + " " + " \"G:\\SGC-PROYECTOS-CBR\\SGC\\" + anorastreo + "\\" + carpetarastreo + "\\" + proyectorastreo + "\\" + carptape + "\" " + " G:\\Sistema\\Temporales\\Temporal"+ "  " + "/E";

     //                   ////////////////////////// Copia archivos a Temporal backup///////////////////////////////////////////////////
     //                   string copiatemporalbackup = Environment.CurrentDirectory;
     //                   copiatemporalbackup = @"/C robocopy" + " " + " \"G:\\SGC-PROYECTOS-CBR\\SGC\\" + anorastreo + "\\" + carpetarastreo + "\\" + proyectorastreo + "\\" + carptape + "\" " +"\""+ "G:\\Sistema\\Temporales-Backup\\" + carpetaf + "\""+"  "+"/E";

     //                   //////////////////////////////////////////////////////////////////////////////////////////////////////

     //                   //////////////////////////// elimina carpeta con permiso previo ///////////////////////////////////////////////////
     //                   string elimina;
     //                   elimina = "/C rd" + " " + " \"G:\\SGC-PROYECTOS-CBR\\SGC\\" + anorastreo + "\\" + carpetarastreo + "\\" + proyectorastreo + "\\" + carptape + "\" /S /Q ";

     //                   ////////////////////////////////////////////////////////////////////////////////////////////////////

     //                   /////////////////////////crea proceso de rastreo de proyecto //////////////////////////////////////
     //                   string nuevoperm = "";
     //                   if (carpetarastreo == "Proyectos-Compacta") { nuevoperm = "CodificadoresLideres-Af-Comp-pos"; }
     //                   if (carpetarastreo == "Proyectos-Alto-Flujo") { nuevoperm = "CodificadoresLideres-Af-Comp-pos"; }
     //                   if (carpetarastreo == "Proyectos-Industrial-Potabiliza") { nuevoperm = "CodificadoresLideres-Indus-Pot-pos"; }
     //                   if (carpetarastreo == "Proyectos-Inox") { nuevoperm = "Proyecto-INOX-pos"; codifica = ""; }

     //                   //////////////////////////// obtiene carpeta con nuevos permisos ///////////////////////////////////////////////////
     //                   string recrea;
     //                   recrea = @"/C  robocopy" + " " + "G:\\SGC\\" + anorastreo + "\\" + nuevoperm + "\\" + codifica + " " + "G:\\SGC-PROYECTOS-CBR\\SGC\\" + anorastreo + "\\" + carpetarastreo + "\\" + proyectorastreo + "" + " " + "/E /SEC /R:0";
     //                   // recrea = @"/C robocopy" + " " + "B:\\SGC\\" + anorastreo + "\\" + nuevoperm + " " + " B:\\Proyectos-CBR\\SGC\\" + anorastreo + "\\" + carpetarastreo + "\\" + proye.Text + "" + " " + "/E /SEC";

     //                   //////////////////////////////////////////////////////////////////////////////////////////////////
     //                   ////////////////////////// obtiene archivos a nueva carpeta con permisos ///////////////////////////////////////////////////
     //                   string copia;

     //                   //recrea = @"/C robocopy" + " " + "B:\\SGC\\" + anorastreo + "\\" + nuevoperm  + " " + "B:\\Proyectos-CBR\\SGC\\"+anorastreo+"\\"+carpetarastreo+"\\"+proye.Text+"\\  /E / SEC";
     //                   copia = @"/C xcopy " + " " + "\""+"G:\\Sistema\\Temporales\\Temporal" + "\"" +"  " +"\""+ "G:\\SGC-PROYECTOS-CBR\\SGC\\" + anorastreo + "\\" + carpetarastreo + "\\" + proyectorastreo + "\\" + carptape + "\""+ "" + " " + "/S /Q  /i";

     //                   //////////////////////////////////////////////////////////////////////////////////////////////////
     //                   //////////////////////////// elimina carpeta temporal ///////////////////////////////////////////////////
     //                   string eliminatem;
     //                   eliminatem = "/C rd" + " " + " \"G:\\Sistema\\Temporales\\Temporal" + "\" /S /Q";




     //                   this.migra_permisosTableAdapter.codificador(usuario,apellido,carpetarastreo, carptape, proyectorastreo, temporal, temporalbackup, copiatemporal, copiatemporalbackup, elimina, recrea, copia, eliminatem);
     //                   //Process crea1 = System.Diagnostics.Process.Start("CMD.exe", temporal);
     //                   //crea1.WaitForExit();

     //                   //Thread.Sleep(1000);
     //                   //Process crea = System.Diagnostics.Process.Start("CMD.exe", temporalbackup);
     //                   //crea.WaitForExit();
     //                   //Thread.Sleep(1000);
     //                   //Process copia1 = System.Diagnostics.Process.Start("CMD.exe", copiatemporal);
     //                   //copia1.WaitForExit();
     //                   //Thread.Sleep(1000);
     //                   //Process copia1back = System.Diagnostics.Process.Start("CMD.exe", copiatemporalbackup);
     //                   //copia1back.WaitForExit();
     //                   //Thread.Sleep(1000);
     //                   //Process elimin = System.Diagnostics.Process.Start("CMD.exe", elimina);
     //                   //elimin.WaitForExit();
     //                   //Thread.Sleep(2000);
     //                   //Process recrear = System.Diagnostics.Process.Start("CMD.exe", recrea);
     //                   //recrear.WaitForExit();
     //                   //Thread.Sleep(1000);
     //                   //Process cop = System.Diagnostics.Process.Start("CMD.exe", copia);
     //                   //cop.WaitForExit();
     //                   //Thread.Sleep(1000);
     //                   //Process elim = System.Diagnostics.Process.Start("CMD.exe", eliminatem);
     //                   //elim.WaitForExit();
     //                   ////Process netStat = new Process();
     //                   ////Process elim = System.Diagnostics.Process.Start(@"C:\Windows\system32\cmd.exe", @"/c psexec \\192.168.1.203 -u 192.168.1.48203\Administrador -p asdasd -s cmd.exe");

     //                   //elim.WaitForExit();
     //                   MessageBox.Show("Actualizacion subida espere correos de confirmacion");





     //               }
     //               else if (Resultado == DialogResult.No)
     //               {
 

     //               }
     //           }
     //           else if (carptape == "")
     //           {
     //               MessageBox.Show("No ha seleccionado la carpeta a actualizar");
     //           }


     //       }
     //   }






      

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        public static string ObtenerCadena()
        {
            return Settings.Default.CBR_IngenieriaConnectionString;/*Este codigo obtiene los datos de la cadena de conexion declarada en los settings de la aplicacion*/

        }
        public void consultacarpeta( string proyecto)
        {


            try
            {

                ////////////////////////////////////SE ABRE LA CONEXION/////////////////////////////////////////////////////////////////////////////////
                SqlConnection conexion = new SqlConnection(ObtenerCadena());

                conexion.Open();

                ////////////////////////////////////SE EJECUTA EL COMANDO HACIA LA VISTA Y SE HACE LA VALIDACION//////////////////////////////////////
                SqlCommand cmd = new SqlCommand(
                                                "select " +
                                                "[ID], " +
                                                "[Norma], " +
                                                "[Tipo], " +
                                                "[Capacidad], " +
                                                 "[Fecha], " +
                                                  "[Folio], " +
                                                    "[Nombre], " +
                                                         "[Ano], " +
                                                "[tipo2]," +
                                                "[Nombre2],"+
                                                "[codificador] " +

                                                "from [Folio_Proyectos] " +
                                                "where [Nombre2]=@Nombre"
                                              
                                                , conexion);
                cmd.Parameters.AddWithValue("Nombre", proyecto);
              

                SqlDataAdapter sda = new SqlDataAdapter(cmd);
                sda.SelectCommand.CommandTimeout = 36000;
                DataTable dt = new DataTable();
                sda.Fill(dt);



                conexion.Dispose();


                if (dt.Rows.Count > 0)
                {

                    //MessageBox.Show("Si tiene acceso a esta planta");
                    // this.Close();
                    string folio = dt.Rows[0][5].ToString();
                    string Nombre = dt.Rows[0][6].ToString();
                    string Año = dt.Rows[0][7].ToString();
                    string tipo = dt.Rows[0][8].ToString();
                    string codificador = dt.Rows[0][10].ToString();
                    proyectoras = folio;
                    anorastreo = Año;
                    proyectorastreo = Nombre;
                    carpetarastreo = tipo;
                    codifica = codificador;
                    nomproyec1 = dt.Rows[0][9].ToString();
                }
                ////////////////////////////////////SI EL USUARIO, CONTRASENA O PLANTA ESTA MAL//////////////////////////////////////////////////////////////////////////
                else
                {


                    MessageBox.Show("Datos incorrectos de acceso");
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            finally
            {

            }
        }



    }
}
