using CBR_ADMIN.Administracion;
using CBR_ADMIN.IT.Tickets;
using CBR_ADMIN.Properties;
using CBR_ADMIN.Sistema;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net;
using aejw.Network;
using System.IO;
using System.Diagnostics;
using CBR_VPN;
using System.Runtime.InteropServices;
using System.Net.NetworkInformation;

namespace CBR_ADMIN
{
    public partial class General : Form
    {
        public General()
        {
            InitializeComponent();
        }
        public string Permiso;
        public string Nombre;
        public string Apellido;
        public string Ventana;
        public string Departamento;
        public string Puesto;
        public string Email;
        public string Cif1, Var1, Cif2, Var2, Cif3, Var3, Cif4, Var4;
        int conexionvpn;
        public int acceso1;


        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }




        public static string ObtenerCadena()
        {
            return Settings.Default.CBR_IngenieriaConnectionString;/*Este codigo obtiene los datos de la cadena de conexion declarada en los settings de la aplicacion*/
        }
        SqlConnection conexion = new SqlConnection(ObtenerCadena());
        public void MostrarCuadroInicio()
        {
            Loading Cargando = new Loading();
            Cargando.ShowDialog();

        }
        public void MostrarCuadroInicio2()
        {
            Loading Cargando = new Loading();
            Cargando.ShowDialog();

        }

        public void MostrarCuadroInicio3()
        {
            Loading Cargando = new Loading();
            Cargando.ShowDialog();
        }
        private void levantamientoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Departamento == "Asistente de direccion" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                {


                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central.Controls.Contains(Genera_No_Confor.Instance))
                        {
                            Ventana = "Pantalla: Levantamiento de No conformidad";
                            Genera_No_Confor.Instance.Nombre = Nombre;

                            Genera_No_Confor.Instance.Ventana1 = Ventana;
                            //Genera_No_Confor.Instance.Email = Email;
                            Central.Controls.Add(Genera_No_Confor.Instance);
                            Genera_No_Confor.Instance.Dock = DockStyle.Fill;
                            Genera_No_Confor.Instance.BringToFront();



                        }
                        else
                        {
                            Genera_No_Confor.Instance.BringToFront();
                            Ventana = "Pantalla: Levantamiento de No conformidad";


                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");

                    }
                }

                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                           // MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {//  MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);    

                    }
                    oNetDrive = null;
                    if (Departamento == "Asistente de direccion" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                    {


                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                        {

                            if (!Central.Controls.Contains(Genera_No_Confor.Instance))
                            {
                                Ventana = "Pantalla: Levantamiento de No conformidad";
                                Genera_No_Confor.Instance.Nombre = Nombre;

                                Genera_No_Confor.Instance.Ventana1 = Ventana;
                                //Genera_No_Confor.Instance.Email = Email;
                                Central.Controls.Add(Genera_No_Confor.Instance);
                                Genera_No_Confor.Instance.Dock = DockStyle.Fill;
                                Genera_No_Confor.Instance.BringToFront();



                            }
                            else
                            {
                                Genera_No_Confor.Instance.BringToFront();
                                Ventana = "Pantalla: Levantamiento de No conformidad";


                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");

                        }
                    }

                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }
        public Int32 NumeroConfo, NumeroOport, NumeroOportDep;
        public string NumeroConformidad;
        private void General_Load(object sender, EventArgs e)


        {
            Screen screen = Screen.PrimaryScreen;

            int Height = screen.Bounds.Width;

            int Width = screen.Bounds.Height;
            string het = Convert.ToString(Height);
            string wit = Convert.ToString(Width);
            //label17.Text = het + "  " + wit;
            if (Height < 1366)
            {
                this.Size = new Size(1217, 650);
                this.Central.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
                this.Central2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
                this.Central3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
                this.Central4.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
                this.Central5.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
                this.Central6.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
                this.Central7.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
                this.Central8.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
                this.Central9.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
                this.central10.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
                this.central11.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
                this.Central12.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
                this.central13.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
        
            }



            else { }
            


            if (Cif1 == Var1 & Cif2 == Var2 & Cif3 == Var3 & Cif4 == Var4) { acceso1 = 0; }
            else if (Cif1 == Var1 & Cif2 == Var2 & Cif3 == Var3 & Cif4 != Var4) { acceso1 = 1; }
            else if (Cif1 == Var1 & Cif2 == Var2 & Cif3 != Var3 & Cif4 != Var4) { acceso1 = 2; }
            else if (Cif1 == Var1 & Cif2 != Var2 & Cif3 != Var3 & Cif4 != Var4) { acceso1 = 3; }
            else if (Cif1 != Var1 & Cif2 != Var2 & Cif3 != Var3 & Cif4 != Var4) { acceso1 = 4; }
            Nombreprin.Text = "Bienvenido  " + "" + Nombre + "" + Apellido + " a Cbr Administracion y Servicios";

            timer1.Enabled = true;
            try
            {
                ////////////////////////////////////SE ABRE LA CONEXION/////////////////////////////////////////////////////////////////////////////////
                SqlConnection conexion = new SqlConnection(ObtenerCadena());
                conexion.Open();
                ////////////////////////////////////SE EJECUTA EL COMANDO HACIA LA VISTA Y SE HACE LA VALIDACION//////////////////////////////////////
                SqlCommand cmd = new SqlCommand(
                                                "select COUNT(*) AS conteo from [No_Conformidades] where  " +
                                                "  (Status ='Abierta')" +
                                                "and (Departamento = @Departamento)"
                                                , conexion);
                cmd.Parameters.AddWithValue("Responsable", Nombre);
                cmd.Parameters.AddWithValue("Departamento", Departamento);

                SqlCommand cmd1 = new SqlCommand(
                                    "select COUNT(*) AS conteo from [No_Conformidades] where (Departamento = @Departamento)  and (Status ='Cerrada')"
                                    , conexion);
                cmd1.Parameters.AddWithValue("Departamento", Departamento);


                SqlCommand cmd2 = new SqlCommand(
                                  "select COUNT([# NC]) AS conteo from [No_Conformidades] "
                                  , conexion);


                SqlCommand cmd3 = new SqlCommand(
              "select COUNT(*) AS conteo from [No_Conformidades] where " +
              "  (Status ='Verificando')" +
              "and (Departamento = @Departamento)"
              , conexion);
                cmd3.Parameters.AddWithValue("Responsable", Nombre);
                cmd3.Parameters.AddWithValue("Departamento", Departamento);


                SqlCommand cmd4 = new SqlCommand(
           "select COUNT(*) AS conteo from [Op_Mejora] "
           , conexion);


                SqlCommand cmd5 = new SqlCommand(
                                                "select COUNT(*) AS conteo from [Op_Mejora] where  " +
                                                "  (Status ='Abierta')" +
                                                "and (Departamento = @Departamento)"
                                                , conexion);

                cmd5.Parameters.AddWithValue("Departamento", Departamento);
                SqlCommand cmd6 = new SqlCommand(
                                    "select COUNT(*) AS conteo from [Op_Mejora] where  " +
                                    "  (Status ='Cerrada') " +
                                    "and (Departamento = @Departamento)"
                                    , conexion);

                cmd6.Parameters.AddWithValue("Departamento", Departamento);


                SqlCommand cmd7 = new SqlCommand(
                                    "select COUNT(*) AS conteo from [No_Conformidades] where  " +
                                    "  (Status ='Abierta')" +
                                    "and (Departamento = 'General')"
                                    , conexion);



                SqlCommand cmd8 = new SqlCommand(
                                    "select COUNT(*) AS conteo from [No_Conformidades] where (Departamento = 'General')  and (Status ='Cerrada')"
                                    , conexion);


                SqlCommand cmd9 = new SqlCommand(
                                    "select COUNT(*) AS conteo from [No_Conformidades] where (Departamento = 'General')  and  (Status ='Verificando')"
                                    , conexion);





                Int32 rows_count = Convert.ToInt32(cmd.ExecuteScalar());
                Int32 rows_count1 = Convert.ToInt32(cmd1.ExecuteScalar());
                Int32 rows_count2 = Convert.ToInt32(cmd3.ExecuteScalar());
                Int32 rows_count3 = Convert.ToInt32(cmd5.ExecuteScalar());
                Int32 rows_count4 = Convert.ToInt32(cmd6.ExecuteScalar());
                Int32 rows_count5 = Convert.ToInt32(cmd7.ExecuteScalar());
                Int32 rows_count6 = Convert.ToInt32(cmd9.ExecuteScalar());
                Int32 rows_count7 = Convert.ToInt32(cmd8.ExecuteScalar());

                NumeroConfo = Convert.ToInt32(cmd2.ExecuteScalar());
                NumeroOport = Convert.ToInt32(cmd4.ExecuteScalar());

                cmd.Dispose();
                cmd1.Dispose();
                cmd2.Dispose();
                cmd3.Dispose();
                cmd4.Dispose();
                cmd5.Dispose();
                cmd6.Dispose();
                cmd7.Dispose();
                cmd8.Dispose();
                cmd9.Dispose();

                conexion.Close();

                NumeroConformidad = NumeroConfo.ToString();
                Nc_Abiertas.Text = rows_count.ToString();
                NC_Concluidas.Text = rows_count1.ToString();
                NC_Verificacion.Text = rows_count2.ToString();
                OP_Abiertas.Text = rows_count3.ToString();
                Op_Cerradas.Text = rows_count4.ToString();
                Ncag.Text = rows_count5.ToString();
                Ncvg.Text = rows_count6.ToString();
                Nccg.Text = rows_count7.ToString();
                Nom.Text = Nombre;
                Ape.Text = Apellido;
            }
            catch (Exception d)
            {
                MessageBox.Show(d.Message);
            }
            finally
            {

            }

            Depart.Text = Departamento;
            Pues.Text = Puesto;
            correo.Text = Email;






        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void generacíonToolStripMenuItem_Click(object sender, EventArgs e)
        {

            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Departamento == "Administracion" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                {


                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {


                        if (!Central.Controls.Contains(Genera_No_Confor.Instance))
                        {
                            Ventana = "Pantalla: Levantamiento de No conformidad";
                            Genera_No_Confor.Instance.Nombre = Nombre;
                            Genera_No_Confor.Instance.Ventana1 = Ventana;

                            Genera_No_Confor.Instance.Departamento = Departamento;
                            Central.Controls.Add(Genera_No_Confor.Instance);
                            Genera_No_Confor.Instance.Dock = DockStyle.Fill;
                            Genera_No_Confor.Instance.BringToFront();


                        }
                        else
                        {
                            Genera_No_Confor.Instance.BringToFront();
                            Ventana = "Pantalla: Levantamiento de No conformidad";

                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");

                    }
                }
                else { MessageBox.Show("No tiene acceso a este modulo"); }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                           // MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {//  MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);     
                    }
                    oNetDrive = null;
                    if (Departamento == "Administracion" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                    {


                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                        {


                            if (!Central.Controls.Contains(Genera_No_Confor.Instance))
                            {
                                Ventana = "Pantalla: Levantamiento de No conformidad";
                                Genera_No_Confor.Instance.Nombre = Nombre;
                                Genera_No_Confor.Instance.Ventana1 = Ventana;

                                Genera_No_Confor.Instance.Departamento = Departamento;
                                Central.Controls.Add(Genera_No_Confor.Instance);
                                Genera_No_Confor.Instance.Dock = DockStyle.Fill;
                                Genera_No_Confor.Instance.BringToFront();


                            }
                            else
                            {
                                Genera_No_Confor.Instance.BringToFront();
                                Ventana = "Pantalla: Levantamiento de No conformidad";

                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");

                        }
                    }
                    else { MessageBox.Show("No tiene acceso a este modulo"); }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }

        }




        private void generacionToolStripMenuItem_Click(object sender, EventArgs e)
        {

            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central2.Controls.Contains(Genera_No_Confor.Instance))
                    {
                        Ventana = "Pantalla: Levantamiento de No conformidad";
                        Genera_No_Confor.Instance.Nombre = Nombre;
                        Genera_No_Confor.Instance.Ventana1 = Ventana;

                        Central2.Controls.Add(Genera_No_Confor.Instance);
                        Genera_No_Confor.Instance.Dock = DockStyle.Fill;
                        Genera_No_Confor.Instance.BringToFront();



                    }
                    else
                    {
                        Genera_No_Confor.Instance.BringToFront();
                        Ventana = "Pantalla: Levantamiento de No conformidad";

                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {

                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                           // MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;


                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); }
                    oNetDrive = null;
                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central2.Controls.Contains(Genera_No_Confor.Instance))
                        {
                            Ventana = "Pantalla: Levantamiento de No conformidad";
                            Genera_No_Confor.Instance.Nombre = Nombre;
                            Genera_No_Confor.Instance.Ventana1 = Ventana;

                            Central2.Controls.Add(Genera_No_Confor.Instance);
                            Genera_No_Confor.Instance.Dock = DockStyle.Fill;
                            Genera_No_Confor.Instance.BringToFront();



                        }
                        else
                        {
                            Genera_No_Confor.Instance.BringToFront();
                            Ventana = "Pantalla: Levantamiento de No conformidad";

                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void generaciónToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central3.Controls.Contains(Genera_No_Confor.Instance))
                    {
                        Ventana = "Pantalla: Levantamiento de No conformidad";
                        Genera_No_Confor.Instance.Nombre = Nombre;
                        Genera_No_Confor.Instance.Ventana1 = Ventana;

                        Central3.Controls.Add(Genera_No_Confor.Instance);
                        Genera_No_Confor.Instance.Dock = DockStyle.Fill;
                        Genera_No_Confor.Instance.BringToFront();



                    }
                    else
                    {
                        Genera_No_Confor.Instance.BringToFront();
                        Ventana = "Pantalla: Levantamiento de No conformidad";

                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                        //    MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); }
                    oNetDrive = null;
                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central3.Controls.Contains(Genera_No_Confor.Instance))
                        {
                            Ventana = "Pantalla: Levantamiento de No conformidad";
                            Genera_No_Confor.Instance.Nombre = Nombre;
                            Genera_No_Confor.Instance.Ventana1 = Ventana;

                            Central3.Controls.Add(Genera_No_Confor.Instance);
                            Genera_No_Confor.Instance.Dock = DockStyle.Fill;
                            Genera_No_Confor.Instance.BringToFront();



                        }
                        else
                        {
                            Genera_No_Confor.Instance.BringToFront();
                            Ventana = "Pantalla: Levantamiento de No conformidad";

                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }



        }

        private void generaciónToolStripMenuItem1_Click(object sender, EventArgs e)
        {

            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central4.Controls.Contains(Genera_No_Confor.Instance))
                    {
                        Ventana = "Pantalla: Levantamiento de No conformidad";
                        Genera_No_Confor.Instance.Nombre = Nombre;
                        Genera_No_Confor.Instance.Ventana1 = Ventana;

                        Central4.Controls.Add(Genera_No_Confor.Instance);
                        Genera_No_Confor.Instance.Dock = DockStyle.Fill;
                        Genera_No_Confor.Instance.BringToFront();



                    }
                    else
                    {
                        Genera_No_Confor.Instance.BringToFront();
                        Ventana = "Pantalla: Levantamiento de No conformidad";

                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");

                }
            }





        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central5.Controls.Contains(Genera_No_Confor.Instance))
                    {
                        Ventana = "Pantalla: Levantamiento de No conformidad";
                        Genera_No_Confor.Instance.Nombre = Nombre;
                        Genera_No_Confor.Instance.Ventana1 = Ventana;

                        Central5.Controls.Add(Genera_No_Confor.Instance);
                        Genera_No_Confor.Instance.Dock = DockStyle.Fill;
                        Genera_No_Confor.Instance.BringToFront();



                    }
                    else
                    {
                        Genera_No_Confor.Instance.BringToFront();
                        Ventana = "Pantalla: Levantamiento de No conformidad";

                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                         //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); }
                    oNetDrive = null;
                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central5.Controls.Contains(Genera_No_Confor.Instance))
                        {
                            Ventana = "Pantalla: Levantamiento de No conformidad";
                            Genera_No_Confor.Instance.Nombre = Nombre;
                            Genera_No_Confor.Instance.Ventana1 = Ventana;

                            Central5.Controls.Add(Genera_No_Confor.Instance);
                            Genera_No_Confor.Instance.Dock = DockStyle.Fill;
                            Genera_No_Confor.Instance.BringToFront();



                        }
                        else
                        {
                            Genera_No_Confor.Instance.BringToFront();
                            Ventana = "Pantalla: Levantamiento de No conformidad";

                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }



        }

        private void toolStripMenuItem9_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central6.Controls.Contains(Genera_No_Confor.Instance))
                    {
                        Ventana = "Pantalla: Levantamiento de No conformidad";
                        Genera_No_Confor.Instance.Nombre = Nombre;
                        Genera_No_Confor.Instance.Ventana1 = Ventana;

                        Central6.Controls.Add(Genera_No_Confor.Instance);
                        Genera_No_Confor.Instance.Dock = DockStyle.Fill;
                        Genera_No_Confor.Instance.BringToFront();



                    }
                    else
                    {
                        Genera_No_Confor.Instance.BringToFront();
                        Ventana = "Pantalla: Levantamiento de No conformidad";

                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                       //     MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); }
                    oNetDrive = null;
                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central6.Controls.Contains(Genera_No_Confor.Instance))
                        {
                            Ventana = "Pantalla: Levantamiento de No conformidad";
                            Genera_No_Confor.Instance.Nombre = Nombre;
                            Genera_No_Confor.Instance.Ventana1 = Ventana;

                            Central6.Controls.Add(Genera_No_Confor.Instance);
                            Genera_No_Confor.Instance.Dock = DockStyle.Fill;
                            Genera_No_Confor.Instance.BringToFront();



                        }
                        else
                        {
                            Genera_No_Confor.Instance.BringToFront();
                            Ventana = "Pantalla: Levantamiento de No conformidad";

                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }










        }

        private void generaciónToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;

                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central7.Controls.Contains(Genera_No_Confor.Instance))
                    {
                        Ventana = "Pantalla: Levantamiento de No conformidad";
                        Genera_No_Confor.Instance.Nombre = Nombre;
                        Genera_No_Confor.Instance.Ventana1 = Ventana;

                        Central7.Controls.Add(Genera_No_Confor.Instance);
                        Genera_No_Confor.Instance.Dock = DockStyle.Fill;
                        Genera_No_Confor.Instance.BringToFront();



                    }
                    else
                    {
                        Genera_No_Confor.Instance.BringToFront();
                        Ventana = "Pantalla: Levantamiento de No conformidad";

                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {

                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);   
                    }
                    oNetDrive = null;

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central7.Controls.Contains(Genera_No_Confor.Instance))
                        {
                            Ventana = "Pantalla: Levantamiento de No conformidad";
                            Genera_No_Confor.Instance.Nombre = Nombre;
                            Genera_No_Confor.Instance.Ventana1 = Ventana;

                            Central7.Controls.Add(Genera_No_Confor.Instance);
                            Genera_No_Confor.Instance.Dock = DockStyle.Fill;
                            Genera_No_Confor.Instance.BringToFront();



                        }
                        else
                        {
                            Genera_No_Confor.Instance.BringToFront();
                            Ventana = "Pantalla: Levantamiento de No conformidad";

                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void consultaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;

                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central2.Controls.Contains(Consulta_No_Conf.Instance))
                    {
                        Ventana = "Pantalla: Consulta de No Conformidad";
                        Consulta_No_Conf.Instance.Departamento = Departamento;
                        Central2.Controls.Add(Consulta_No_Conf.Instance);
                        Consulta_No_Conf.Instance.Dock = DockStyle.Fill;
                        Consulta_No_Conf.Instance.BringToFront();



                    }
                    else
                    {
                        Consulta_No_Conf.Instance.BringToFront();
                        Ventana = "Pantalla: Consulta de No Conformidad";

                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);    
                    }
                    oNetDrive = null;

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central2.Controls.Contains(Consulta_No_Conf.Instance))
                        {
                            Ventana = "Pantalla: Consulta de No Conformidad";
                            Consulta_No_Conf.Instance.Departamento = Departamento;
                            Central2.Controls.Add(Consulta_No_Conf.Instance);
                            Consulta_No_Conf.Instance.Dock = DockStyle.Fill;
                            Consulta_No_Conf.Instance.BringToFront();



                        }
                        else
                        {
                            Consulta_No_Conf.Instance.BringToFront();
                            Ventana = "Pantalla: Consulta de No Conformidad";

                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void consultaToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;

                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central3.Controls.Contains(Consulta_No_Conf.Instance))
                    {
                        Ventana = "Pantalla: Consulta de No Conformidad";
                        Consulta_No_Conf.Instance.Departamento = Departamento;
                        Central3.Controls.Add(Consulta_No_Conf.Instance);
                        Consulta_No_Conf.Instance.Dock = DockStyle.Fill;
                        Consulta_No_Conf.Instance.BringToFront();

                    }
                    else
                    {
                        Consulta_No_Conf.Instance.BringToFront();
                        Ventana = "Pantalla: Consulta de No Conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);   
                    }
                    oNetDrive = null;

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central3.Controls.Contains(Consulta_No_Conf.Instance))
                        {
                            Ventana = "Pantalla: Consulta de No Conformidad";
                            Consulta_No_Conf.Instance.Departamento = Departamento;
                            Central3.Controls.Add(Consulta_No_Conf.Instance);
                            Consulta_No_Conf.Instance.Dock = DockStyle.Fill;
                            Consulta_No_Conf.Instance.BringToFront();

                        }
                        else
                        {
                            Consulta_No_Conf.Instance.BringToFront();
                            Ventana = "Pantalla: Consulta de No Conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void consultaToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Departamento == "Administracion" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                {

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central.Controls.Contains(Consulta_No_Conf.Instance))
                        {
                            Ventana = "Pantalla: Consulta de No Conformidad";
                            Consulta_No_Conf.Instance.Departamento = Departamento;
                            Central.Controls.Add(Consulta_No_Conf.Instance);
                            Consulta_No_Conf.Instance.Dock = DockStyle.Fill;
                            Consulta_No_Conf.Instance.BringToFront();


                        }
                        else
                        {
                            Consulta_No_Conf.Instance.BringToFront();
                            Ventana = "Pantalla: Consulta de No Conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                            conexionvpn = 0;
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();

                    }
                    catch (Exception err) { }
                    //   { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  //   }
                    oNetDrive = null;
                    if (Departamento == "Administracion" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                    {

                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                        {

                            if (!Central.Controls.Contains(Consulta_No_Conf.Instance))
                            {
                                Ventana = "Pantalla: Consulta de No Conformidad";
                                Consulta_No_Conf.Instance.Departamento = Departamento;
                                Central.Controls.Add(Consulta_No_Conf.Instance);
                                Consulta_No_Conf.Instance.Dock = DockStyle.Fill;
                                Consulta_No_Conf.Instance.BringToFront();


                            }
                            else
                            {
                                Consulta_No_Conf.Instance.BringToFront();
                                Ventana = "Pantalla: Consulta de No Conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void toolStripMenuItem4_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;

                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central5.Controls.Contains(Consulta_No_Conf.Instance))
                    {
                        Ventana = "Pantalla: Consulta de No Conformidad";
                        Consulta_No_Conf.Instance.Departamento = Departamento;
                        Central5.Controls.Add(Consulta_No_Conf.Instance);
                        Consulta_No_Conf.Instance.Dock = DockStyle.Fill;
                        Consulta_No_Conf.Instance.BringToFront();


                    }
                    else
                    {
                        Consulta_No_Conf.Instance.BringToFront();
                        Ventana = "Pantalla: Consulta de No Conformidad";
                    }
                }
                else
                {

                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {

                        //   MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central5.Controls.Contains(Consulta_No_Conf.Instance))
                        {
                            Ventana = "Pantalla: Consulta de No Conformidad";
                            Consulta_No_Conf.Instance.Departamento = Departamento;
                            Central5.Controls.Add(Consulta_No_Conf.Instance);
                            Consulta_No_Conf.Instance.Dock = DockStyle.Fill;
                            Consulta_No_Conf.Instance.BringToFront();


                        }
                        else
                        {
                            Consulta_No_Conf.Instance.BringToFront();
                            Ventana = "Pantalla: Consulta de No Conformidad";
                        }
                    }
                    else
                    {

                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void toolStripMenuItem12_Click(object sender, EventArgs e)
        {

            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central6.Controls.Contains(Consulta_No_Conf.Instance))
                    {
                        Ventana = "Pantalla: Consulta de No Conformidad";
                        Consulta_No_Conf.Instance.Departamento = Departamento;
                        Central6.Controls.Add(Consulta_No_Conf.Instance);
                        Consulta_No_Conf.Instance.Dock = DockStyle.Fill;
                        Consulta_No_Conf.Instance.BringToFront();


                    }
                    else
                    {
                        Consulta_No_Conf.Instance.BringToFront();
                        Ventana = "Pantalla: Consulta de No Conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;
                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central6.Controls.Contains(Consulta_No_Conf.Instance))
                        {
                            Ventana = "Pantalla: Consulta de No Conformidad";
                            Consulta_No_Conf.Instance.Departamento = Departamento;
                            Central6.Controls.Add(Consulta_No_Conf.Instance);
                            Consulta_No_Conf.Instance.Dock = DockStyle.Fill;
                            Consulta_No_Conf.Instance.BringToFront();


                        }
                        else
                        {
                            Consulta_No_Conf.Instance.BringToFront();
                            Ventana = "Pantalla: Consulta de No Conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void consultaToolStripMenuItem3_Click(object sender, EventArgs e)
        {

        }

        private void consultaToolStripMenuItem4_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central4.Controls.Contains(Consulta_No_Conf.Instance))
                    {
                        Ventana = "Pantalla: Consulta de No Conformidad";
                        Consulta_No_Conf.Instance.Departamento = Departamento;
                        Central4.Controls.Add(Consulta_No_Conf.Instance);
                        Consulta_No_Conf.Instance.Dock = DockStyle.Fill;
                        Consulta_No_Conf.Instance.BringToFront();


                    }
                    else
                    {
                        Consulta_No_Conf.Instance.BringToFront();
                        Ventana = "Pantalla: Consulta de No Conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {//  MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);   
                    }
                    oNetDrive = null;
                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5") if (Permiso == "1" || Permiso == "2")
                        {

                            if (!Central4.Controls.Contains(Consulta_No_Conf.Instance))
                            {
                                Ventana = "Pantalla: Consulta de No Conformidad";
                                Consulta_No_Conf.Instance.Departamento = Departamento;
                                Central4.Controls.Add(Consulta_No_Conf.Instance);
                                Consulta_No_Conf.Instance.Dock = DockStyle.Fill;
                                Consulta_No_Conf.Instance.BringToFront();


                            }
                            else
                            {
                                Consulta_No_Conf.Instance.BringToFront();
                                Ventana = "Pantalla: Consulta de No Conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }



        }

        private void generalToolStripMenuItem8_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else { conexionvpn = 2; }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);   
                    }
                    oNetDrive = null;
                    if (Departamento == "Asistente de direccion" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                    {

                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5") if (Permiso == "1" || Permiso == "2")
                            {

                                if (!Central.Controls.Contains(Consulta_Oportunidades_Mejora.Instance))
                                {
                                    Ventana = "Pantalla: Consulta de No Conformidad";
                                    Consulta_Oportunidades_Mejora.Instance.Departamento = Departamento;
                                    Central.Controls.Add(Consulta_Oportunidades_Mejora.Instance);

                                    Consulta_Oportunidades_Mejora.Instance.Dock = DockStyle.Fill;
                                    Consulta_Oportunidades_Mejora.Instance.BringToFront();


                                }
                                else
                                {
                                    Consulta_Oportunidades_Mejora.Instance.BringToFront();
                                    Ventana = "Pantalla: Consulta de No Conformidad";
                                }
                            }
                            else
                            {
                                MessageBox.Show("No tiene acceso a este modulo");
                            }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void generalToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Departamento == "Administracion" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                {

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central.Controls.Contains(Consulta_No_Conf_general.Instance))
                        {
                            Ventana = "Pantalla: Consulta de No Conformidad";

                            Central.Controls.Add(Consulta_No_Conf_general.Instance);
                            Consulta_No_Conf_general.Instance.Dock = DockStyle.Fill;
                            Consulta_No_Conf_general.Instance.BringToFront();


                        }
                        else
                        {
                            Consulta_No_Conf_general.Instance.BringToFront();
                            Ventana = "Pantalla: Consulta de No Conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //    MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                            conexionvpn = 0;
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);   
                    }
                    oNetDrive = null;
                    if (Departamento == "Administracion" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                    {

                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                        {

                            if (!Central.Controls.Contains(Consulta_No_Conf_general.Instance))
                            {
                                Ventana = "Pantalla: Consulta de No Conformidad";

                                Central.Controls.Add(Consulta_No_Conf_general.Instance);
                                Consulta_No_Conf_general.Instance.Dock = DockStyle.Fill;
                                Consulta_No_Conf_general.Instance.BringToFront();


                            }
                            else
                            {
                                Consulta_No_Conf_general.Instance.BringToFront();
                                Ventana = "Pantalla: Consulta de No Conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void generalToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Departamento == "Administracion" || Departamento == "SuperAdmin" || Departamento == "Direccion" || Departamento == "Compras")
                {

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central.Controls.Contains(Consulta_No_Conf_general.Instance))
                        {
                            Ventana = "Pantalla: Consulta de No Conformidad";

                            Central.Controls.Add(Consulta_No_Conf_general.Instance);
                            Consulta_No_Conf_general.Instance.Dock = DockStyle.Fill;
                            Consulta_No_Conf_general.Instance.BringToFront();


                        }
                        else
                        {
                            Consulta_No_Conf_general.Instance.BringToFront();
                            Ventana = "Pantalla: Consulta de No Conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //    MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                            conexionvpn = 0;
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);   
                    }
                    oNetDrive = null;
                    if (Departamento == "Administracion" || Departamento == "SuperAdmin" || Departamento == "Direccion" || Departamento == "Compras")
                    {

                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                        {

                            if (!Central.Controls.Contains(Consulta_No_Conf_general.Instance))
                            {
                                Ventana = "Pantalla: Consulta de No Conformidad";

                                Central.Controls.Add(Consulta_No_Conf_general.Instance);
                                Consulta_No_Conf_general.Instance.Dock = DockStyle.Fill;
                                Consulta_No_Conf_general.Instance.BringToFront();


                            }
                            else
                            {
                                Consulta_No_Conf_general.Instance.BringToFront();
                                Ventana = "Pantalla: Consulta de No Conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void generalToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;

            }
            else
            {
                conexionvpn = 2;

                if (Departamento == "Cobranza" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                {

                    if (Permiso == "1" || Permiso == "2")
                    {

                        if (!Central.Controls.Contains(Consulta_No_Conf_general.Instance))
                        {
                            Ventana = "Pantalla: Consulta de No Conformidad";

                            Central.Controls.Add(Consulta_No_Conf_general.Instance);
                            Consulta_No_Conf_general.Instance.Dock = DockStyle.Fill;
                            Consulta_No_Conf_general.Instance.BringToFront();


                        }
                        else
                        {
                            Consulta_No_Conf_general.Instance.BringToFront();
                            Ventana = "Pantalla: Consulta de No Conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {
                        //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;

                    if (Departamento == "Cobranza" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                    {

                        if (Permiso == "1" || Permiso == "2")
                        {

                            if (!Central.Controls.Contains(Consulta_No_Conf_general.Instance))
                            {
                                Ventana = "Pantalla: Consulta de No Conformidad";

                                Central.Controls.Add(Consulta_No_Conf_general.Instance);
                                Consulta_No_Conf_general.Instance.Dock = DockStyle.Fill;
                                Consulta_No_Conf_general.Instance.BringToFront();


                            }
                            else
                            {
                                Consulta_No_Conf_general.Instance.BringToFront();
                                Ventana = "Pantalla: Consulta de No Conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void generalToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;

                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central2.Controls.Contains(Consulta_No_Conf_general.Instance))
                    {
                        Ventana = "Pantalla: Consulta de No Conformidad";

                        Central2.Controls.Add(Consulta_No_Conf_general.Instance);
                        Consulta_No_Conf_general.Instance.Dock = DockStyle.Fill;
                        Consulta_No_Conf_general.Instance.BringToFront();


                    }
                    else
                    {
                        Consulta_No_Conf_general.Instance.BringToFront();
                        Ventana = "Pantalla: Consulta de No Conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //    MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {
                        //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);   
                    }
                    oNetDrive = null;

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central2.Controls.Contains(Consulta_No_Conf_general.Instance))
                        {
                            Ventana = "Pantalla: Consulta de No Conformidad";

                            Central2.Controls.Add(Consulta_No_Conf_general.Instance);
                            Consulta_No_Conf_general.Instance.Dock = DockStyle.Fill;
                            Consulta_No_Conf_general.Instance.BringToFront();


                        }
                        else
                        {
                            Consulta_No_Conf_general.Instance.BringToFront();
                            Ventana = "Pantalla: Consulta de No Conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void generalToolStripMenuItem4_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;

                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central3.Controls.Contains(Consulta_No_Conf_general.Instance))
                    {
                        Ventana = "Pantalla: Consulta de No Conformidad";

                        Central3.Controls.Add(Consulta_No_Conf_general.Instance);
                        Consulta_No_Conf_general.Instance.Dock = DockStyle.Fill;
                        Consulta_No_Conf_general.Instance.BringToFront();


                    }
                    else
                    {
                        Consulta_No_Conf_general.Instance.BringToFront();
                        Ventana = "Pantalla: Consulta de No Conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            // MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);   
                    }
                    oNetDrive = null;

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central3.Controls.Contains(Consulta_No_Conf_general.Instance))
                        {
                            Ventana = "Pantalla: Consulta de No Conformidad";

                            Central3.Controls.Add(Consulta_No_Conf_general.Instance);
                            Consulta_No_Conf_general.Instance.Dock = DockStyle.Fill;
                            Consulta_No_Conf_general.Instance.BringToFront();


                        }
                        else
                        {
                            Consulta_No_Conf_general.Instance.BringToFront();
                            Ventana = "Pantalla: Consulta de No Conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void generalToolStripMenuItem5_Click(object sender, EventArgs e)
        {

            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central4.Controls.Contains(Consulta_No_Conf_general.Instance))
                    {
                        Ventana = "Pantalla: Consulta de No Conformidad";

                        Central4.Controls.Add(Consulta_No_Conf_general.Instance);
                        Consulta_No_Conf_general.Instance.Dock = DockStyle.Fill;
                        Consulta_No_Conf_general.Instance.BringToFront();


                    }
                    else
                    {
                        Consulta_No_Conf_general.Instance.BringToFront();
                        Ventana = "Pantalla: Consulta de No Conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); 
                    }
                    oNetDrive = null;
                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central4.Controls.Contains(Consulta_No_Conf_general.Instance))
                        {
                            Ventana = "Pantalla: Consulta de No Conformidad";

                            Central4.Controls.Add(Consulta_No_Conf_general.Instance);
                            Consulta_No_Conf_general.Instance.Dock = DockStyle.Fill;
                            Consulta_No_Conf_general.Instance.BringToFront();


                        }
                        else
                        {
                            Consulta_No_Conf_general.Instance.BringToFront();
                            Ventana = "Pantalla: Consulta de No Conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void generalToolStripMenuItem6_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central5.Controls.Contains(Consulta_No_Conf_general.Instance))
                    {
                        Ventana = "Pantalla: Consulta de No Conformidad";

                        Central5.Controls.Add(Consulta_No_Conf_general.Instance);
                        Consulta_No_Conf_general.Instance.Dock = DockStyle.Fill;
                        Consulta_No_Conf_general.Instance.BringToFront();


                    }
                    else
                    {
                        Consulta_No_Conf_general.Instance.BringToFront();
                        Ventana = "Pantalla: Consulta de No Conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            // MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); 
                    }
                    oNetDrive = null;
                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central5.Controls.Contains(Consulta_No_Conf_general.Instance))
                        {
                            Ventana = "Pantalla: Consulta de No Conformidad";

                            Central5.Controls.Add(Consulta_No_Conf_general.Instance);
                            Consulta_No_Conf_general.Instance.Dock = DockStyle.Fill;
                            Consulta_No_Conf_general.Instance.BringToFront();


                        }
                        else
                        {
                            Consulta_No_Conf_general.Instance.BringToFront();
                            Ventana = "Pantalla: Consulta de No Conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }




        }

        private void generalToolStripMenuItem7_Click(object sender, EventArgs e)
        {

            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central6.Controls.Contains(Consulta_No_Conf_general.Instance))
                    {
                        Ventana = "Pantalla: Consulta de No Conformidad";

                        Central6.Controls.Add(Consulta_No_Conf_general.Instance);
                        Consulta_No_Conf_general.Instance.Dock = DockStyle.Fill;
                        Consulta_No_Conf_general.Instance.BringToFront();


                    }
                    else
                    {
                        Consulta_No_Conf_general.Instance.BringToFront();
                        Ventana = "Pantalla: Consulta de No Conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //    MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {
                        //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); 
                    }
                    oNetDrive = null;
                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central6.Controls.Contains(Consulta_No_Conf_general.Instance))
                        {
                            Ventana = "Pantalla: Consulta de No Conformidad";

                            Central6.Controls.Add(Consulta_No_Conf_general.Instance);
                            Consulta_No_Conf_general.Instance.Dock = DockStyle.Fill;
                            Consulta_No_Conf_general.Instance.BringToFront();


                        }
                        else
                        {
                            Consulta_No_Conf_general.Instance.BringToFront();
                            Ventana = "Pantalla: Consulta de No Conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }



        }

        private void administracionToolStripMenuItem_Click(object sender, EventArgs e)
        {


        }

        private void seguimientoToolStripMenuItem4_Click(object sender, EventArgs e)
        {


            if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
            {

                if (!Central.Controls.Contains(Seguimiento_No_confor.Instance))
                {
                    Ventana = "Pantalla: Seguimiento de No conformidad";
                    Seguimiento_No_confor.Instance.Departamento = Departamento;
                    Seguimiento_No_confor.Instance.Nombre = Nombre;
                    Seguimiento_No_confor.Instance.Apellido = Apellido;
                    Seguimiento_No_confor.Instance.NoConfor = NumeroConformidad;
                    Seguimiento_No_confor.Instance.Ventana1 = Ventana;
                    Central.Controls.Add(Seguimiento_No_confor.Instance);
                    Seguimiento_No_confor.Instance.Dock = DockStyle.Fill;
                    Seguimiento_No_confor.Instance.BringToFront();


                }
                else
                {
                    Seguimiento_No_confor.Instance.BringToFront();
                    Ventana = "Pantalla: Seguimiento de No conformidad";
                }
            }
            else
            {
                MessageBox.Show("No tiene acceso a este modulo");
            }
        }

        private void cierreToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Departamento == "Asistente de direccion" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                {

                    if (Permiso == "1" || Permiso == "2")
                    {

                        if (!Central.Controls.Contains(Cierre_No_confor.Instance))
                        {
                            Ventana = "Pantalla: Seguimiento de No conformidad";
                            Central.Controls.Add(Cierre_No_confor.Instance);
                            Cierre_No_confor.Instance.Dock = DockStyle.Fill;
                            Cierre_No_confor.Instance.BringToFront();


                        }
                        else
                        {
                            Cierre_No_confor.Instance.BringToFront();
                            Ventana = "Pantalla: Seguimiento de No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //    MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {//  MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);   
                    }
                    oNetDrive = null;
                    if (Departamento == "Asistente de direccion" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                    {

                        if (Permiso == "1" || Permiso == "2")
                        {

                            if (!Central.Controls.Contains(Cierre_No_confor.Instance))
                            {
                                Ventana = "Pantalla: Seguimiento de No conformidad";
                                Central.Controls.Add(Cierre_No_confor.Instance);
                                Cierre_No_confor.Instance.Dock = DockStyle.Fill;
                                Cierre_No_confor.Instance.BringToFront();


                            }
                            else
                            {
                                Cierre_No_confor.Instance.BringToFront();
                                Ventana = "Pantalla: Seguimiento de No conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }
        public static string FolderPath => string.Concat(Directory.GetCurrentDirectory(),
         "\\VPN");
        private void pictureBox2_Click(object sender, EventArgs e)
        {


            if (conexionvpn == 1) { }
            else if (conexionvpn == 2)
            { 
            File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");

            var newProcess = new Process
            {
                StartInfo =
                {
                    FileName = FolderPath + "\\VpnDisconnect.bat",
                    WindowStyle = ProcessWindowStyle.Minimized
                }
            };

                newProcess.Start();
                newProcess.WaitForExit();
            }

            NetworkDrive oNetDrive1 = new NetworkDrive();
            try
            {
                //set propertys
                oNetDrive1.Force = true;
                oNetDrive1.LocalDrive = "G";
                oNetDrive1.ShareName = url.Text;
                //match call to options provided
                oNetDrive1.MapDrive();

                //update status
            }
            catch (Exception err)
            {
                //report error

           //     MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);

            }

            NetworkDrive oNetDrive = new NetworkDrive();

            try
            {

                oNetDrive.LocalDrive = "G";
                oNetDrive.UnMapDrive();

                //update status
            }
            catch (Exception err)
            {
                //report error

              //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);

            }

           

            this.Hide();
            Application.Restart();
            Login genera = new Login();
            Form fc = Application.OpenForms["Form2"];

            if (fc != null) { fc.Close(); }



            genera.Show();
          
         

     

       
        }

        private void detalleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Departamento == "Asistente de direccion" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                {

                    if (Permiso == "1" || Permiso == "2")
                    {

                        if (!Central.Controls.Contains(Detalle.Instance))
                        {
                            Ventana = "Pantalla: Detalle No conformidad";
                            Central.Controls.Add(Detalle.Instance);
                            Detalle.Instance.Dock = DockStyle.Fill;
                            Detalle.Instance.BringToFront();


                        }
                        else
                        {
                            Detalle.Instance.BringToFront();
                            Detalle.Instance.Dock = DockStyle.Fill;
                            Ventana = "Pantalla: Detalle No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); 
                    }
                    oNetDrive = null;
                    if (Departamento == "Asistente de direccion" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                    {

                        if (Permiso == "1" || Permiso == "2")
                        {

                            if (!Central.Controls.Contains(Detalle.Instance))
                            {
                                Ventana = "Pantalla: Detalle No conformidad";
                                Central.Controls.Add(Detalle.Instance);
                                Detalle.Instance.Dock = DockStyle.Fill;
                                Detalle.Instance.BringToFront();


                            }
                            else
                            {
                                Detalle.Instance.BringToFront();
                                Detalle.Instance.Dock = DockStyle.Fill;
                                Ventana = "Pantalla: Detalle No conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void seguimientoToolStripMenuItem7_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else { conexionvpn = 2; }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);   
                    }
                    oNetDrive = null;

                    if (Departamento == "Asistente de direccion" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                    {

                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                        {

                            if (!Central.Controls.Contains(Seguimiento_AM.Instance))
                            {
                                Ventana = "Pantalla: Detalle No conformidad";
                                Seguimiento_AM.Instance.Nombre = Nombre;
                                Seguimiento_AM.Instance.Apellido = Apellido;
                                Seguimiento_AM.Instance.Departamento = Departamento;
                                Seguimiento_AM.Instance.Ventana1 = Ventana;
                                Central.Controls.Add(Seguimiento_AM.Instance);
                                Seguimiento_AM.Instance.Dock = DockStyle.Fill;
                                Seguimiento_AM.Instance.BringToFront();


                            }
                            else
                            {
                                Seguimiento_AM.Instance.BringToFront();
                                Ventana = "Pantalla: Detalle No conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                    }

                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }

        }

        private void cierreToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else { conexionvpn = 2; }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); 
                    }
                    oNetDrive = null;
                    if (Departamento == "Asistente de direccion" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                    {

                        if (Permiso == "1" || Permiso == "2")
                        {

                            if (!Central.Controls.Contains(Ciere_Accion_Mejora.Instance))
                            {
                                Ventana = "Pantalla: Detalle No conformidad";
                                Ciere_Accion_Mejora.Instance.Nombre = Nombre;
                                Ciere_Accion_Mejora.Instance.Departamento = Departamento;
                                Ciere_Accion_Mejora.Instance.Ventana1 = Ventana;
                                Central.Controls.Add(Ciere_Accion_Mejora.Instance);
                                Ciere_Accion_Mejora.Instance.Dock = DockStyle.Fill;
                                Ciere_Accion_Mejora.Instance.BringToFront();


                            }
                            else
                            {
                                Seguimiento_AM.Instance.BringToFront();
                                Ventana = "Pantalla: Detalle No conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                    }

                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }



        }

        private void consultaToolStripMenuItem7_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else { conexionvpn = 2; }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {
                        //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); 
                    }
                    oNetDrive = null;

                    if (Departamento == "Asistente de direccion" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                    {

                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                        {

                            if (!Central.Controls.Contains(Consulta_OportunidadesMejora_General.Instance))
                            {
                                Ventana = "Pantalla: Detalle No conformidad";

                                Central.Controls.Add(Consulta_OportunidadesMejora_General.Instance);
                                Consulta_OportunidadesMejora_General.Instance.Dock = DockStyle.Fill;
                                Consulta_OportunidadesMejora_General.Instance.BringToFront();


                            }
                            else
                            {
                                Consulta_OportunidadesMejora_General.Instance.BringToFront();
                                Ventana = "Pantalla: Detalle No conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }

        }

        private void consultaToolStripMenuItem8_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central7.Controls.Contains(Consulta_No_Conf.Instance))
                    {
                        Ventana = "Pantalla: Consulta de No Conformidad";
                        Consulta_No_Conf.Instance.Departamento = Departamento;
                        Central7.Controls.Add(Consulta_No_Conf.Instance);
                        Consulta_No_Conf.Instance.Dock = DockStyle.Fill;
                        Consulta_No_Conf.Instance.BringToFront();


                    }
                    else
                    {
                        Consulta_No_Conf.Instance.BringToFront();
                        Ventana = "Pantalla: Consulta de No Conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {
                        //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); 
                    }
                    oNetDrive = null;
                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central7.Controls.Contains(Consulta_No_Conf.Instance))
                        {
                            Ventana = "Pantalla: Consulta de No Conformidad";
                            Consulta_No_Conf.Instance.Departamento = Departamento;
                            Central7.Controls.Add(Consulta_No_Conf.Instance);
                            Consulta_No_Conf.Instance.Dock = DockStyle.Fill;
                            Consulta_No_Conf.Instance.BringToFront();


                        }
                        else
                        {
                            Consulta_No_Conf.Instance.BringToFront();
                            Ventana = "Pantalla: Consulta de No Conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            Hora.Text = DateTime.Now.ToString("hh:mm:ss");
            horra.Text = DateTime.Now.ToString("hh:mm:ss");
        }

        private void tabControl1_Selected(object sender, TabControlEventArgs e)
        {


            if (e.TabPage == Administracion)
            {

                if (acceso1 == 0)
                {
                    if (Departamento == "Administracion" || Departamento == "Sistemas" || Departamento == "Asistente de direccion" || Departamento == "Cobranza" || Departamento == "Compras" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                    {
                        MessageBox.Show("Acceso concedido");


                    }
                    else { Administracion.Parent = null; MessageBox.Show("No tiene acceso a este modulo"); }
                    Administracion.Parent = Base;
                }
                else { MessageBox.Show("Se ha bloqueado la aplicacion con cifrado nivel" + acceso1); Administracion.Parent = null; }
            }

            if (e.TabPage == Potabiliza)
            {
                if (acceso1 == 0)
                {
                    if (Departamento == "Potabiliza" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                    {
                        MessageBox.Show("Acceso concedido");

                    }
                    else { Potabiliza.Parent = null; MessageBox.Show("No tiene acceso a este modulo"); }
                    Potabiliza.Parent = Base;
                }
                else { MessageBox.Show("Se ha bloqueado la aplicacion con cifrado nivel" + acceso1); Potabiliza.Parent = null; }

            }

            if (e.TabPage == Atencion)
            {
                if (acceso1 == 0)
                {
                    if (Departamento == "Atencion a Clientes" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                    {
                        MessageBox.Show("Acceso concedido");

                    }
                    else { Atencion.Parent = null; MessageBox.Show("No tiene acceso a este modulo"); }
                    Atencion.Parent = Base;
                }
                else
                {
                    MessageBox.Show("Se ha bloqueado la aplicacion con cifrado nivel" + acceso1); Atencion.Parent = null;

                }
            }

            if (e.TabPage == Proyectos)
            {
                if (acceso1 == 0)
                {
                    if (Departamento == "Proyectos" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                    {
                        MessageBox.Show("Acceso concedido");

                    }
                    else { Proyectos.Parent = null; MessageBox.Show("No tiene acceso a este modulo"); }
                    Proyectos.Parent = Base;
                }
                else { MessageBox.Show("Se ha bloqueado la aplicacion con cifrado nivel" + acceso1); Proyectos.Parent = null; }
            }

            if (e.TabPage == Ventas)
            {
                if (acceso1 == 0)
                {
                    if (Departamento == "Ventas" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                    {
                        MessageBox.Show("Acceso concedido");

                    }
                    else { Ventas.Parent = null; MessageBox.Show("No tiene acceso a este modulo"); }
                    Ventas.Parent = Base;
                }
                else { MessageBox.Show("Se ha bloqueado la aplicacion con cifrado nivel" + acceso1); Ventas.Parent = null; }

            }
            if (e.TabPage == Produccion)
            {
                if (acceso1 == 0)
                {
                    if (Departamento == "Produccion" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                    {
                        MessageBox.Show("Acceso concedido");

                    }
                    else { Produccion.Parent = null; MessageBox.Show("No tiene acceso a este modulo"); }
                    Produccion.Parent = Base;
                }
                else { MessageBox.Show("Se ha bloqueado la aplicacion con cifrado nivel" + acceso1); Produccion.Parent = null; }
            }

            if (e.TabPage == Infraes)
            {
                if (acceso1 == 0)
                {
                    if (Departamento == "Infraestructura" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                    {
                        MessageBox.Show("Acceso concedido");


                    }
                    else { Infraes.Parent = null; MessageBox.Show("No tiene acceso a este modulo"); }
                    Infraes.Parent = Base;
                }
                else { MessageBox.Show("Se ha bloqueado la aplicacion con cifrado nivel " + acceso1); Infraes.Parent = null; }
            }

            if (e.TabPage == Almacen)
            {
                if (acceso1 == 0)
                {
                    if (Departamento == "Almacen" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                    {
                        MessageBox.Show("Acceso concedido");
                    }
                    else { Almacen.Parent = null; MessageBox.Show("No tiene acceso a este modulo"); }
                    Almacen.Parent = Base;
                }
                else { MessageBox.Show("Se ha bloqueado la aplicacion con cifrado nivel " + acceso1); Almacen.Parent = null; }
            }

            if (e.TabPage == OpyMan)
            {
                if (acceso1 == 0)
                {
                    if (Departamento == "Operacion y Mantenimiento" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                    {
                        MessageBox.Show("Acceso concedido");
                    }
                    else { OpyMan.Parent = null; MessageBox.Show("No tiene acceso a este modulo"); }
                    OpyMan.Parent = Base;
                }
                else { MessageBox.Show("Se ha bloqueado la aplicacion con cifrado nivel " + acceso1); OpyMan.Parent = null; }
            }

            if (e.TabPage == ConsIns)
            {
                if (acceso1 == 0)
                {
                    if (Departamento == "Construccion e Instalaciones" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                    {
                        MessageBox.Show("Acceso concedido");
                    }
                    else { ConsIns.Parent = null; MessageBox.Show("No tiene acceso a este modulo"); }
                    ConsIns.Parent = Base;
                }
                else { MessageBox.Show("Se ha bloqueado la aplicacion con cifrado nivel " + acceso1); ConsIns.Parent = null; }
            }

            if (e.TabPage == Configuraciones)
            {
                if (acceso1 == 0)
                {
                    if (Permiso == "1" || Permiso == "2")
                    {
                        MessageBox.Show("Acceso concedido");


                    }
                    else { Configuraciones.Parent = null; MessageBox.Show("No tiene acceso a este modulo"); }
                    Configuraciones.Parent = Base;

                }
                else { MessageBox.Show("Se ha bloqueado la aplicacion con cifrado nivel " + acceso1); Configuraciones.Parent = null; }
            }


            if (e.TabPage == Administradores)
            {
                if (acceso1 == 0)
                {
                    if (Permiso == "1" || Permiso == "2")
                    {
                        MessageBox.Show("Acceso concedido");


                    }
                    else { Administradores.Parent = null; MessageBox.Show("No tiene acceso a este modulo"); }
                    Administradores.Parent = Base;

                }
                else { MessageBox.Show("Se ha bloqueado la aplicacion con cifrado nivel " + acceso1); Administradores.Parent = null; }
            }




        }

        private void noConformidadToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void noConformidadesToolStripMenuItem5_Click(object sender, EventArgs e)
        {

        }

        private void consultaToolStripMenuItem5_Click(object sender, EventArgs e)
        {


            if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
            {

                if (!Central.Controls.Contains(Consulta_No_Conf.Instance))
                {
                    Ventana = "Pantalla: Consulta de No Conformidad";
                    Consulta_No_Conf.Instance.Departamento = Departamento;
                    Central.Controls.Add(Consulta_No_Conf.Instance);
                    Consulta_No_Conf.Instance.Dock = DockStyle.Fill;
                    Consulta_No_Conf.Instance.BringToFront();


                }
                else
                {
                    Consulta_No_Conf.Instance.BringToFront();
                    Ventana = "Pantalla: Consulta de No Conformidad";
                }
            }
            else
            {
                MessageBox.Show("No tiene acceso a este modulo");
            }
        }

        private void seguimientoToolStripMenuItem5_Click(object sender, EventArgs e)
        {

        }

        private void toolStripMenuItem27_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Departamento == "Administracion" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                {

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central.Controls.Contains(Genera_AM.Instance))
                        {
                            Ventana = "Pantalla: Genera Oportunidad de Mejora";
                            Genera_AM.Instance.Departamento = Departamento;
                            Genera_AM.Instance.Nombre = Nombre;
                            Genera_AM.Instance.Ventana1 = Ventana;
                            Genera_AM.Instance.NumeroAC = NumeroOport;
                            Central.Controls.Add(Genera_AM.Instance);
                            Genera_AM.Instance.Dock = DockStyle.Fill;
                            Genera_AM.Instance.BringToFront();


                        }
                        else
                        {
                            Genera_AM.Instance.BringToFront();
                            Ventana = "Pantalla: Genera Oportunidad de Mejora";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }

                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                            conexionvpn = 0;
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {
                        //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);
                    }
                    oNetDrive = null;
                    if (Departamento == "Administracion" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                    {

                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                        {

                            if (!Central.Controls.Contains(Genera_AM.Instance))
                            {
                                Ventana = "Pantalla: Genera Oportunidad de Mejora";
                                Genera_AM.Instance.Departamento = Departamento;
                                Genera_AM.Instance.Nombre = Nombre;
                                Genera_AM.Instance.Ventana1 = Ventana;
                                Genera_AM.Instance.NumeroAC = NumeroOport;
                                Central.Controls.Add(Genera_AM.Instance);
                                Genera_AM.Instance.Dock = DockStyle.Fill;
                                Genera_AM.Instance.BringToFront();


                            }
                            else
                            {
                                Genera_AM.Instance.BringToFront();
                                Ventana = "Pantalla: Genera Oportunidad de Mejora";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                    }

                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }



        }

        private void toolStripMenuItem33_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;

                if (Departamento == "Compras" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                {

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central.Controls.Contains(Genera_AM.Instance))
                        {
                            Ventana = "Pantalla: Genera Oportunidad de Mejora";
                            Genera_AM.Instance.Departamento = Departamento;
                            Genera_AM.Instance.Nombre = Nombre;
                            Genera_AM.Instance.Ventana1 = Ventana;
                            Genera_AM.Instance.NumeroAC = NumeroOport;
                            Central.Controls.Add(Genera_AM.Instance);
                            Genera_AM.Instance.Dock = DockStyle.Fill;
                            Genera_AM.Instance.BringToFront();


                        }
                        else
                        {
                            Genera_AM.Instance.BringToFront();
                            Ventana = "Pantalla: Genera Oportunidad de Mejora";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }

                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {
                        //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);   
                    }
                    oNetDrive = null;

                    if (Departamento == "Compras" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                    {

                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                        {

                            if (!Central.Controls.Contains(Genera_AM.Instance))
                            {
                                Ventana = "Pantalla: Genera Oportunidad de Mejora";
                                Genera_AM.Instance.Departamento = Departamento;
                                Genera_AM.Instance.Nombre = Nombre;
                                Genera_AM.Instance.Ventana1 = Ventana;
                                Genera_AM.Instance.NumeroAC = NumeroOport;
                                Central.Controls.Add(Genera_AM.Instance);
                                Genera_AM.Instance.Dock = DockStyle.Fill;
                                Genera_AM.Instance.BringToFront();


                            }
                            else
                            {
                                Genera_AM.Instance.BringToFront();
                                Ventana = "Pantalla: Genera Oportunidad de Mejora";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                    }

                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void toolStripMenuItem38_Click(object sender, EventArgs e)
        {

        }

        private void toolStripMenuItem43_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;

                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central2.Controls.Contains(Genera_AM.Instance))
                    {
                        Ventana = "Pantalla: Genera Oportunidad de Mejora";
                        Genera_AM.Instance.Departamento = Departamento;
                        Genera_AM.Instance.Nombre = Nombre;
                        Genera_AM.Instance.Ventana1 = Ventana;
                        Genera_AM.Instance.NumeroAC = NumeroOport;
                        Central2.Controls.Add(Genera_AM.Instance);
                        Genera_AM.Instance.Dock = DockStyle.Fill;
                        Genera_AM.Instance.BringToFront();


                    }
                    else
                    {
                        Genera_AM.Instance.BringToFront();
                        Ventana = "Pantalla: Genera Oportunidad de Mejora";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {
                        //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central2.Controls.Contains(Genera_AM.Instance))
                        {
                            Ventana = "Pantalla: Genera Oportunidad de Mejora";
                            Genera_AM.Instance.Departamento = Departamento;
                            Genera_AM.Instance.Nombre = Nombre;
                            Genera_AM.Instance.Ventana1 = Ventana;
                            Genera_AM.Instance.NumeroAC = NumeroOport;
                            Central2.Controls.Add(Genera_AM.Instance);
                            Genera_AM.Instance.Dock = DockStyle.Fill;
                            Genera_AM.Instance.BringToFront();


                        }
                        else
                        {
                            Genera_AM.Instance.BringToFront();
                            Ventana = "Pantalla: Genera Oportunidad de Mejora";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }









        }

        private void toolStripMenuItem48_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;

                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central3.Controls.Contains(Genera_AM.Instance))
                    {
                        Ventana = "Pantalla: Genera Oportunidad de Mejora";
                        Genera_AM.Instance.Departamento = Departamento;
                        Genera_AM.Instance.Nombre = Nombre;
                        Genera_AM.Instance.Ventana1 = Ventana;
                        Genera_AM.Instance.NumeroAC = NumeroOport;
                        Central3.Controls.Add(Genera_AM.Instance);
                        Genera_AM.Instance.Dock = DockStyle.Fill;
                        Genera_AM.Instance.BringToFront();


                    }
                    else
                    {
                        Genera_AM.Instance.BringToFront();
                        Ventana = "Pantalla: Genera Oportunidad de Mejora";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {

                        //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);   
                    }
                    oNetDrive = null;

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central3.Controls.Contains(Genera_AM.Instance))
                        {
                            Ventana = "Pantalla: Genera Oportunidad de Mejora";
                            Genera_AM.Instance.Departamento = Departamento;
                            Genera_AM.Instance.Nombre = Nombre;
                            Genera_AM.Instance.Ventana1 = Ventana;
                            Genera_AM.Instance.NumeroAC = NumeroOport;
                            Central3.Controls.Add(Genera_AM.Instance);
                            Genera_AM.Instance.Dock = DockStyle.Fill;
                            Genera_AM.Instance.BringToFront();


                        }
                        else
                        {
                            Genera_AM.Instance.BringToFront();
                            Ventana = "Pantalla: Genera Oportunidad de Mejora";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }








        }

        private void toolStripMenuItem53_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;


                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central4.Controls.Contains(Genera_AM.Instance))
                    {
                        Ventana = "Pantalla: Genera Oportunidad de Mejora";
                        Genera_AM.Instance.Departamento = Departamento;
                        Genera_AM.Instance.Nombre = Nombre;
                        Genera_AM.Instance.Ventana1 = Ventana;
                        Genera_AM.Instance.NumeroAC = NumeroOport;
                        Central4.Controls.Add(Genera_AM.Instance);
                        Genera_AM.Instance.Dock = DockStyle.Fill;
                        Genera_AM.Instance.BringToFront();


                    }
                    else
                    {
                        Genera_AM.Instance.BringToFront();
                        Ventana = "Pantalla: Genera Oportunidad de Mejora";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {

                        //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central4.Controls.Contains(Genera_AM.Instance))
                        {
                            Ventana = "Pantalla: Genera Oportunidad de Mejora";
                            Genera_AM.Instance.Departamento = Departamento;
                            Genera_AM.Instance.Nombre = Nombre;
                            Genera_AM.Instance.Ventana1 = Ventana;
                            Genera_AM.Instance.NumeroAC = NumeroOport;
                            Central4.Controls.Add(Genera_AM.Instance);
                            Genera_AM.Instance.Dock = DockStyle.Fill;
                            Genera_AM.Instance.BringToFront();


                        }
                        else
                        {
                            Genera_AM.Instance.BringToFront();
                            Ventana = "Pantalla: Genera Oportunidad de Mejora";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void toolStripMenuItem58_Click(object sender, EventArgs e)
        {

            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central5.Controls.Contains(Genera_AM.Instance))
                    {
                        Ventana = "Pantalla: Genera Oportunidad de Mejora";
                        Genera_AM.Instance.Departamento = Departamento;
                        Genera_AM.Instance.Nombre = Nombre;
                        Genera_AM.Instance.Ventana1 = Ventana;
                        Genera_AM.Instance.NumeroAC = NumeroOport;
                        Central5.Controls.Add(Genera_AM.Instance);
                        Genera_AM.Instance.Dock = DockStyle.Fill;
                        Genera_AM.Instance.BringToFront();


                    }
                    else
                    {
                        Genera_AM.Instance.BringToFront();
                        Ventana = "Pantalla: Genera Oportunidad de Mejora";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {

                        //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central5.Controls.Contains(Genera_AM.Instance))
                        {
                            Ventana = "Pantalla: Genera Oportunidad de Mejora";
                            Genera_AM.Instance.Departamento = Departamento;
                            Genera_AM.Instance.Nombre = Nombre;
                            Genera_AM.Instance.Ventana1 = Ventana;
                            Genera_AM.Instance.NumeroAC = NumeroOport;
                            Central5.Controls.Add(Genera_AM.Instance);
                            Genera_AM.Instance.Dock = DockStyle.Fill;
                            Genera_AM.Instance.BringToFront();


                        }
                        else
                        {
                            Genera_AM.Instance.BringToFront();
                            Ventana = "Pantalla: Genera Oportunidad de Mejora";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void toolStripMenuItem63_Click(object sender, EventArgs e)
        {

            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central6.Controls.Contains(Genera_AM.Instance))
                    {
                        Ventana = "Pantalla: Genera Oportunidad de Mejora";
                        Genera_AM.Instance.Departamento = Departamento;
                        Genera_AM.Instance.Nombre = Nombre;
                        Genera_AM.Instance.Ventana1 = Ventana;
                        Genera_AM.Instance.NumeroAC = NumeroOport;
                        Central6.Controls.Add(Genera_AM.Instance);
                        Genera_AM.Instance.Dock = DockStyle.Fill;
                        Genera_AM.Instance.BringToFront();


                    }
                    else
                    {
                        Genera_AM.Instance.BringToFront();
                        Ventana = "Pantalla: Genera Oportunidad de Mejora";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {

                        //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);   
                    }
                    oNetDrive = null;
                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central6.Controls.Contains(Genera_AM.Instance))
                        {
                            Ventana = "Pantalla: Genera Oportunidad de Mejora";
                            Genera_AM.Instance.Departamento = Departamento;
                            Genera_AM.Instance.Nombre = Nombre;
                            Genera_AM.Instance.Ventana1 = Ventana;
                            Genera_AM.Instance.NumeroAC = NumeroOport;
                            Central6.Controls.Add(Genera_AM.Instance);
                            Genera_AM.Instance.Dock = DockStyle.Fill;
                            Genera_AM.Instance.BringToFront();


                        }
                        else
                        {
                            Genera_AM.Instance.BringToFront();
                            Ventana = "Pantalla: Genera Oportunidad de Mejora";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void toolStripMenuItem68_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;


                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central7.Controls.Contains(Genera_AM.Instance))
                    {
                        Ventana = "Pantalla: Genera Oportunidad de Mejora";
                        Genera_AM.Instance.Departamento = Departamento;
                        Genera_AM.Instance.Nombre = Nombre;
                        Genera_AM.Instance.Ventana1 = Ventana;
                        Genera_AM.Instance.NumeroAC = NumeroOport;
                        Central7.Controls.Add(Genera_AM.Instance);
                        Genera_AM.Instance.Dock = DockStyle.Fill;
                        Genera_AM.Instance.BringToFront();


                    }
                    else
                    {
                        Genera_AM.Instance.BringToFront();
                        Ventana = "Pantalla: Genera Oportunidad de Mejora";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;


                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {// MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);   
                    }
                    oNetDrive = null;


                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central7.Controls.Contains(Genera_AM.Instance))
                        {
                            Ventana = "Pantalla: Genera Oportunidad de Mejora";
                            Genera_AM.Instance.Departamento = Departamento;
                            Genera_AM.Instance.Nombre = Nombre;
                            Genera_AM.Instance.Ventana1 = Ventana;
                            Genera_AM.Instance.NumeroAC = NumeroOport;
                            Central7.Controls.Add(Genera_AM.Instance);
                            Genera_AM.Instance.Dock = DockStyle.Fill;
                            Genera_AM.Instance.BringToFront();


                        }
                        else
                        {
                            Genera_AM.Instance.BringToFront();
                            Ventana = "Pantalla: Genera Oportunidad de Mejora";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }

        }

        private void toolStripMenuItem73_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "AUXILIAR DE  PRODUCCION")
                {

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central8.Controls.Contains(Genera_AM.Instance))
                        {
                            Ventana = "Pantalla: Genera Oportunidad de Mejora";
                            Genera_AM.Instance.Departamento = Departamento;
                            Genera_AM.Instance.Nombre = Nombre;
                            Genera_AM.Instance.Ventana1 = Ventana;
                            Genera_AM.Instance.NumeroAC = NumeroOport;
                            Central8.Controls.Add(Genera_AM.Instance);
                            Genera_AM.Instance.Dock = DockStyle.Fill;
                            Genera_AM.Instance.BringToFront();


                        }
                        else
                        {
                            Genera_AM.Instance.BringToFront();
                            Ventana = "Pantalla: Genera Oportunidad de Mejora";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else { MessageBox.Show("No tiene acceso a este modulo"); }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {

                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;
                    if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "AUXILIAR DE  PRODUCCION")
                    {

                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                        {

                            if (!Central8.Controls.Contains(Genera_AM.Instance))
                            {
                                Ventana = "Pantalla: Genera Oportunidad de Mejora";
                                Genera_AM.Instance.Departamento = Departamento;
                                Genera_AM.Instance.Nombre = Nombre;
                                Genera_AM.Instance.Ventana1 = Ventana;
                                Genera_AM.Instance.NumeroAC = NumeroOport;
                                Central8.Controls.Add(Genera_AM.Instance);
                                Genera_AM.Instance.Dock = DockStyle.Fill;
                                Genera_AM.Instance.BringToFront();


                            }
                            else
                            {
                                Genera_AM.Instance.BringToFront();
                                Ventana = "Pantalla: Genera Oportunidad de Mejora";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                    }
                    else { MessageBox.Show("No tiene acceso a este modulo"); }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void toolStripMenuItem28_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Departamento == "Administracion" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                {

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central.Controls.Contains(Seguimiento_AM.Instance))
                        {
                            Ventana = "Pantalla: Detalle No conformidad";
                            Seguimiento_AM.Instance.Nombre = Nombre;
                            Seguimiento_AM.Instance.Apellido = Apellido;
                            Seguimiento_AM.Instance.Departamento = Departamento;
                            Seguimiento_AM.Instance.Ventana1 = Ventana;
                            Central.Controls.Add(Seguimiento_AM.Instance);
                            Seguimiento_AM.Instance.Dock = DockStyle.Fill;
                            Seguimiento_AM.Instance.BringToFront();


                        }
                        else
                        {
                            Seguimiento_AM.Instance.BringToFront();
                            Ventana = "Pantalla: Detalle No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }

                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                            conexionvpn = 0;
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); 
                    }
                    oNetDrive = null;
                    if (Departamento == "Administracion" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                    {

                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                        {

                            if (!Central.Controls.Contains(Seguimiento_AM.Instance))
                            {
                                Ventana = "Pantalla: Detalle No conformidad";
                                Seguimiento_AM.Instance.Nombre = Nombre;
                                Seguimiento_AM.Instance.Apellido = Apellido;
                                Seguimiento_AM.Instance.Departamento = Departamento;
                                Seguimiento_AM.Instance.Ventana1 = Ventana;
                                Central.Controls.Add(Seguimiento_AM.Instance);
                                Seguimiento_AM.Instance.Dock = DockStyle.Fill;
                                Seguimiento_AM.Instance.BringToFront();


                            }
                            else
                            {
                                Seguimiento_AM.Instance.BringToFront();
                                Ventana = "Pantalla: Detalle No conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                    }

                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }



        }

        private void toolStripMenuItem34_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;

                if (Departamento == "Compras" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                {

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central.Controls.Contains(Seguimiento_AM.Instance))
                        {
                            Ventana = "Pantalla: Detalle No conformidad";
                            Seguimiento_AM.Instance.Nombre = Nombre;
                            Seguimiento_AM.Instance.Apellido = Apellido;
                            Seguimiento_AM.Instance.Departamento = Departamento;
                            Seguimiento_AM.Instance.Ventana1 = Ventana;
                            Central.Controls.Add(Seguimiento_AM.Instance);
                            Seguimiento_AM.Instance.Dock = DockStyle.Fill;
                            Seguimiento_AM.Instance.BringToFront();


                        }
                        else
                        {
                            Seguimiento_AM.Instance.BringToFront();
                            Ventana = "Pantalla: Detalle No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;

                    if (Departamento == "Compras" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                    {

                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                        {

                            if (!Central.Controls.Contains(Seguimiento_AM.Instance))
                            {
                                Ventana = "Pantalla: Detalle No conformidad";
                                Seguimiento_AM.Instance.Nombre = Nombre;
                                Seguimiento_AM.Instance.Apellido = Apellido;
                                Seguimiento_AM.Instance.Departamento = Departamento;
                                Seguimiento_AM.Instance.Ventana1 = Ventana;
                                Central.Controls.Add(Seguimiento_AM.Instance);
                                Seguimiento_AM.Instance.Dock = DockStyle.Fill;
                                Seguimiento_AM.Instance.BringToFront();


                            }
                            else
                            {
                                Seguimiento_AM.Instance.BringToFront();
                                Ventana = "Pantalla: Detalle No conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }

                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }










        }

        private void toolStripMenuItem39_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Departamento == "Cobranza" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                {

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {


                        if (!Central.Controls.Contains(Seguimiento_AM.Instance))
                        {
                            Ventana = "Pantalla: Detalle No conformidad";
                            Seguimiento_AM.Instance.Nombre = Nombre;
                            Seguimiento_AM.Instance.Apellido = Apellido;
                            Seguimiento_AM.Instance.Departamento = Departamento;
                            Seguimiento_AM.Instance.Ventana1 = Ventana;
                            Central.Controls.Add(Seguimiento_AM.Instance);
                            Seguimiento_AM.Instance.Dock = DockStyle.Fill;
                            Seguimiento_AM.Instance.BringToFront();


                        }
                        else
                        {
                            Seguimiento_AM.Instance.BringToFront();
                            Ventana = "Pantalla: Detalle No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {

                        //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); 
                    }
                    oNetDrive = null;
                    if (Departamento == "Cobranza" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                    {

                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                        {


                            if (!Central.Controls.Contains(Seguimiento_AM.Instance))
                            {
                                Ventana = "Pantalla: Detalle No conformidad";
                                Seguimiento_AM.Instance.Nombre = Nombre;
                                Seguimiento_AM.Instance.Apellido = Apellido;
                                Seguimiento_AM.Instance.Departamento = Departamento;
                                Seguimiento_AM.Instance.Ventana1 = Ventana;
                                Central.Controls.Add(Seguimiento_AM.Instance);
                                Seguimiento_AM.Instance.Dock = DockStyle.Fill;
                                Seguimiento_AM.Instance.BringToFront();


                            }
                            else
                            {
                                Seguimiento_AM.Instance.BringToFront();
                                Ventana = "Pantalla: Detalle No conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }



        }

        private void toolStripMenuItem44_Click(object sender, EventArgs e)
        {

            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central2.Controls.Contains(Seguimiento_AM.Instance))
                    {
                        Ventana = "Pantalla: Detalle No conformidad";
                        Seguimiento_AM.Instance.Nombre = Nombre;
                        Seguimiento_AM.Instance.Apellido = Apellido;
                        Seguimiento_AM.Instance.Departamento = Departamento;
                        Seguimiento_AM.Instance.Ventana1 = Ventana;
                        Central2.Controls.Add(Seguimiento_AM.Instance);
                        Seguimiento_AM.Instance.Dock = DockStyle.Fill;
                        Seguimiento_AM.Instance.BringToFront();


                    }
                    else
                    {
                        Seguimiento_AM.Instance.BringToFront();
                        Ventana = "Pantalla: Detalle No conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //    MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);   
                    }
                    oNetDrive = null;
                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central2.Controls.Contains(Seguimiento_AM.Instance))
                        {
                            Ventana = "Pantalla: Detalle No conformidad";
                            Seguimiento_AM.Instance.Nombre = Nombre;
                            Seguimiento_AM.Instance.Apellido = Apellido;
                            Seguimiento_AM.Instance.Departamento = Departamento;
                            Seguimiento_AM.Instance.Ventana1 = Ventana;
                            Central2.Controls.Add(Seguimiento_AM.Instance);
                            Seguimiento_AM.Instance.Dock = DockStyle.Fill;
                            Seguimiento_AM.Instance.BringToFront();


                        }
                        else
                        {
                            Seguimiento_AM.Instance.BringToFront();
                            Ventana = "Pantalla: Detalle No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }









        }

        private void toolStripMenuItem49_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;




                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central3.Controls.Contains(Seguimiento_AM.Instance))
                    {
                        Ventana = "Pantalla: Detalle No conformidad";
                        Seguimiento_AM.Instance.Nombre = Nombre;
                        Seguimiento_AM.Instance.Apellido = Apellido;
                        Seguimiento_AM.Instance.Departamento = Departamento;
                        Seguimiento_AM.Instance.Ventana1 = Ventana;
                        Central3.Controls.Add(Seguimiento_AM.Instance);
                        Seguimiento_AM.Instance.Dock = DockStyle.Fill;
                        Seguimiento_AM.Instance.BringToFront();


                    }
                    else
                    {
                        Seguimiento_AM.Instance.BringToFront();
                        Ventana = "Pantalla: Detalle No conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {

                        //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);   
                    }
                    oNetDrive = null;




                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central3.Controls.Contains(Seguimiento_AM.Instance))
                        {
                            Ventana = "Pantalla: Detalle No conformidad";
                            Seguimiento_AM.Instance.Nombre = Nombre;
                            Seguimiento_AM.Instance.Apellido = Apellido;
                            Seguimiento_AM.Instance.Departamento = Departamento;
                            Seguimiento_AM.Instance.Ventana1 = Ventana;
                            Central3.Controls.Add(Seguimiento_AM.Instance);
                            Seguimiento_AM.Instance.Dock = DockStyle.Fill;
                            Seguimiento_AM.Instance.BringToFront();


                        }
                        else
                        {
                            Seguimiento_AM.Instance.BringToFront();
                            Ventana = "Pantalla: Detalle No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }






        }

        private void toolStripMenuItem54_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;

                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central4.Controls.Contains(Seguimiento_AM.Instance))
                    {
                        Ventana = "Pantalla: Detalle No conformidad";
                        Seguimiento_AM.Instance.Nombre = Nombre;
                        Seguimiento_AM.Instance.Apellido = Apellido;
                        Seguimiento_AM.Instance.Departamento = Departamento;
                        Seguimiento_AM.Instance.Ventana1 = Ventana;
                        Central4.Controls.Add(Seguimiento_AM.Instance);
                        Seguimiento_AM.Instance.Dock = DockStyle.Fill;
                        Seguimiento_AM.Instance.BringToFront();


                    }
                    else
                    {
                        Seguimiento_AM.Instance.BringToFront();
                        Ventana = "Pantalla: Detalle No conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central4.Controls.Contains(Seguimiento_AM.Instance))
                        {
                            Ventana = "Pantalla: Detalle No conformidad";
                            Seguimiento_AM.Instance.Nombre = Nombre;
                            Seguimiento_AM.Instance.Apellido = Apellido;
                            Seguimiento_AM.Instance.Departamento = Departamento;
                            Seguimiento_AM.Instance.Ventana1 = Ventana;
                            Central4.Controls.Add(Seguimiento_AM.Instance);
                            Seguimiento_AM.Instance.Dock = DockStyle.Fill;
                            Seguimiento_AM.Instance.BringToFront();


                        }
                        else
                        {
                            Seguimiento_AM.Instance.BringToFront();
                            Ventana = "Pantalla: Detalle No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void toolStripMenuItem59_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;


                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central5.Controls.Contains(Seguimiento_AM.Instance))
                    {
                        Ventana = "Pantalla: Detalle No conformidad";
                        Seguimiento_AM.Instance.Nombre = Nombre;
                        Seguimiento_AM.Instance.Apellido = Apellido;
                        Seguimiento_AM.Instance.Departamento = Departamento;
                        Seguimiento_AM.Instance.Ventana1 = Ventana;
                        Central5.Controls.Add(Seguimiento_AM.Instance);
                        Seguimiento_AM.Instance.Dock = DockStyle.Fill;
                        Seguimiento_AM.Instance.BringToFront();


                    }
                    else
                    {
                        Seguimiento_AM.Instance.BringToFront();
                        Ventana = "Pantalla: Detalle No conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //    MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);   
                    }
                    oNetDrive = null;


                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central5.Controls.Contains(Seguimiento_AM.Instance))
                        {
                            Ventana = "Pantalla: Detalle No conformidad";
                            Seguimiento_AM.Instance.Nombre = Nombre;
                            Seguimiento_AM.Instance.Apellido = Apellido;
                            Seguimiento_AM.Instance.Departamento = Departamento;
                            Seguimiento_AM.Instance.Ventana1 = Ventana;
                            Central5.Controls.Add(Seguimiento_AM.Instance);
                            Seguimiento_AM.Instance.Dock = DockStyle.Fill;
                            Seguimiento_AM.Instance.BringToFront();


                        }
                        else
                        {
                            Seguimiento_AM.Instance.BringToFront();
                            Ventana = "Pantalla: Detalle No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }

        }

        private void toolStripMenuItem64_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central6.Controls.Contains(Seguimiento_AM.Instance))
                    {
                        Ventana = "Pantalla: Detalle No conformidad";
                        Seguimiento_AM.Instance.Nombre = Nombre;
                        Seguimiento_AM.Instance.Apellido = Apellido;
                        Seguimiento_AM.Instance.Departamento = Departamento;
                        Seguimiento_AM.Instance.Ventana1 = Ventana;
                        Central6.Controls.Add(Seguimiento_AM.Instance);
                        Seguimiento_AM.Instance.Dock = DockStyle.Fill;
                        Seguimiento_AM.Instance.BringToFront();


                    }
                    else
                    {
                        Seguimiento_AM.Instance.BringToFront();
                        Ventana = "Pantalla: Detalle No conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            // MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;


                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); 
                    }
                    oNetDrive = null;
                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central6.Controls.Contains(Seguimiento_AM.Instance))
                        {
                            Ventana = "Pantalla: Detalle No conformidad";
                            Seguimiento_AM.Instance.Nombre = Nombre;
                            Seguimiento_AM.Instance.Apellido = Apellido;
                            Seguimiento_AM.Instance.Departamento = Departamento;
                            Seguimiento_AM.Instance.Ventana1 = Ventana;
                            Central6.Controls.Add(Seguimiento_AM.Instance);
                            Seguimiento_AM.Instance.Dock = DockStyle.Fill;
                            Seguimiento_AM.Instance.BringToFront();


                        }
                        else
                        {
                            Seguimiento_AM.Instance.BringToFront();
                            Ventana = "Pantalla: Detalle No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }



        }

        private void toolStripMenuItem69_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central7.Controls.Contains(Seguimiento_AM.Instance))
                    {
                        Ventana = "Pantalla: Detalle No conformidad";
                        Seguimiento_AM.Instance.Nombre = Nombre;
                        Seguimiento_AM.Instance.Apellido = Apellido;
                        Seguimiento_AM.Instance.Departamento = Departamento;
                        Seguimiento_AM.Instance.Ventana1 = Ventana;
                        Central7.Controls.Add(Seguimiento_AM.Instance);
                        Seguimiento_AM.Instance.Dock = DockStyle.Fill;
                        Seguimiento_AM.Instance.BringToFront();


                    }
                    else
                    {
                        Seguimiento_AM.Instance.BringToFront();
                        Ventana = "Pantalla: Detalle No conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;
                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central7.Controls.Contains(Seguimiento_AM.Instance))
                        {
                            Ventana = "Pantalla: Detalle No conformidad";
                            Seguimiento_AM.Instance.Nombre = Nombre;
                            Seguimiento_AM.Instance.Apellido = Apellido;
                            Seguimiento_AM.Instance.Departamento = Departamento;
                            Seguimiento_AM.Instance.Ventana1 = Ventana;
                            Central7.Controls.Add(Seguimiento_AM.Instance);
                            Seguimiento_AM.Instance.Dock = DockStyle.Fill;
                            Seguimiento_AM.Instance.BringToFront();


                        }
                        else
                        {
                            Seguimiento_AM.Instance.BringToFront();
                            Ventana = "Pantalla: Detalle No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }



        }

        private void toolStripMenuItem74_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "AUXILIAR DE  PRODUCCION")
                {

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central8.Controls.Contains(Seguimiento_AM.Instance))
                        {
                            Ventana = "Pantalla: Detalle No conformidad";
                            Seguimiento_AM.Instance.Nombre = Nombre;
                            Seguimiento_AM.Instance.Apellido = Apellido;
                            Seguimiento_AM.Instance.Departamento = Departamento;
                            Seguimiento_AM.Instance.Ventana1 = Ventana;
                            Central8.Controls.Add(Seguimiento_AM.Instance);
                            Seguimiento_AM.Instance.Dock = DockStyle.Fill;
                            Seguimiento_AM.Instance.BringToFront();


                        }
                        else
                        {
                            Seguimiento_AM.Instance.BringToFront();
                            Ventana = "Pantalla: Detalle No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else { MessageBox.Show("No tiene acceso a este modulo"); }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {

                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //     MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); 
                    }
                    oNetDrive = null;
                    if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "AUXILIAR DE  PRODUCCION")
                    {

                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                        {

                            if (!Central8.Controls.Contains(Seguimiento_AM.Instance))
                            {
                                Ventana = "Pantalla: Detalle No conformidad";
                                Seguimiento_AM.Instance.Nombre = Nombre;
                                Seguimiento_AM.Instance.Apellido = Apellido;
                                Seguimiento_AM.Instance.Departamento = Departamento;
                                Seguimiento_AM.Instance.Ventana1 = Ventana;
                                Central8.Controls.Add(Seguimiento_AM.Instance);
                                Seguimiento_AM.Instance.Dock = DockStyle.Fill;
                                Seguimiento_AM.Instance.BringToFront();


                            }
                            else
                            {
                                Seguimiento_AM.Instance.BringToFront();
                                Ventana = "Pantalla: Detalle No conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                    }
                    else { MessageBox.Show("No tiene acceso a este modulo"); }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void toolStripMenuItem31_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Departamento == "Administracion" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                {

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central.Controls.Contains(Consulta_OportunidadesMejora_General.Instance))
                        {
                            Ventana = "Pantalla: Detalle No conformidad";

                            Central.Controls.Add(Consulta_OportunidadesMejora_General.Instance);
                            Consulta_OportunidadesMejora_General.Instance.Dock = DockStyle.Fill;
                            Consulta_OportunidadesMejora_General.Instance.BringToFront();


                        }
                        else
                        {
                            Consulta_OportunidadesMejora_General.Instance.BringToFront();
                            Ventana = "Pantalla: Detalle No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //    MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                            conexionvpn = 0;
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;

                    if (Departamento == "Administracion" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                    {

                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                        {

                            if (!Central.Controls.Contains(Consulta_OportunidadesMejora_General.Instance))
                            {
                                Ventana = "Pantalla: Detalle No conformidad";

                                Central.Controls.Add(Consulta_OportunidadesMejora_General.Instance);
                                Consulta_OportunidadesMejora_General.Instance.Dock = DockStyle.Fill;
                                Consulta_OportunidadesMejora_General.Instance.BringToFront();


                            }
                            else
                            {
                                Consulta_OportunidadesMejora_General.Instance.BringToFront();
                                Ventana = "Pantalla: Detalle No conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void toolStripMenuItem35_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;




                if (Departamento == "Compras" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                {

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central.Controls.Contains(Consulta_OportunidadesMejora_General.Instance))
                        {
                            Ventana = "Pantalla: Detalle No conformidad";

                            Central.Controls.Add(Consulta_OportunidadesMejora_General.Instance);
                            Consulta_OportunidadesMejora_General.Instance.Dock = DockStyle.Fill;
                            Consulta_OportunidadesMejora_General.Instance.BringToFront();


                        }
                        else
                        {
                            Consulta_OportunidadesMejora_General.Instance.BringToFront();
                            Ventana = "Pantalla: Detalle No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }

                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //     MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {//  MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;




                    if (Departamento == "Compras" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                    {

                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                        {

                            if (!Central.Controls.Contains(Consulta_OportunidadesMejora_General.Instance))
                            {
                                Ventana = "Pantalla: Detalle No conformidad";

                                Central.Controls.Add(Consulta_OportunidadesMejora_General.Instance);
                                Consulta_OportunidadesMejora_General.Instance.Dock = DockStyle.Fill;
                                Consulta_OportunidadesMejora_General.Instance.BringToFront();


                            }
                            else
                            {
                                Consulta_OportunidadesMejora_General.Instance.BringToFront();
                                Ventana = "Pantalla: Detalle No conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                    }

                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }









            if (Departamento == "Compras" || Departamento == "SuperAdmin" || Departamento == "Direccion")
            {

                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central.Controls.Contains(Consulta_OportunidadesMejora_General.Instance))
                    {
                        Ventana = "Pantalla: Detalle No conformidad";

                        Central.Controls.Add(Consulta_OportunidadesMejora_General.Instance);
                        Consulta_OportunidadesMejora_General.Instance.Dock = DockStyle.Fill;
                        Consulta_OportunidadesMejora_General.Instance.BringToFront();


                    }
                    else
                    {
                        Consulta_OportunidadesMejora_General.Instance.BringToFront();
                        Ventana = "Pantalla: Detalle No conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }

            else
            {
                MessageBox.Show("No tiene acceso a este modulo");
            }

        }

        private void toolStripMenuItem40_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Departamento == "Cobranza" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                {

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {
                        if (!Central.Controls.Contains(Consulta_OportunidadesMejora_General.Instance))
                        {
                            Ventana = "Pantalla: Detalle No conformidad";

                            Central.Controls.Add(Consulta_OportunidadesMejora_General.Instance);
                            Consulta_OportunidadesMejora_General.Instance.Dock = DockStyle.Fill;
                            Consulta_OportunidadesMejora_General.Instance.BringToFront();


                        }
                        else
                        {
                            Consulta_OportunidadesMejora_General.Instance.BringToFront();
                            Ventana = "Pantalla: Detalle No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }

                }

                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {//  MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;
                    if (Departamento == "Cobranza" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                    {

                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                        {
                            if (!Central.Controls.Contains(Consulta_OportunidadesMejora_General.Instance))
                            {
                                Ventana = "Pantalla: Detalle No conformidad";

                                Central.Controls.Add(Consulta_OportunidadesMejora_General.Instance);
                                Consulta_OportunidadesMejora_General.Instance.Dock = DockStyle.Fill;
                                Consulta_OportunidadesMejora_General.Instance.BringToFront();


                            }
                            else
                            {
                                Consulta_OportunidadesMejora_General.Instance.BringToFront();
                                Ventana = "Pantalla: Detalle No conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }

                    }

                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }



        }

        private void toolStripMenuItem45_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central2.Controls.Contains(Consulta_OportunidadesMejora_General.Instance))
                    {
                        Ventana = "Pantalla: Detalle No conformidad";

                        Central2.Controls.Add(Consulta_OportunidadesMejora_General.Instance);
                        Consulta_OportunidadesMejora_General.Instance.Dock = DockStyle.Fill;
                        Consulta_OportunidadesMejora_General.Instance.BringToFront();


                    }
                    else
                    {
                        Consulta_OportunidadesMejora_General.Instance.BringToFront();
                        Ventana = "Pantalla: Detalle No conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;
                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central2.Controls.Contains(Consulta_OportunidadesMejora_General.Instance))
                        {
                            Ventana = "Pantalla: Detalle No conformidad";

                            Central2.Controls.Add(Consulta_OportunidadesMejora_General.Instance);
                            Consulta_OportunidadesMejora_General.Instance.Dock = DockStyle.Fill;
                            Consulta_OportunidadesMejora_General.Instance.BringToFront();


                        }
                        else
                        {
                            Consulta_OportunidadesMejora_General.Instance.BringToFront();
                            Ventana = "Pantalla: Detalle No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }

                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }










        }

        private void toolStripMenuItem50_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central3.Controls.Contains(Consulta_OportunidadesMejora_General.Instance))
                    {
                        Ventana = "Pantalla: Detalle No conformidad";

                        Central3.Controls.Add(Consulta_OportunidadesMejora_General.Instance);
                        Consulta_OportunidadesMejora_General.Instance.Dock = DockStyle.Fill;
                        Consulta_OportunidadesMejora_General.Instance.BringToFront();


                    }
                    else
                    {
                        Consulta_OportunidadesMejora_General.Instance.BringToFront();
                        Ventana = "Pantalla: Detalle No conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); 
                    }
                    oNetDrive = null;

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central3.Controls.Contains(Consulta_OportunidadesMejora_General.Instance))
                        {
                            Ventana = "Pantalla: Detalle No conformidad";

                            Central3.Controls.Add(Consulta_OportunidadesMejora_General.Instance);
                            Consulta_OportunidadesMejora_General.Instance.Dock = DockStyle.Fill;
                            Consulta_OportunidadesMejora_General.Instance.BringToFront();


                        }
                        else
                        {
                            Consulta_OportunidadesMejora_General.Instance.BringToFront();
                            Ventana = "Pantalla: Detalle No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }










        }

        private void toolStripMenuItem55_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central4.Controls.Contains(Consulta_OportunidadesMejora_General.Instance))
                    {
                        Ventana = "Pantalla: Detalle No conformidad";

                        Central4.Controls.Add(Consulta_OportunidadesMejora_General.Instance);
                        Consulta_OportunidadesMejora_General.Instance.Dock = DockStyle.Fill;
                        Consulta_OportunidadesMejora_General.Instance.BringToFront();


                    }
                    else
                    {
                        Consulta_OportunidadesMejora_General.Instance.BringToFront();
                        Ventana = "Pantalla: Detalle No conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //    MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); 
                    }
                    oNetDrive = null;
                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central4.Controls.Contains(Consulta_OportunidadesMejora_General.Instance))
                        {
                            Ventana = "Pantalla: Detalle No conformidad";

                            Central4.Controls.Add(Consulta_OportunidadesMejora_General.Instance);
                            Consulta_OportunidadesMejora_General.Instance.Dock = DockStyle.Fill;
                            Consulta_OportunidadesMejora_General.Instance.BringToFront();


                        }
                        else
                        {
                            Consulta_OportunidadesMejora_General.Instance.BringToFront();
                            Ventana = "Pantalla: Detalle No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }



        }

        private void toolStripMenuItem60_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;

                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central5.Controls.Contains(Consulta_OportunidadesMejora_General.Instance))
                    {
                        Ventana = "Pantalla: Detalle No conformidad";

                        Central5.Controls.Add(Consulta_OportunidadesMejora_General.Instance);
                        Consulta_OportunidadesMejora_General.Instance.Dock = DockStyle.Fill;
                        Consulta_OportunidadesMejora_General.Instance.BringToFront();


                    }
                    else
                    {
                        Consulta_OportunidadesMejora_General.Instance.BringToFront();
                        Ventana = "Pantalla: Detalle No conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);   
                    }
                    oNetDrive = null;

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central5.Controls.Contains(Consulta_OportunidadesMejora_General.Instance))
                        {
                            Ventana = "Pantalla: Detalle No conformidad";

                            Central5.Controls.Add(Consulta_OportunidadesMejora_General.Instance);
                            Consulta_OportunidadesMejora_General.Instance.Dock = DockStyle.Fill;
                            Consulta_OportunidadesMejora_General.Instance.BringToFront();


                        }
                        else
                        {
                            Consulta_OportunidadesMejora_General.Instance.BringToFront();
                            Ventana = "Pantalla: Detalle No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void toolStripMenuItem65_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;

                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central6.Controls.Contains(Consulta_OportunidadesMejora_General.Instance))
                    {
                        Ventana = "Pantalla: Detalle No conformidad";

                        Central6.Controls.Add(Consulta_OportunidadesMejora_General.Instance);
                        Consulta_OportunidadesMejora_General.Instance.Dock = DockStyle.Fill;
                        Consulta_OportunidadesMejora_General.Instance.BringToFront();


                    }
                    else
                    {
                        Consulta_OportunidadesMejora_General.Instance.BringToFront();
                        Ventana = "Pantalla: Detalle No conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); }
                    oNetDrive = null;

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central6.Controls.Contains(Consulta_OportunidadesMejora_General.Instance))
                        {
                            Ventana = "Pantalla: Detalle No conformidad";

                            Central6.Controls.Add(Consulta_OportunidadesMejora_General.Instance);
                            Consulta_OportunidadesMejora_General.Instance.Dock = DockStyle.Fill;
                            Consulta_OportunidadesMejora_General.Instance.BringToFront();


                        }
                        else
                        {
                            Consulta_OportunidadesMejora_General.Instance.BringToFront();
                            Ventana = "Pantalla: Detalle No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void toolStripMenuItem70_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central7.Controls.Contains(Consulta_OportunidadesMejora_General.Instance))
                    {
                        Ventana = "Pantalla: Detalle No conformidad";

                        Central7.Controls.Add(Consulta_OportunidadesMejora_General.Instance);
                        Consulta_OportunidadesMejora_General.Instance.Dock = DockStyle.Fill;
                        Consulta_OportunidadesMejora_General.Instance.BringToFront();


                    }
                    else
                    {
                        Consulta_OportunidadesMejora_General.Instance.BringToFront();
                        Ventana = "Pantalla: Detalle No conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);
                    }
                    oNetDrive = null;
                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central7.Controls.Contains(Consulta_OportunidadesMejora_General.Instance))
                        {
                            Ventana = "Pantalla: Detalle No conformidad";

                            Central7.Controls.Add(Consulta_OportunidadesMejora_General.Instance);
                            Consulta_OportunidadesMejora_General.Instance.Dock = DockStyle.Fill;
                            Consulta_OportunidadesMejora_General.Instance.BringToFront();


                        }
                        else
                        {
                            Consulta_OportunidadesMejora_General.Instance.BringToFront();
                            Ventana = "Pantalla: Detalle No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }



        }

        private void toolStripMenuItem71_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central7.Controls.Contains(Consulta_Oportunidades_Mejora.Instance))
                    {
                        Ventana = "Pantalla: Detalle No conformidad";


                        Consulta_Oportunidades_Mejora.Instance.Nombre = Nombre;
                        Consulta_Oportunidades_Mejora.Instance.Departamento = Departamento;
                        Consulta_Oportunidades_Mejora.Instance.Ventana1 = Ventana;

                        Central7.Controls.Add(Consulta_Oportunidades_Mejora.Instance);
                        Consulta_Oportunidades_Mejora.Instance.Dock = DockStyle.Fill;
                        Consulta_Oportunidades_Mejora.Instance.BringToFront();



                    }
                    else
                    {
                        Consulta_OportunidadesMejora_General.Instance.BringToFront();
                        Ventana = "Pantalla: Detalle No conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //    MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;
                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central7.Controls.Contains(Consulta_Oportunidades_Mejora.Instance))
                        {
                            Ventana = "Pantalla: Detalle No conformidad";


                            Consulta_Oportunidades_Mejora.Instance.Nombre = Nombre;
                            Consulta_Oportunidades_Mejora.Instance.Departamento = Departamento;
                            Consulta_Oportunidades_Mejora.Instance.Ventana1 = Ventana;

                            Central7.Controls.Add(Consulta_Oportunidades_Mejora.Instance);
                            Consulta_Oportunidades_Mejora.Instance.Dock = DockStyle.Fill;
                            Consulta_Oportunidades_Mejora.Instance.BringToFront();



                        }
                        else
                        {
                            Consulta_OportunidadesMejora_General.Instance.BringToFront();
                            Ventana = "Pantalla: Detalle No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }



        }

        private void toolStripMenuItem75_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;

                if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Departamento == "Infraestructura")
                {

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central8.Controls.Contains(Consulta_OportunidadesMejora_General.Instance))
                        {
                            Ventana = "Pantalla: Detalle No conformidad";

                            Central8.Controls.Add(Consulta_OportunidadesMejora_General.Instance);
                            Consulta_OportunidadesMejora_General.Instance.Dock = DockStyle.Fill;
                            Consulta_OportunidadesMejora_General.Instance.BringToFront();


                        }
                        else
                        {
                            Consulta_OportunidadesMejora_General.Instance.BringToFront();
                            Ventana = "Pantalla: Detalle No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else { MessageBox.Show("No tiene acceso a este modulo"); }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);   
                    }
                    oNetDrive = null;

                    if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "AUXILIAR DE  PRODUCCION")
                    {

                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                        {

                            if (!Central8.Controls.Contains(Consulta_OportunidadesMejora_General.Instance))
                            {
                                Ventana = "Pantalla: Detalle No conformidad";

                                Central8.Controls.Add(Consulta_OportunidadesMejora_General.Instance);
                                Consulta_OportunidadesMejora_General.Instance.Dock = DockStyle.Fill;
                                Consulta_OportunidadesMejora_General.Instance.BringToFront();


                            }
                            else
                            {
                                Consulta_OportunidadesMejora_General.Instance.BringToFront();
                                Ventana = "Pantalla: Detalle No conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                    }
                    else { MessageBox.Show("No tiene acceso a este modulo"); }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void toolStripMenuItem32_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Departamento == "Administracion" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                {

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central.Controls.Contains(Consulta_Oportunidades_Mejora.Instance))
                        {
                            Ventana = "Pantalla: Detalle No conformidad";


                            Consulta_Oportunidades_Mejora.Instance.Nombre = Nombre;
                            Consulta_Oportunidades_Mejora.Instance.Departamento = Departamento;
                            Consulta_Oportunidades_Mejora.Instance.Ventana1 = Ventana;

                            Central.Controls.Add(Consulta_Oportunidades_Mejora.Instance);
                            Consulta_Oportunidades_Mejora.Instance.Dock = DockStyle.Fill;
                            Consulta_Oportunidades_Mejora.Instance.BringToFront();



                        }
                        else
                        {
                            Consulta_OportunidadesMejora_General.Instance.BringToFront();
                            Ventana = "Pantalla: Detalle No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {


                    NetworkDrive oNetDrive1 = new NetworkDrive();
                    try
                    {
                        //set propertys
                        oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                        oNetDrive1.ShareName = url.Text;
                        //match call to options provided
                        oNetDrive1.MapDrive();
                        conexionvpn = 1;
                        //update status
                    }
                    catch (Exception err)
                    {
                        //report error

                        //MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        conexionvpn = 0;
                    }
                    NetworkDrive oNetDrive = new NetworkDrive();
                    try
                    {

                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);    
                    }
                    oNetDrive = null;
                    if (Departamento == "Administracion" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                    {

                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                        {

                            if (!Central.Controls.Contains(Consulta_Oportunidades_Mejora.Instance))
                            {
                                Ventana = "Pantalla: Detalle No conformidad";


                                Consulta_Oportunidades_Mejora.Instance.Nombre = Nombre;
                                Consulta_Oportunidades_Mejora.Instance.Departamento = Departamento;
                                Consulta_Oportunidades_Mejora.Instance.Ventana1 = Ventana;

                                Central.Controls.Add(Consulta_Oportunidades_Mejora.Instance);
                                Consulta_Oportunidades_Mejora.Instance.Dock = DockStyle.Fill;
                                Consulta_Oportunidades_Mejora.Instance.BringToFront();



                            }
                            else
                            {
                                Consulta_OportunidadesMejora_General.Instance.BringToFront();
                                Ventana = "Pantalla: Detalle No conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }



        }

        private void toolStripMenuItem36_Click(object sender, EventArgs e)
        {

            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;


                if (Departamento == "Compras" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                {

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central.Controls.Contains(Consulta_Oportunidades_Mejora.Instance))
                        {
                            Ventana = "Pantalla: Consulta de No Conformidad";
                            Consulta_Oportunidades_Mejora.Instance.Departamento = Departamento;
                            Central.Controls.Add(Consulta_Oportunidades_Mejora.Instance);

                            Consulta_Oportunidades_Mejora.Instance.Dock = DockStyle.Fill;
                            Consulta_Oportunidades_Mejora.Instance.BringToFront();


                        }
                        else
                        {
                            Consulta_Oportunidades_Mejora.Instance.BringToFront();
                            Ventana = "Pantalla: Consulta de No Conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //    MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);   
                    }
                    oNetDrive = null;


                    if (Departamento == "Compras" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                    {

                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                        {

                            if (!Central.Controls.Contains(Consulta_Oportunidades_Mejora.Instance))
                            {
                                Ventana = "Pantalla: Consulta de No Conformidad";
                                Consulta_Oportunidades_Mejora.Instance.Departamento = Departamento;
                                Central.Controls.Add(Consulta_Oportunidades_Mejora.Instance);

                                Consulta_Oportunidades_Mejora.Instance.Dock = DockStyle.Fill;
                                Consulta_Oportunidades_Mejora.Instance.BringToFront();


                            }
                            else
                            {
                                Consulta_Oportunidades_Mejora.Instance.BringToFront();
                                Ventana = "Pantalla: Consulta de No Conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }

                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }








        }

        private void toolStripMenuItem41_Click(object sender, EventArgs e)
        {

            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central.Controls.Contains(Consulta_Oportunidades_Mejora.Instance))
                    {
                        Ventana = "Pantalla: Detalle No conformidad";


                        Consulta_Oportunidades_Mejora.Instance.Nombre = Nombre;
                        Consulta_Oportunidades_Mejora.Instance.Departamento = Departamento;
                        Consulta_Oportunidades_Mejora.Instance.Ventana1 = Ventana;

                        Central.Controls.Add(Consulta_Oportunidades_Mejora.Instance);
                        Consulta_Oportunidades_Mejora.Instance.Dock = DockStyle.Fill;
                        Consulta_Oportunidades_Mejora.Instance.BringToFront();



                    }
                    else
                    {
                        Consulta_OportunidadesMejora_General.Instance.BringToFront();
                        Ventana = "Pantalla: Detalle No conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //    MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);   
                    }
                    oNetDrive = null;
                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central.Controls.Contains(Consulta_Oportunidades_Mejora.Instance))
                        {
                            Ventana = "Pantalla: Detalle No conformidad";


                            Consulta_Oportunidades_Mejora.Instance.Nombre = Nombre;
                            Consulta_Oportunidades_Mejora.Instance.Departamento = Departamento;
                            Consulta_Oportunidades_Mejora.Instance.Ventana1 = Ventana;

                            Central.Controls.Add(Consulta_Oportunidades_Mejora.Instance);
                            Consulta_Oportunidades_Mejora.Instance.Dock = DockStyle.Fill;
                            Consulta_Oportunidades_Mejora.Instance.BringToFront();



                        }
                        else
                        {
                            Consulta_OportunidadesMejora_General.Instance.BringToFront();
                            Ventana = "Pantalla: Detalle No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void toolStripMenuItem46_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central2.Controls.Contains(Consulta_Oportunidades_Mejora.Instance))
                    {
                        Ventana = "Pantalla: Detalle No conformidad";


                        Consulta_Oportunidades_Mejora.Instance.Nombre = Nombre;
                        Consulta_Oportunidades_Mejora.Instance.Departamento = Departamento;
                        Consulta_Oportunidades_Mejora.Instance.Ventana1 = Ventana;

                        Central2.Controls.Add(Consulta_Oportunidades_Mejora.Instance);
                        Consulta_Oportunidades_Mejora.Instance.Dock = DockStyle.Fill;
                        Consulta_Oportunidades_Mejora.Instance.BringToFront();



                    }
                    else
                    {
                        Consulta_OportunidadesMejora_General.Instance.BringToFront();
                        Ventana = "Pantalla: Detalle No conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;
                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central2.Controls.Contains(Consulta_Oportunidades_Mejora.Instance))
                        {
                            Ventana = "Pantalla: Detalle No conformidad";


                            Consulta_Oportunidades_Mejora.Instance.Nombre = Nombre;
                            Consulta_Oportunidades_Mejora.Instance.Departamento = Departamento;
                            Consulta_Oportunidades_Mejora.Instance.Ventana1 = Ventana;

                            Central2.Controls.Add(Consulta_Oportunidades_Mejora.Instance);
                            Consulta_Oportunidades_Mejora.Instance.Dock = DockStyle.Fill;
                            Consulta_Oportunidades_Mejora.Instance.BringToFront();



                        }
                        else
                        {
                            Consulta_OportunidadesMejora_General.Instance.BringToFront();
                            Ventana = "Pantalla: Detalle No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }










        }

        private void toolStripMenuItem10_Click(object sender, EventArgs e)
        {

        }

        private void toolStripMenuItem51_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;

                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central3.Controls.Contains(Consulta_Oportunidades_Mejora.Instance))
                    {
                        Ventana = "Pantalla: Detalle No conformidad";


                        Consulta_Oportunidades_Mejora.Instance.Nombre = Nombre;
                        Consulta_Oportunidades_Mejora.Instance.Departamento = Departamento;
                        Consulta_Oportunidades_Mejora.Instance.Ventana1 = Ventana;

                        Central3.Controls.Add(Consulta_Oportunidades_Mejora.Instance);
                        Consulta_Oportunidades_Mejora.Instance.Dock = DockStyle.Fill;
                        Consulta_Oportunidades_Mejora.Instance.BringToFront();



                    }
                    else
                    {
                        Consulta_OportunidadesMejora_General.Instance.BringToFront();
                        Ventana = "Pantalla: Detalle No conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); 
                    }
                    oNetDrive = null;

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central3.Controls.Contains(Consulta_Oportunidades_Mejora.Instance))
                        {
                            Ventana = "Pantalla: Detalle No conformidad";


                            Consulta_Oportunidades_Mejora.Instance.Nombre = Nombre;
                            Consulta_Oportunidades_Mejora.Instance.Departamento = Departamento;
                            Consulta_Oportunidades_Mejora.Instance.Ventana1 = Ventana;

                            Central3.Controls.Add(Consulta_Oportunidades_Mejora.Instance);
                            Consulta_Oportunidades_Mejora.Instance.Dock = DockStyle.Fill;
                            Consulta_Oportunidades_Mejora.Instance.BringToFront();



                        }
                        else
                        {
                            Consulta_OportunidadesMejora_General.Instance.BringToFront();
                            Ventana = "Pantalla: Detalle No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }

                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }









        }

        private void toolStripMenuItem56_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;

                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central4.Controls.Contains(Consulta_Oportunidades_Mejora.Instance))
                    {
                        Ventana = "Pantalla: Detalle No conformidad";


                        Consulta_Oportunidades_Mejora.Instance.Nombre = Nombre;
                        Consulta_Oportunidades_Mejora.Instance.Departamento = Departamento;
                        Consulta_Oportunidades_Mejora.Instance.Ventana1 = Ventana;

                        Central4.Controls.Add(Consulta_Oportunidades_Mejora.Instance);
                        Consulta_Oportunidades_Mejora.Instance.Dock = DockStyle.Fill;
                        Consulta_Oportunidades_Mejora.Instance.BringToFront();



                    }
                    else
                    {
                        Consulta_OportunidadesMejora_General.Instance.BringToFront();
                        Ventana = "Pantalla: Detalle No conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);    
                    }
                    oNetDrive = null;

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central4.Controls.Contains(Consulta_Oportunidades_Mejora.Instance))
                        {
                            Ventana = "Pantalla: Detalle No conformidad";


                            Consulta_Oportunidades_Mejora.Instance.Nombre = Nombre;
                            Consulta_Oportunidades_Mejora.Instance.Departamento = Departamento;
                            Consulta_Oportunidades_Mejora.Instance.Ventana1 = Ventana;

                            Central4.Controls.Add(Consulta_Oportunidades_Mejora.Instance);
                            Consulta_Oportunidades_Mejora.Instance.Dock = DockStyle.Fill;
                            Consulta_Oportunidades_Mejora.Instance.BringToFront();



                        }
                        else
                        {
                            Consulta_OportunidadesMejora_General.Instance.BringToFront();
                            Ventana = "Pantalla: Detalle No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void toolStripMenuItem61_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;

                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central5.Controls.Contains(Consulta_Oportunidades_Mejora.Instance))
                    {
                        Ventana = "Pantalla: Detalle No conformidad";


                        Consulta_Oportunidades_Mejora.Instance.Nombre = Nombre;
                        Consulta_Oportunidades_Mejora.Instance.Departamento = Departamento;
                        Consulta_Oportunidades_Mejora.Instance.Ventana1 = Ventana;

                        Central5.Controls.Add(Consulta_Oportunidades_Mejora.Instance);
                        Consulta_Oportunidades_Mejora.Instance.Dock = DockStyle.Fill;
                        Consulta_Oportunidades_Mejora.Instance.BringToFront();



                    }
                    else
                    {
                        Consulta_OportunidadesMejora_General.Instance.BringToFront();
                        Ventana = "Pantalla: Detalle No conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //      MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); 
                    }
                    oNetDrive = null;

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central5.Controls.Contains(Consulta_Oportunidades_Mejora.Instance))
                        {
                            Ventana = "Pantalla: Detalle No conformidad";


                            Consulta_Oportunidades_Mejora.Instance.Nombre = Nombre;
                            Consulta_Oportunidades_Mejora.Instance.Departamento = Departamento;
                            Consulta_Oportunidades_Mejora.Instance.Ventana1 = Ventana;

                            Central5.Controls.Add(Consulta_Oportunidades_Mejora.Instance);
                            Consulta_Oportunidades_Mejora.Instance.Dock = DockStyle.Fill;
                            Consulta_Oportunidades_Mejora.Instance.BringToFront();



                        }
                        else
                        {
                            Consulta_OportunidadesMejora_General.Instance.BringToFront();
                            Ventana = "Pantalla: Detalle No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }

                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }









        }

        private void toolStripMenuItem66_Click(object sender, EventArgs e)
        {

            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central6.Controls.Contains(Consulta_Oportunidades_Mejora.Instance))
                    {
                        Ventana = "Pantalla: Detalle No conformidad";


                        Consulta_Oportunidades_Mejora.Instance.Nombre = Nombre;
                        Consulta_Oportunidades_Mejora.Instance.Departamento = Departamento;
                        Consulta_Oportunidades_Mejora.Instance.Ventana1 = Ventana;

                        Central6.Controls.Add(Consulta_Oportunidades_Mejora.Instance);
                        Consulta_Oportunidades_Mejora.Instance.Dock = DockStyle.Fill;
                        Consulta_Oportunidades_Mejora.Instance.BringToFront();



                    }
                    else
                    {
                        Consulta_OportunidadesMejora_General.Instance.BringToFront();
                        Ventana = "Pantalla: Detalle No conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;
                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central6.Controls.Contains(Consulta_Oportunidades_Mejora.Instance))
                        {
                            Ventana = "Pantalla: Detalle No conformidad";


                            Consulta_Oportunidades_Mejora.Instance.Nombre = Nombre;
                            Consulta_Oportunidades_Mejora.Instance.Departamento = Departamento;
                            Consulta_Oportunidades_Mejora.Instance.Ventana1 = Ventana;

                            Central6.Controls.Add(Consulta_Oportunidades_Mejora.Instance);
                            Consulta_Oportunidades_Mejora.Instance.Dock = DockStyle.Fill;
                            Consulta_Oportunidades_Mejora.Instance.BringToFront();



                        }
                        else
                        {
                            Consulta_OportunidadesMejora_General.Instance.BringToFront();
                            Ventana = "Pantalla: Detalle No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }









        }

        private void toolStripMenuItem76_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;

                if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Departamento == "Infraestructura")
                {


                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central8.Controls.Contains(Consulta_Oportunidades_Mejora.Instance))
                        {
                            Ventana = "Pantalla: Detalle No conformidad";


                            Consulta_Oportunidades_Mejora.Instance.Nombre = Nombre;
                            Consulta_Oportunidades_Mejora.Instance.Departamento = Departamento;
                            Consulta_Oportunidades_Mejora.Instance.Ventana1 = Ventana;

                            Central8.Controls.Add(Consulta_Oportunidades_Mejora.Instance);
                            Consulta_Oportunidades_Mejora.Instance.Dock = DockStyle.Fill;
                            Consulta_Oportunidades_Mejora.Instance.BringToFront();



                        }
                        else
                        {
                            Consulta_OportunidadesMejora_General.Instance.BringToFront();
                            Ventana = "Pantalla: Detalle No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else { MessageBox.Show("No tiene acceso a este modulo"); }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //    MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);   
                    }
                    oNetDrive = null;

                    if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "AUXILIAR DE  PRODUCCION")
                    {


                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                        {

                            if (!Central8.Controls.Contains(Consulta_Oportunidades_Mejora.Instance))
                            {
                                Ventana = "Pantalla: Detalle No conformidad";


                                Consulta_Oportunidades_Mejora.Instance.Nombre = Nombre;
                                Consulta_Oportunidades_Mejora.Instance.Departamento = Departamento;
                                Consulta_Oportunidades_Mejora.Instance.Ventana1 = Ventana;

                                Central8.Controls.Add(Consulta_Oportunidades_Mejora.Instance);
                                Consulta_Oportunidades_Mejora.Instance.Dock = DockStyle.Fill;
                                Consulta_Oportunidades_Mejora.Instance.BringToFront();



                            }
                            else
                            {
                                Consulta_OportunidadesMejora_General.Instance.BringToFront();
                                Ventana = "Pantalla: Detalle No conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                    }
                    else { MessageBox.Show("No tiene acceso a este modulo"); }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void toolStripMenuItem16_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "AUXILIAR DE  PRODUCCION")
                {

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central8.Controls.Contains(Genera_No_Confor.Instance))
                        {
                            Ventana = "Pantalla: Levantamiento de No conformidad";
                            Genera_No_Confor.Instance.Nombre = Nombre;
                            Genera_No_Confor.Instance.Ventana1 = Ventana;

                            Central8.Controls.Add(Genera_No_Confor.Instance);
                            Genera_No_Confor.Instance.Dock = DockStyle.Fill;
                            Genera_No_Confor.Instance.BringToFront();


                        }
                        else
                        {
                            Genera_No_Confor.Instance.BringToFront();
                            Ventana = "Pantalla: Levantamiento de No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else { MessageBox.Show("No tiene acceso a este modulo"); }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {

                            try
                            {
                                //set propertys
                                oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                                oNetDrive1.ShareName = url.Text;
                                //match call to options provided
                                oNetDrive1.MapDrive();
                                conexionvpn = 1;
                                //update status
                            }
                            catch (Exception err)
                            {
                                //report error

                                //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                            }
                            oNetDrive1 = null;

                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {

                        //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);    
                    }
                    oNetDrive = null;
                    if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "AUXILIAR DE  PRODUCCION")
                    {

                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                        {

                            if (!Central8.Controls.Contains(Genera_No_Confor.Instance))
                            {
                                Ventana = "Pantalla: Levantamiento de No conformidad";
                                Genera_No_Confor.Instance.Nombre = Nombre;
                                Genera_No_Confor.Instance.Ventana1 = Ventana;

                                Central8.Controls.Add(Genera_No_Confor.Instance);
                                Genera_No_Confor.Instance.Dock = DockStyle.Fill;
                                Genera_No_Confor.Instance.BringToFront();


                            }
                            else
                            {
                                Genera_No_Confor.Instance.BringToFront();
                                Ventana = "Pantalla: Levantamiento de No conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                    }
                    else { MessageBox.Show("No tiene acceso a este modulo"); }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }



        }




        private void toolStripMenuItem17_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "AUXILIAR DE  PRODUCCION")
                {


                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central8.Controls.Contains(Consulta_No_Conf.Instance))
                        {
                            Ventana = "Pantalla: Consulta de No Conformidad";
                            Consulta_No_Conf.Instance.Departamento = Departamento;
                            Central8.Controls.Add(Consulta_No_Conf.Instance);
                            Consulta_No_Conf.Instance.Dock = DockStyle.Fill;
                            Consulta_No_Conf.Instance.BringToFront();


                        }
                        else
                        {
                            Consulta_No_Conf.Instance.BringToFront();
                            Ventana = "Pantalla: Consulta de No Conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else { MessageBox.Show("No tiene acceso a este modulo"); }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); 
                    }
                    oNetDrive = null;
                    if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "AUXILIAR DE  PRODUCCION")
                    {


                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                        {

                            if (!Central8.Controls.Contains(Consulta_No_Conf.Instance))
                            {
                                Ventana = "Pantalla: Consulta de No Conformidad";
                                Consulta_No_Conf.Instance.Departamento = Departamento;
                                Central8.Controls.Add(Consulta_No_Conf.Instance);
                                Consulta_No_Conf.Instance.Dock = DockStyle.Fill;
                                Consulta_No_Conf.Instance.BringToFront();


                            }
                            else
                            {
                                Consulta_No_Conf.Instance.BringToFront();
                                Ventana = "Pantalla: Consulta de No Conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                    }
                    else { MessageBox.Show("No tiene acceso a este modulo"); }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }



        }

        private void toolStripMenuItem18_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "AUXILIAR DE  PRODUCCION")
                {

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central8.Controls.Contains(Consulta_No_Conf_general.Instance))
                        {
                            Ventana = "Pantalla: Consulta de No Conformidad";

                            Central8.Controls.Add(Consulta_No_Conf_general.Instance);
                            Consulta_No_Conf_general.Instance.Dock = DockStyle.Fill;
                            Consulta_No_Conf_general.Instance.BringToFront();


                        }
                        else
                        {
                            Consulta_No_Conf_general.Instance.BringToFront();
                            Ventana = "Pantalla: Consulta de No Conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else { MessageBox.Show("No tiene acceso a este modulo"); }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);
                    }
                    oNetDrive = null;

                    if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "AUXILIAR DE  PRODUCCION")
                    {

                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                        {

                            if (!Central8.Controls.Contains(Consulta_No_Conf_general.Instance))
                            {
                                Ventana = "Pantalla: Consulta de No Conformidad";

                                Central8.Controls.Add(Consulta_No_Conf_general.Instance);
                                Consulta_No_Conf_general.Instance.Dock = DockStyle.Fill;
                                Consulta_No_Conf_general.Instance.BringToFront();


                            }
                            else
                            {
                                Consulta_No_Conf_general.Instance.BringToFront();
                                Ventana = "Pantalla: Consulta de No Conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                    }
                    else { MessageBox.Show("No tiene acceso a este modulo"); }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void seguimientoToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Departamento == "Compras" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                {

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central.Controls.Contains(Seguimiento_No_confor.Instance))
                        {
                            Ventana = "Pantalla: Seguimiento de No conformidad";
                            Seguimiento_No_confor.Instance.Departamento = Departamento;
                            Seguimiento_No_confor.Instance.Nombre = Nombre;
                            Seguimiento_No_confor.Instance.Apellido = Apellido;
                            Seguimiento_No_confor.Instance.NoConfor = NumeroConformidad;
                            Seguimiento_No_confor.Instance.Ventana1 = Ventana;
                            Central.Controls.Add(Seguimiento_No_confor.Instance);
                            Seguimiento_No_confor.Instance.Dock = DockStyle.Fill;
                            Seguimiento_No_confor.Instance.BringToFront();


                        }
                        else
                        {
                            Seguimiento_No_confor.Instance.BringToFront();
                            Ventana = "Pantalla: Seguimiento de No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;
                    if (Departamento == "Compras" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                    {

                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                        {

                            if (!Central.Controls.Contains(Seguimiento_No_confor.Instance))
                            {
                                Ventana = "Pantalla: Seguimiento de No conformidad";
                                Seguimiento_No_confor.Instance.Departamento = Departamento;
                                Seguimiento_No_confor.Instance.Nombre = Nombre;
                                Seguimiento_No_confor.Instance.Apellido = Apellido;
                                Seguimiento_No_confor.Instance.NoConfor = NumeroConformidad;
                                Seguimiento_No_confor.Instance.Ventana1 = Ventana;
                                Central.Controls.Add(Seguimiento_No_confor.Instance);
                                Seguimiento_No_confor.Instance.Dock = DockStyle.Fill;
                                Seguimiento_No_confor.Instance.BringToFront();


                            }
                            else
                            {
                                Seguimiento_No_confor.Instance.BringToFront();
                                Ventana = "Pantalla: Seguimiento de No conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void consultaToolStripMenuItem6_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Departamento == "Cobranza" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                {

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central.Controls.Contains(Consulta_No_Conf.Instance))
                        {
                            Ventana = "Pantalla: Consulta de No Conformidad";
                            Consulta_No_Conf.Instance.Departamento = Departamento;
                            Central.Controls.Add(Consulta_No_Conf.Instance);
                            Consulta_No_Conf.Instance.Dock = DockStyle.Fill;
                            Consulta_No_Conf.Instance.BringToFront();


                        }
                        else
                        {
                            Consulta_No_Conf.Instance.BringToFront();
                            Ventana = "Pantalla: Consulta de No Conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }

                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;
                    if (Departamento == "Cobranza" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                    {

                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                        {

                            if (!Central.Controls.Contains(Consulta_No_Conf.Instance))
                            {
                                Ventana = "Pantalla: Consulta de No Conformidad";
                                Consulta_No_Conf.Instance.Departamento = Departamento;
                                Central.Controls.Add(Consulta_No_Conf.Instance);
                                Consulta_No_Conf.Instance.Dock = DockStyle.Fill;
                                Consulta_No_Conf.Instance.BringToFront();


                            }
                            else
                            {
                                Consulta_No_Conf.Instance.BringToFront();
                                Ventana = "Pantalla: Consulta de No Conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                    }

                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }

        }

        private void cobranzaToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void seguimientoToolStripMenuItem6_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Departamento == "Administracion" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                {
                    ThreadStart proceso = new ThreadStart(MostrarCuadroInicio2); Thread hilo = new Thread(proceso);

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central.Controls.Contains(Seguimiento_No_confor.Instance))
                        {
                            Ventana = "Pantalla: Seguimiento de No conformidad";
                            Seguimiento_No_confor.Instance.Departamento = Departamento;
                            Seguimiento_No_confor.Instance.Nombre = Nombre;
                            Seguimiento_No_confor.Instance.Apellido = Apellido;
                            Seguimiento_No_confor.Instance.NoConfor = NumeroConformidad;
                            Seguimiento_No_confor.Instance.Ventana1 = Ventana;
                            Central.Controls.Add(Seguimiento_No_confor.Instance);
                            Seguimiento_No_confor.Instance.Dock = DockStyle.Fill;
                            Seguimiento_No_confor.Instance.BringToFront();


                        }
                        else
                        {
                            Seguimiento_No_confor.Instance.BringToFront();
                            Ventana = "Pantalla: Seguimiento de No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //      MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                            conexionvpn = 0;
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);   
                    }
                    oNetDrive = null;
                    if (Departamento == "Administracion" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                    {
                        ThreadStart proceso = new ThreadStart(MostrarCuadroInicio2); Thread hilo = new Thread(proceso);

                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                        {

                            if (!Central.Controls.Contains(Seguimiento_No_confor.Instance))
                            {
                                Ventana = "Pantalla: Seguimiento de No conformidad";
                                Seguimiento_No_confor.Instance.Departamento = Departamento;
                                Seguimiento_No_confor.Instance.Nombre = Nombre;
                                Seguimiento_No_confor.Instance.Apellido = Apellido;
                                Seguimiento_No_confor.Instance.NoConfor = NumeroConformidad;
                                Seguimiento_No_confor.Instance.Ventana1 = Ventana;
                                Central.Controls.Add(Seguimiento_No_confor.Instance);
                                Seguimiento_No_confor.Instance.Dock = DockStyle.Fill;
                                Seguimiento_No_confor.Instance.BringToFront();


                            }
                            else
                            {
                                Seguimiento_No_confor.Instance.BringToFront();
                                Ventana = "Pantalla: Seguimiento de No conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }

        }

        private void seguimientoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;

                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central.Controls.Contains(Seguimiento_No_confor.Instance))
                    {
                        Ventana = "Pantalla: Seguimiento de No conformidad";
                        Seguimiento_No_confor.Instance.Departamento = Departamento;
                        Seguimiento_No_confor.Instance.Nombre = Nombre;
                        Seguimiento_No_confor.Instance.Apellido = Apellido;
                        Seguimiento_No_confor.Instance.NoConfor = NumeroConformidad;
                        Seguimiento_No_confor.Instance.Ventana1 = Ventana;
                        Central2.Controls.Add(Seguimiento_No_confor.Instance);
                        Seguimiento_No_confor.Instance.Dock = DockStyle.Fill;
                        Seguimiento_No_confor.Instance.BringToFront();


                    }
                    else
                    {
                        Seguimiento_No_confor.Instance.BringToFront();
                        Ventana = "Pantalla: Seguimiento de No conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //     MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);    
                    }
                    oNetDrive = null;
                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central.Controls.Contains(Seguimiento_No_confor.Instance))
                        {
                            Ventana = "Pantalla: Seguimiento de No conformidad";
                            Seguimiento_No_confor.Instance.Departamento = Departamento;
                            Seguimiento_No_confor.Instance.Nombre = Nombre;
                            Seguimiento_No_confor.Instance.Apellido = Apellido;
                            Seguimiento_No_confor.Instance.NoConfor = NumeroConformidad;
                            Seguimiento_No_confor.Instance.Ventana1 = Ventana;
                            Central2.Controls.Add(Seguimiento_No_confor.Instance);
                            Seguimiento_No_confor.Instance.Dock = DockStyle.Fill;
                            Seguimiento_No_confor.Instance.BringToFront();


                        }
                        else
                        {
                            Seguimiento_No_confor.Instance.BringToFront();
                            Ventana = "Pantalla: Seguimiento de No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }
        }

        private void seguimientoToolStripMenuItem1_Click(object sender, EventArgs e)
        {

            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;

                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central3.Controls.Contains(Seguimiento_No_confor.Instance))
                    {
                        Ventana = "Pantalla: Seguimiento de No conformidad";
                        Seguimiento_No_confor.Instance.Departamento = Departamento;
                        Seguimiento_No_confor.Instance.Nombre = Nombre;
                        Seguimiento_No_confor.Instance.Apellido = Apellido;
                        Seguimiento_No_confor.Instance.NoConfor = NumeroConformidad;
                        Seguimiento_No_confor.Instance.Ventana1 = Ventana;
                        Central3.Controls.Add(Seguimiento_No_confor.Instance);
                        Seguimiento_No_confor.Instance.Dock = DockStyle.Fill;
                        Seguimiento_No_confor.Instance.BringToFront();


                    }
                    else
                    {
                        Seguimiento_No_confor.Instance.BringToFront();
                        Ventana = "Pantalla: Seguimiento de No conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //    MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);   
                    }
                    oNetDrive = null;

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central3.Controls.Contains(Seguimiento_No_confor.Instance))
                        {
                            Ventana = "Pantalla: Seguimiento de No conformidad";
                            Seguimiento_No_confor.Instance.Departamento = Departamento;
                            Seguimiento_No_confor.Instance.Nombre = Nombre;
                            Seguimiento_No_confor.Instance.Apellido = Apellido;
                            Seguimiento_No_confor.Instance.NoConfor = NumeroConformidad;
                            Seguimiento_No_confor.Instance.Ventana1 = Ventana;
                            Central3.Controls.Add(Seguimiento_No_confor.Instance);
                            Seguimiento_No_confor.Instance.Dock = DockStyle.Fill;
                            Seguimiento_No_confor.Instance.BringToFront();


                        }
                        else
                        {
                            Seguimiento_No_confor.Instance.BringToFront();
                            Ventana = "Pantalla: Seguimiento de No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void seguimientoToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central4.Controls.Contains(Seguimiento_No_confor.Instance))
                    {
                        Ventana = "Pantalla: Seguimiento de No conformidad";
                        Seguimiento_No_confor.Instance.Departamento = Departamento;
                        Seguimiento_No_confor.Instance.Nombre = Nombre;
                        Seguimiento_No_confor.Instance.Apellido = Apellido;
                        Seguimiento_No_confor.Instance.NoConfor = NumeroConformidad;
                        Seguimiento_No_confor.Instance.Ventana1 = Ventana;
                        Central4.Controls.Add(Seguimiento_No_confor.Instance);
                        Seguimiento_No_confor.Instance.Dock = DockStyle.Fill;
                        Seguimiento_No_confor.Instance.BringToFront();


                    }
                    else
                    {
                        Seguimiento_No_confor.Instance.BringToFront();
                        Ventana = "Pantalla: Seguimiento de No conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //    MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;
                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central4.Controls.Contains(Seguimiento_No_confor.Instance))
                        {
                            Ventana = "Pantalla: Seguimiento de No conformidad";
                            Seguimiento_No_confor.Instance.Departamento = Departamento;
                            Seguimiento_No_confor.Instance.Nombre = Nombre;
                            Seguimiento_No_confor.Instance.Apellido = Apellido;
                            Seguimiento_No_confor.Instance.NoConfor = NumeroConformidad;
                            Seguimiento_No_confor.Instance.Ventana1 = Ventana;
                            Central4.Controls.Add(Seguimiento_No_confor.Instance);
                            Seguimiento_No_confor.Instance.Dock = DockStyle.Fill;
                            Seguimiento_No_confor.Instance.BringToFront();


                        }
                        else
                        {
                            Seguimiento_No_confor.Instance.BringToFront();
                            Ventana = "Pantalla: Seguimiento de No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void toolStripMenuItem6_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;

                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central5.Controls.Contains(Seguimiento_No_confor.Instance))
                    {
                        Ventana = "Pantalla: Seguimiento de No conformidad";
                        Seguimiento_No_confor.Instance.Departamento = Departamento;
                        Seguimiento_No_confor.Instance.Nombre = Nombre;
                        Seguimiento_No_confor.Instance.Apellido = Apellido;
                        Seguimiento_No_confor.Instance.NoConfor = NumeroConformidad;
                        Seguimiento_No_confor.Instance.Ventana1 = Ventana;
                        Central5.Controls.Add(Seguimiento_No_confor.Instance);
                        Seguimiento_No_confor.Instance.Dock = DockStyle.Fill;
                        Seguimiento_No_confor.Instance.BringToFront();


                    }
                    else
                    {
                        Seguimiento_No_confor.Instance.BringToFront();
                        Ventana = "Pantalla: Seguimiento de No conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //    MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);   
                    }
                    oNetDrive = null;

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central5.Controls.Contains(Seguimiento_No_confor.Instance))
                        {
                            Ventana = "Pantalla: Seguimiento de No conformidad";
                            Seguimiento_No_confor.Instance.Departamento = Departamento;
                            Seguimiento_No_confor.Instance.Nombre = Nombre;
                            Seguimiento_No_confor.Instance.Apellido = Apellido;
                            Seguimiento_No_confor.Instance.NoConfor = NumeroConformidad;
                            Seguimiento_No_confor.Instance.Ventana1 = Ventana;
                            Central5.Controls.Add(Seguimiento_No_confor.Instance);
                            Seguimiento_No_confor.Instance.Dock = DockStyle.Fill;
                            Seguimiento_No_confor.Instance.BringToFront();


                        }
                        else
                        {
                            Seguimiento_No_confor.Instance.BringToFront();
                            Ventana = "Pantalla: Seguimiento de No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }



        }

        private void toolStripMenuItem13_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;

                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central6.Controls.Contains(Seguimiento_No_confor.Instance))
                    {
                        Ventana = "Pantalla: Seguimiento de No conformidad";
                        Seguimiento_No_confor.Instance.Departamento = Departamento;
                        Seguimiento_No_confor.Instance.Nombre = Nombre;
                        Seguimiento_No_confor.Instance.Apellido = Apellido;
                        Seguimiento_No_confor.Instance.NoConfor = NumeroConformidad;
                        Seguimiento_No_confor.Instance.Ventana1 = Ventana;
                        Central6.Controls.Add(Seguimiento_No_confor.Instance);
                        Seguimiento_No_confor.Instance.Dock = DockStyle.Fill;
                        Seguimiento_No_confor.Instance.BringToFront();


                    }
                    else
                    {
                        Seguimiento_No_confor.Instance.BringToFront();
                        Ventana = "Pantalla: Seguimiento de No conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //    MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);   
                    }
                    oNetDrive = null;

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central6.Controls.Contains(Seguimiento_No_confor.Instance))
                        {
                            Ventana = "Pantalla: Seguimiento de No conformidad";
                            Seguimiento_No_confor.Instance.Departamento = Departamento;
                            Seguimiento_No_confor.Instance.Nombre = Nombre;
                            Seguimiento_No_confor.Instance.Apellido = Apellido;
                            Seguimiento_No_confor.Instance.NoConfor = NumeroConformidad;
                            Seguimiento_No_confor.Instance.Ventana1 = Ventana;
                            Central6.Controls.Add(Seguimiento_No_confor.Instance);
                            Seguimiento_No_confor.Instance.Dock = DockStyle.Fill;
                            Seguimiento_No_confor.Instance.BringToFront();


                        }
                        else
                        {
                            Seguimiento_No_confor.Instance.BringToFront();
                            Ventana = "Pantalla: Seguimiento de No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void toolStripMenuItem19_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "AUXILIAR DE  PRODUCCION")
                {

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central8.Controls.Contains(Seguimiento_No_confor.Instance))
                        {
                            Ventana = "Pantalla: Seguimiento de No conformidad";
                            Seguimiento_No_confor.Instance.Departamento = Departamento;
                            Seguimiento_No_confor.Instance.Nombre = Nombre;
                            Seguimiento_No_confor.Instance.Apellido = Apellido;
                            Seguimiento_No_confor.Instance.NoConfor = NumeroConformidad;
                            Seguimiento_No_confor.Instance.Ventana1 = Ventana;
                            Central8.Controls.Add(Seguimiento_No_confor.Instance);
                            Seguimiento_No_confor.Instance.Dock = DockStyle.Fill;
                            Seguimiento_No_confor.Instance.BringToFront();


                        }
                        else
                        {
                            Seguimiento_No_confor.Instance.BringToFront();
                            Ventana = "Pantalla: Seguimiento de No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else { MessageBox.Show("No tiene acceso a este modulo"); }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); 
                    }
                    oNetDrive = null;
                    if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "AUXILIAR DE  PRODUCCION")
                    {

                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                        {

                            if (!Central8.Controls.Contains(Seguimiento_No_confor.Instance))
                            {
                                Ventana = "Pantalla: Seguimiento de No conformidad";
                                Seguimiento_No_confor.Instance.Departamento = Departamento;
                                Seguimiento_No_confor.Instance.Nombre = Nombre;
                                Seguimiento_No_confor.Instance.Apellido = Apellido;
                                Seguimiento_No_confor.Instance.NoConfor = NumeroConformidad;
                                Seguimiento_No_confor.Instance.Ventana1 = Ventana;
                                Central8.Controls.Add(Seguimiento_No_confor.Instance);
                                Seguimiento_No_confor.Instance.Dock = DockStyle.Fill;
                                Seguimiento_No_confor.Instance.BringToFront();


                            }
                            else
                            {
                                Seguimiento_No_confor.Instance.BringToFront();
                                Ventana = "Pantalla: Seguimiento de No conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                    }
                    else { MessageBox.Show("No tiene acceso a este modulo"); }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }



        }

        private void altaDeUsuariosToolStripMenuItem_Click(object sender, EventArgs e)
        {


            if (Permiso == "1" || Permiso == "2")
            {

                if (!Central9.Controls.Contains(Alta_usuarios.Instance))
                {
                    Ventana = "Pantalla: Seguimiento de No conformidad";

                    Central9.Controls.Add(Alta_usuarios.Instance);
                    Alta_usuarios.Instance.Dock = DockStyle.Fill;
                    Alta_usuarios.Instance.BringToFront();


                }
                else
                {
                    Alta_usuarios.Instance.BringToFront();
                    Ventana = "Pantalla: Seguimiento de No conformidad";
                }
            }
            else
            {
                MessageBox.Show("No tiene acceso a este modulo");
            }
        }

        private void bajaDeUsuariosToolStripMenuItem_Click(object sender, EventArgs e)
        {



            if (Permiso == "1")
            {

                if (!Central9.Controls.Contains(Baja_usuarios.Instance))
                {
                    Ventana = "Pantalla: Seguimiento de No conformidad";

                    Central9.Controls.Add(Baja_usuarios.Instance);
                    Baja_usuarios.Instance.Dock = DockStyle.Fill;
                    Baja_usuarios.Instance.BringToFront();


                }
                else
                {
                    Baja_usuarios.Instance.BringToFront();
                    Ventana = "Pantalla: Seguimiento de No conformidad";
                }
            }
            else
            {
                MessageBox.Show("No tiene acceso a este modulo");
            }
        }

        private void cambioDeContraseñaToolStripMenuItem_Click(object sender, EventArgs e)
        {



            if (Permiso == "1")
            {

                if (!Central9.Controls.Contains(Config_usuarios.Instance))
                {


                    Central9.Controls.Add(Config_usuarios.Instance);
                    Config_usuarios.Instance.Dock = DockStyle.Fill;
                    Config_usuarios.Instance.BringToFront();


                }
                else
                {
                    Config_usuarios.Instance.BringToFront();

                }
            }
            else
            {
                MessageBox.Show("No tiene acceso a este modulo");
            }
        }

        private void administracionToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (!Central.Controls.Contains(IT.Administracion.Instance))
                {
                    CBR_ADMIN.IT.Administracion.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Administracion.Instance.Apellido = Apellido;
                    Central.Controls.Add(CBR_ADMIN.IT.Administracion.Instance);
                    CBR_ADMIN.IT.Administracion.Instance.Dock = DockStyle.Fill;
                    CBR_ADMIN.IT.Administracion.Instance.BringToFront();


                }
                else
                {
                    CBR_ADMIN.IT.Administracion.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //    MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); 
                    }
                    oNetDrive = null;

                    if (!Central.Controls.Contains(IT.Administracion.Instance))
                    {
                        CBR_ADMIN.IT.Administracion.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Administracion.Instance.Apellido = Apellido;
                        Central.Controls.Add(CBR_ADMIN.IT.Administracion.Instance);
                        CBR_ADMIN.IT.Administracion.Instance.Dock = DockStyle.Fill;
                        CBR_ADMIN.IT.Administracion.Instance.BringToFront();


                    }
                    else
                    {
                        CBR_ADMIN.IT.Administracion.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }




        }

        private void aplicativoToolStripMenuItem_Click(object sender, EventArgs e)
        {


            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (!Central.Controls.Contains(IT.Aplicativo.Instance))
                {
                    CBR_ADMIN.IT.Aplicativo.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Aplicativo.Instance.Apellido = Apellido;

                    Central.Controls.Add(IT.Aplicativo.Instance);
                    IT.Aplicativo.Instance.Dock = DockStyle.Fill;
                    IT.Aplicativo.Instance.BringToFront();


                }
                else
                {
                    IT.Administracion.Instance.BringToFront();

                }

            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);   
                    }
                    oNetDrive = null;
                    if (!Central.Controls.Contains(IT.Aplicativo.Instance))
                    {
                        CBR_ADMIN.IT.Aplicativo.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Aplicativo.Instance.Apellido = Apellido;

                        Central.Controls.Add(IT.Aplicativo.Instance);
                        IT.Aplicativo.Instance.Dock = DockStyle.Fill;
                        IT.Aplicativo.Instance.BringToFront();


                    }
                    else
                    {
                        IT.Administracion.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }



        }

        private void optimizacionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (!Central.Controls.Contains(IT.Optimizacion.Instance))
                {
                    CBR_ADMIN.IT.Optimizacion.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Optimizacion.Instance.Apellido = Apellido;

                    Central.Controls.Add(CBR_ADMIN.IT.Optimizacion.Instance);
                    CBR_ADMIN.IT.Optimizacion.Instance.Dock = DockStyle.Fill;
                    CBR_ADMIN.IT.Optimizacion.Instance.BringToFront();


                }
                else
                {
                    CBR_ADMIN.IT.Optimizacion.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);    
                    }
                    oNetDrive = null;
                    if (!Central.Controls.Contains(IT.Optimizacion.Instance))
                    {
                        CBR_ADMIN.IT.Optimizacion.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Optimizacion.Instance.Apellido = Apellido;

                        Central.Controls.Add(CBR_ADMIN.IT.Optimizacion.Instance);
                        CBR_ADMIN.IT.Optimizacion.Instance.Dock = DockStyle.Fill;
                        CBR_ADMIN.IT.Optimizacion.Instance.BringToFront();


                    }
                    else
                    {
                        CBR_ADMIN.IT.Optimizacion.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }
        public void validacion()
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else { conexionvpn = 2; }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); 
                    }
                    oNetDrive = null;
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }
        }
        private void otrosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (!Central.Controls.Contains(IT.Otros.Instance))
                {
                    CBR_ADMIN.IT.Otros.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Otros.Instance.Apellido = Apellido;
                    Central.Controls.Add(CBR_ADMIN.IT.Otros.Instance);
                    CBR_ADMIN.IT.Otros.Instance.Dock = DockStyle.Fill;
                    CBR_ADMIN.IT.Otros.Instance.BringToFront();


                }
                else
                {
                    CBR_ADMIN.IT.Otros.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            // MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); 
                    }
                    oNetDrive = null;
                    if (!Central.Controls.Contains(IT.Otros.Instance))
                    {
                        CBR_ADMIN.IT.Otros.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Otros.Instance.Apellido = Apellido;
                        Central.Controls.Add(CBR_ADMIN.IT.Otros.Instance);
                        CBR_ADMIN.IT.Otros.Instance.Dock = DockStyle.Fill;
                        CBR_ADMIN.IT.Otros.Instance.BringToFront();


                    }
                    else
                    {
                        CBR_ADMIN.IT.Otros.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void serviciosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (!Central.Controls.Contains(IT.Servicio.Instance))
                {
                    CBR_ADMIN.IT.Servicio.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Servicio.Instance.Apellido = Apellido;
                    Central.Controls.Add(CBR_ADMIN.IT.Servicio.Instance);
                    CBR_ADMIN.IT.Servicio.Instance.Dock = DockStyle.Fill;
                    CBR_ADMIN.IT.Servicio.Instance.BringToFront();


                }
                else
                {
                    CBR_ADMIN.IT.Servicio.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //    MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {//  MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);   
                    }
                    oNetDrive = null;
                    if (!Central.Controls.Contains(IT.Servicio.Instance))
                    {
                        CBR_ADMIN.IT.Servicio.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Servicio.Instance.Apellido = Apellido;
                        Central.Controls.Add(CBR_ADMIN.IT.Servicio.Instance);
                        CBR_ADMIN.IT.Servicio.Instance.Dock = DockStyle.Fill;
                        CBR_ADMIN.IT.Servicio.Instance.BringToFront();


                    }
                    else
                    {
                        CBR_ADMIN.IT.Servicio.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }



        }

        private void soporteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (!Central.Controls.Contains(IT.Soporte.Instance))
                {
                    CBR_ADMIN.IT.Soporte.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Soporte.Instance.Apellido = Apellido;
                    Central.Controls.Add(CBR_ADMIN.IT.Soporte.Instance);
                    CBR_ADMIN.IT.Soporte.Instance.Dock = DockStyle.Fill;
                    CBR_ADMIN.IT.Soporte.Instance.BringToFront();


                }
                else
                {
                    CBR_ADMIN.IT.Soporte.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //    MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {
                        //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);    
                    }
                    oNetDrive = null;
                    if (!Central.Controls.Contains(IT.Soporte.Instance))
                    {
                        CBR_ADMIN.IT.Soporte.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Soporte.Instance.Apellido = Apellido;
                        Central.Controls.Add(CBR_ADMIN.IT.Soporte.Instance);
                        CBR_ADMIN.IT.Soporte.Instance.Dock = DockStyle.Fill;
                        CBR_ADMIN.IT.Soporte.Instance.BringToFront();


                    }
                    else
                    {
                        CBR_ADMIN.IT.Soporte.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }



        }

        private void toolStripMenuItem78_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (!Central2.Controls.Contains(IT.Administracion.Instance))
                {
                    CBR_ADMIN.IT.Administracion.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Administracion.Instance.Apellido = Apellido;
                    Central2.Controls.Add(CBR_ADMIN.IT.Administracion.Instance);
                    CBR_ADMIN.IT.Administracion.Instance.Dock = DockStyle.Fill;
                    CBR_ADMIN.IT.Administracion.Instance.BringToFront();


                }
                else
                {
                    CBR_ADMIN.IT.Administracion.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); 
                    }
                    oNetDrive = null;
                    if (!Central2.Controls.Contains(IT.Administracion.Instance))
                    {
                        CBR_ADMIN.IT.Administracion.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Administracion.Instance.Apellido = Apellido;
                        Central2.Controls.Add(CBR_ADMIN.IT.Administracion.Instance);
                        CBR_ADMIN.IT.Administracion.Instance.Dock = DockStyle.Fill;
                        CBR_ADMIN.IT.Administracion.Instance.BringToFront();


                    }
                    else
                    {
                        CBR_ADMIN.IT.Administracion.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }



        }

        private void toolStripMenuItem79_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (!Central2.Controls.Contains(IT.Aplicativo.Instance))
                {

                    CBR_ADMIN.IT.Aplicativo.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Aplicativo.Instance.Apellido = Apellido;
                    Central2.Controls.Add(IT.Aplicativo.Instance);
                    IT.Aplicativo.Instance.Dock = DockStyle.Fill;
                    IT.Aplicativo.Instance.BringToFront();


                }
                else
                {
                    IT.Administracion.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //    MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); 
                    }
                    oNetDrive = null;
                    if (!Central2.Controls.Contains(IT.Aplicativo.Instance))
                    {

                        CBR_ADMIN.IT.Aplicativo.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Aplicativo.Instance.Apellido = Apellido;
                        Central2.Controls.Add(IT.Aplicativo.Instance);
                        IT.Aplicativo.Instance.Dock = DockStyle.Fill;
                        IT.Aplicativo.Instance.BringToFront();


                    }
                    else
                    {
                        IT.Administracion.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void toolStripMenuItem80_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;

                if (!Central2.Controls.Contains(IT.Optimizacion.Instance))
                {

                    CBR_ADMIN.IT.Optimizacion.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Optimizacion.Instance.Apellido = Apellido;
                    Central2.Controls.Add(CBR_ADMIN.IT.Optimizacion.Instance);
                    CBR_ADMIN.IT.Optimizacion.Instance.Dock = DockStyle.Fill;
                    CBR_ADMIN.IT.Optimizacion.Instance.BringToFront();


                }
                else
                {
                    CBR_ADMIN.IT.Optimizacion.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //      MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {//  MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);    
                    }
                    oNetDrive = null;

                    if (!Central2.Controls.Contains(IT.Optimizacion.Instance))
                    {

                        CBR_ADMIN.IT.Optimizacion.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Optimizacion.Instance.Apellido = Apellido;
                        Central2.Controls.Add(CBR_ADMIN.IT.Optimizacion.Instance);
                        CBR_ADMIN.IT.Optimizacion.Instance.Dock = DockStyle.Fill;
                        CBR_ADMIN.IT.Optimizacion.Instance.BringToFront();


                    }
                    else
                    {
                        CBR_ADMIN.IT.Optimizacion.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }

        }

        private void toolStripMenuItem81_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (!Central2.Controls.Contains(IT.Otros.Instance))
                {
                    CBR_ADMIN.IT.Otros.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Otros.Instance.Apellido = Apellido;
                    Central2.Controls.Add(CBR_ADMIN.IT.Otros.Instance);
                    CBR_ADMIN.IT.Otros.Instance.Dock = DockStyle.Fill;
                    CBR_ADMIN.IT.Otros.Instance.BringToFront();


                }
                else
                {
                    CBR_ADMIN.IT.Otros.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;
                    if (!Central2.Controls.Contains(IT.Otros.Instance))
                    {
                        CBR_ADMIN.IT.Otros.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Otros.Instance.Apellido = Apellido;
                        Central2.Controls.Add(CBR_ADMIN.IT.Otros.Instance);
                        CBR_ADMIN.IT.Otros.Instance.Dock = DockStyle.Fill;
                        CBR_ADMIN.IT.Otros.Instance.BringToFront();


                    }
                    else
                    {
                        CBR_ADMIN.IT.Otros.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void toolStripMenuItem82_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (!Central2.Controls.Contains(IT.Servicio.Instance))
                {

                    CBR_ADMIN.IT.Soporte.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Soporte.Instance.Apellido = Apellido;
                    Central2.Controls.Add(CBR_ADMIN.IT.Servicio.Instance);
                    CBR_ADMIN.IT.Servicio.Instance.Dock = DockStyle.Fill;
                    CBR_ADMIN.IT.Servicio.Instance.BringToFront();


                }
                else
                {
                    CBR_ADMIN.IT.Servicio.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;
                    if (!Central2.Controls.Contains(IT.Servicio.Instance))
                    {

                        CBR_ADMIN.IT.Soporte.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Soporte.Instance.Apellido = Apellido;
                        Central2.Controls.Add(CBR_ADMIN.IT.Servicio.Instance);
                        CBR_ADMIN.IT.Servicio.Instance.Dock = DockStyle.Fill;
                        CBR_ADMIN.IT.Servicio.Instance.BringToFront();


                    }
                    else
                    {
                        CBR_ADMIN.IT.Servicio.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }



        }

        private void toolStripMenuItem83_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (!Central2.Controls.Contains(IT.Soporte.Instance))
                {
                    CBR_ADMIN.IT.Servicio.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Servicio.Instance.Apellido = Apellido;

                    Central2.Controls.Add(CBR_ADMIN.IT.Soporte.Instance);
                    CBR_ADMIN.IT.Soporte.Instance.Dock = DockStyle.Fill;
                    CBR_ADMIN.IT.Soporte.Instance.BringToFront();


                }
                else
                {
                    CBR_ADMIN.IT.Soporte.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); 
                    }
                    oNetDrive = null;
                    if (!Central2.Controls.Contains(IT.Soporte.Instance))
                    {
                        CBR_ADMIN.IT.Servicio.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Servicio.Instance.Apellido = Apellido;

                        Central2.Controls.Add(CBR_ADMIN.IT.Soporte.Instance);
                        CBR_ADMIN.IT.Soporte.Instance.Dock = DockStyle.Fill;
                        CBR_ADMIN.IT.Soporte.Instance.BringToFront();


                    }
                    else
                    {
                        CBR_ADMIN.IT.Soporte.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void toolStripMenuItem85_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (!Central3.Controls.Contains(IT.Administracion.Instance))
                {
                    CBR_ADMIN.IT.Administracion.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Administracion.Instance.Apellido = Apellido;


                    Central3.Controls.Add(CBR_ADMIN.IT.Administracion.Instance);
                    CBR_ADMIN.IT.Administracion.Instance.Dock = DockStyle.Fill;
                    CBR_ADMIN.IT.Administracion.Instance.BringToFront();


                }
                else
                {
                    CBR_ADMIN.IT.Administracion.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); 
                    }
                    oNetDrive = null;
                    if (!Central3.Controls.Contains(IT.Administracion.Instance))
                    {
                        CBR_ADMIN.IT.Administracion.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Administracion.Instance.Apellido = Apellido;


                        Central3.Controls.Add(CBR_ADMIN.IT.Administracion.Instance);
                        CBR_ADMIN.IT.Administracion.Instance.Dock = DockStyle.Fill;
                        CBR_ADMIN.IT.Administracion.Instance.BringToFront();


                    }
                    else
                    {
                        CBR_ADMIN.IT.Administracion.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }



        }

        private void toolStripMenuItem86_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;

                if (!Central3.Controls.Contains(IT.Aplicativo.Instance))
                {

                    CBR_ADMIN.IT.Aplicativo.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Aplicativo.Instance.Apellido = Apellido;
                    Central3.Controls.Add(IT.Aplicativo.Instance);
                    IT.Aplicativo.Instance.Dock = DockStyle.Fill;
                    IT.Aplicativo.Instance.BringToFront();


                }
                else
                {
                    IT.Administracion.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); 
                    }
                    oNetDrive = null;

                    if (!Central3.Controls.Contains(IT.Aplicativo.Instance))
                    {

                        CBR_ADMIN.IT.Aplicativo.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Aplicativo.Instance.Apellido = Apellido;
                        Central3.Controls.Add(IT.Aplicativo.Instance);
                        IT.Aplicativo.Instance.Dock = DockStyle.Fill;
                        IT.Aplicativo.Instance.BringToFront();


                    }
                    else
                    {
                        IT.Administracion.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }

        }

        private void toolStripMenuItem87_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (!Central3.Controls.Contains(IT.Optimizacion.Instance))
                {

                    CBR_ADMIN.IT.Optimizacion.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Optimizacion.Instance.Apellido = Apellido;
                    Central3.Controls.Add(CBR_ADMIN.IT.Optimizacion.Instance);
                    CBR_ADMIN.IT.Optimizacion.Instance.Dock = DockStyle.Fill;
                    CBR_ADMIN.IT.Optimizacion.Instance.BringToFront();


                }
                else
                {
                    CBR_ADMIN.IT.Optimizacion.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {//  MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);   
                    }
                    oNetDrive = null;
                    if (!Central3.Controls.Contains(IT.Optimizacion.Instance))
                    {

                        CBR_ADMIN.IT.Optimizacion.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Optimizacion.Instance.Apellido = Apellido;
                        Central3.Controls.Add(CBR_ADMIN.IT.Optimizacion.Instance);
                        CBR_ADMIN.IT.Optimizacion.Instance.Dock = DockStyle.Fill;
                        CBR_ADMIN.IT.Optimizacion.Instance.BringToFront();


                    }
                    else
                    {
                        CBR_ADMIN.IT.Optimizacion.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }



        }

        private void toolStripMenuItem88_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (!Central3.Controls.Contains(IT.Otros.Instance))
                {
                    CBR_ADMIN.IT.Otros.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Otros.Instance.Apellido = Apellido;
                    Central3.Controls.Add(CBR_ADMIN.IT.Otros.Instance);
                    CBR_ADMIN.IT.Otros.Instance.Dock = DockStyle.Fill;
                    CBR_ADMIN.IT.Otros.Instance.BringToFront();


                }
                else
                {
                    CBR_ADMIN.IT.Otros.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);     

                    }
                    oNetDrive = null;
                    if (!Central3.Controls.Contains(IT.Otros.Instance))
                    {
                        CBR_ADMIN.IT.Otros.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Otros.Instance.Apellido = Apellido;
                        Central3.Controls.Add(CBR_ADMIN.IT.Otros.Instance);
                        CBR_ADMIN.IT.Otros.Instance.Dock = DockStyle.Fill;
                        CBR_ADMIN.IT.Otros.Instance.BringToFront();


                    }
                    else
                    {
                        CBR_ADMIN.IT.Otros.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }



        }

        private void toolStripMenuItem89_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (!Central3.Controls.Contains(IT.Servicio.Instance))
                {
                    CBR_ADMIN.IT.Servicio.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Servicio.Instance.Apellido = Apellido;
                    Central3.Controls.Add(CBR_ADMIN.IT.Servicio.Instance);
                    CBR_ADMIN.IT.Servicio.Instance.Dock = DockStyle.Fill;
                    CBR_ADMIN.IT.Servicio.Instance.BringToFront();


                }
                else
                {
                    CBR_ADMIN.IT.Servicio.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;
                    if (!Central3.Controls.Contains(IT.Servicio.Instance))
                    {
                        CBR_ADMIN.IT.Servicio.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Servicio.Instance.Apellido = Apellido;
                        Central3.Controls.Add(CBR_ADMIN.IT.Servicio.Instance);
                        CBR_ADMIN.IT.Servicio.Instance.Dock = DockStyle.Fill;
                        CBR_ADMIN.IT.Servicio.Instance.BringToFront();


                    }
                    else
                    {
                        CBR_ADMIN.IT.Servicio.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void toolStripMenuItem90_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (!Central3.Controls.Contains(IT.Soporte.Instance))
                {
                    CBR_ADMIN.IT.Soporte.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Soporte.Instance.Apellido = Apellido;
                    Central3.Controls.Add(CBR_ADMIN.IT.Soporte.Instance);
                    CBR_ADMIN.IT.Soporte.Instance.Dock = DockStyle.Fill;
                    CBR_ADMIN.IT.Soporte.Instance.BringToFront();


                }
                else
                {
                    CBR_ADMIN.IT.Soporte.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);     
                    }
                    oNetDrive = null;
                    if (!Central3.Controls.Contains(IT.Soporte.Instance))
                    {
                        CBR_ADMIN.IT.Soporte.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Soporte.Instance.Apellido = Apellido;
                        Central3.Controls.Add(CBR_ADMIN.IT.Soporte.Instance);
                        CBR_ADMIN.IT.Soporte.Instance.Dock = DockStyle.Fill;
                        CBR_ADMIN.IT.Soporte.Instance.BringToFront();


                    }
                    else
                    {
                        CBR_ADMIN.IT.Soporte.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void toolStripMenuItem92_Click(object sender, EventArgs e)
        {

            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;

                if (!Central4.Controls.Contains(IT.Administracion.Instance))
                {
                    CBR_ADMIN.IT.Administracion.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Administracion.Instance.Apellido = Apellido;
                    Central4.Controls.Add(CBR_ADMIN.IT.Administracion.Instance);
                    CBR_ADMIN.IT.Administracion.Instance.Dock = DockStyle.Fill;
                    CBR_ADMIN.IT.Administracion.Instance.BringToFront();


                }
                else
                {
                    CBR_ADMIN.IT.Administracion.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);    
                    }
                    oNetDrive = null;

                    if (!Central4.Controls.Contains(IT.Administracion.Instance))
                    {
                        CBR_ADMIN.IT.Administracion.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Administracion.Instance.Apellido = Apellido;
                        Central4.Controls.Add(CBR_ADMIN.IT.Administracion.Instance);
                        CBR_ADMIN.IT.Administracion.Instance.Dock = DockStyle.Fill;
                        CBR_ADMIN.IT.Administracion.Instance.BringToFront();


                    }
                    else
                    {
                        CBR_ADMIN.IT.Administracion.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }




        }

        private void toolStripMenuItem93_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (!Central4.Controls.Contains(IT.Aplicativo.Instance))
                {
                    CBR_ADMIN.IT.Aplicativo.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Aplicativo.Instance.Apellido = Apellido;
                    Central4.Controls.Add(IT.Aplicativo.Instance);
                    IT.Aplicativo.Instance.Dock = DockStyle.Fill;
                    IT.Aplicativo.Instance.BringToFront();


                }
                else
                {
                    IT.Administracion.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //     MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);     
                    }
                    oNetDrive = null;
                    if (!Central4.Controls.Contains(IT.Aplicativo.Instance))
                    {
                        CBR_ADMIN.IT.Aplicativo.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Aplicativo.Instance.Apellido = Apellido;
                        Central4.Controls.Add(IT.Aplicativo.Instance);
                        IT.Aplicativo.Instance.Dock = DockStyle.Fill;
                        IT.Aplicativo.Instance.BringToFront();


                    }
                    else
                    {
                        IT.Administracion.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void toolStripMenuItem94_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;

                if (!Central4.Controls.Contains(IT.Optimizacion.Instance))
                {
                    CBR_ADMIN.IT.Optimizacion.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Optimizacion.Instance.Apellido = Apellido;
                    Central4.Controls.Add(CBR_ADMIN.IT.Optimizacion.Instance);
                    CBR_ADMIN.IT.Optimizacion.Instance.Dock = DockStyle.Fill;
                    CBR_ADMIN.IT.Optimizacion.Instance.BringToFront();


                }
                else
                {
                    CBR_ADMIN.IT.Optimizacion.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);
                    }
                    oNetDrive = null;

                    if (!Central4.Controls.Contains(IT.Optimizacion.Instance))
                    {
                        CBR_ADMIN.IT.Optimizacion.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Optimizacion.Instance.Apellido = Apellido;
                        Central4.Controls.Add(CBR_ADMIN.IT.Optimizacion.Instance);
                        CBR_ADMIN.IT.Optimizacion.Instance.Dock = DockStyle.Fill;
                        CBR_ADMIN.IT.Optimizacion.Instance.BringToFront();


                    }
                    else
                    {
                        CBR_ADMIN.IT.Optimizacion.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }

        }

        private void toolStripMenuItem95_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (!Central4.Controls.Contains(IT.Otros.Instance))
                {
                    CBR_ADMIN.IT.Otros.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Otros.Instance.Apellido = Apellido;

                    Central4.Controls.Add(CBR_ADMIN.IT.Otros.Instance);
                    CBR_ADMIN.IT.Otros.Instance.Dock = DockStyle.Fill;
                    CBR_ADMIN.IT.Otros.Instance.BringToFront();


                }
                else
                {
                    CBR_ADMIN.IT.Otros.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);    
                    }
                    oNetDrive = null;
                    if (!Central4.Controls.Contains(IT.Otros.Instance))
                    {
                        CBR_ADMIN.IT.Otros.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Otros.Instance.Apellido = Apellido;

                        Central4.Controls.Add(CBR_ADMIN.IT.Otros.Instance);
                        CBR_ADMIN.IT.Otros.Instance.Dock = DockStyle.Fill;
                        CBR_ADMIN.IT.Otros.Instance.BringToFront();


                    }
                    else
                    {
                        CBR_ADMIN.IT.Otros.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void toolStripMenuItem96_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (!Central4.Controls.Contains(IT.Servicio.Instance))
                {
                    CBR_ADMIN.IT.Servicio.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Servicio.Instance.Apellido = Apellido;
                    Central4.Controls.Add(CBR_ADMIN.IT.Servicio.Instance);
                    CBR_ADMIN.IT.Servicio.Instance.Dock = DockStyle.Fill;
                    CBR_ADMIN.IT.Servicio.Instance.BringToFront();


                }
                else
                {
                    CBR_ADMIN.IT.Servicio.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);    
                    }
                    oNetDrive = null;
                    if (!Central4.Controls.Contains(IT.Servicio.Instance))
                    {
                        CBR_ADMIN.IT.Servicio.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Servicio.Instance.Apellido = Apellido;
                        Central4.Controls.Add(CBR_ADMIN.IT.Servicio.Instance);
                        CBR_ADMIN.IT.Servicio.Instance.Dock = DockStyle.Fill;
                        CBR_ADMIN.IT.Servicio.Instance.BringToFront();


                    }
                    else
                    {
                        CBR_ADMIN.IT.Servicio.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }



        }

        private void toolStripMenuItem97_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;

                if (!Central4.Controls.Contains(IT.Soporte.Instance))
                {
                    CBR_ADMIN.IT.Soporte.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Soporte.Instance.Apellido = Apellido;
                    Central4.Controls.Add(CBR_ADMIN.IT.Soporte.Instance);
                    CBR_ADMIN.IT.Soporte.Instance.Dock = DockStyle.Fill;
                    CBR_ADMIN.IT.Soporte.Instance.BringToFront();


                }
                else
                {
                    CBR_ADMIN.IT.Soporte.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;

                    if (!Central4.Controls.Contains(IT.Soporte.Instance))
                    {
                        CBR_ADMIN.IT.Soporte.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Soporte.Instance.Apellido = Apellido;
                        Central4.Controls.Add(CBR_ADMIN.IT.Soporte.Instance);
                        CBR_ADMIN.IT.Soporte.Instance.Dock = DockStyle.Fill;
                        CBR_ADMIN.IT.Soporte.Instance.BringToFront();


                    }
                    else
                    {
                        CBR_ADMIN.IT.Soporte.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }

        }

        private void toolStripMenuItem99_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;

                if (!Central5.Controls.Contains(IT.Administracion.Instance))
                {
                    CBR_ADMIN.IT.Administracion.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Administracion.Instance.Apellido = Apellido;
                    Central5.Controls.Add(CBR_ADMIN.IT.Administracion.Instance);
                    CBR_ADMIN.IT.Administracion.Instance.Dock = DockStyle.Fill;
                    CBR_ADMIN.IT.Administracion.Instance.BringToFront();


                }
                else
                {
                    CBR_ADMIN.IT.Administracion.Instance.BringToFront();

                }

            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); 
                    }
                    oNetDrive = null;

                    if (!Central5.Controls.Contains(IT.Administracion.Instance))
                    {
                        CBR_ADMIN.IT.Administracion.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Administracion.Instance.Apellido = Apellido;
                        Central5.Controls.Add(CBR_ADMIN.IT.Administracion.Instance);
                        CBR_ADMIN.IT.Administracion.Instance.Dock = DockStyle.Fill;
                        CBR_ADMIN.IT.Administracion.Instance.BringToFront();


                    }
                    else
                    {
                        CBR_ADMIN.IT.Administracion.Instance.BringToFront();

                    }

                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }








        }

        private void toolStripMenuItem100_Click(object sender, EventArgs e)
        {


            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;

                if (!Central5.Controls.Contains(IT.Aplicativo.Instance))
                {
                    CBR_ADMIN.IT.Aplicativo.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Aplicativo.Instance.Apellido = Apellido;
                    Central5.Controls.Add(IT.Aplicativo.Instance);
                    IT.Aplicativo.Instance.Dock = DockStyle.Fill;
                    IT.Aplicativo.Instance.BringToFront();


                }
                else
                {
                    IT.Administracion.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //     MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);   
                    }
                    oNetDrive = null;

                    if (!Central5.Controls.Contains(IT.Aplicativo.Instance))
                    {
                        CBR_ADMIN.IT.Aplicativo.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Aplicativo.Instance.Apellido = Apellido;
                        Central5.Controls.Add(IT.Aplicativo.Instance);
                        IT.Aplicativo.Instance.Dock = DockStyle.Fill;
                        IT.Aplicativo.Instance.BringToFront();


                    }
                    else
                    {
                        IT.Administracion.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }










        }

        private void toolStripMenuItem101_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;



                if (!Central5.Controls.Contains(IT.Optimizacion.Instance))
                {
                    CBR_ADMIN.IT.Optimizacion.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Optimizacion.Instance.Apellido = Apellido;
                    Central5.Controls.Add(CBR_ADMIN.IT.Optimizacion.Instance);
                    CBR_ADMIN.IT.Optimizacion.Instance.Dock = DockStyle.Fill;
                    CBR_ADMIN.IT.Optimizacion.Instance.BringToFront();


                }
                else
                {
                    CBR_ADMIN.IT.Optimizacion.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);
                    }
                    oNetDrive = null;



                    if (!Central5.Controls.Contains(IT.Optimizacion.Instance))
                    {
                        CBR_ADMIN.IT.Optimizacion.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Optimizacion.Instance.Apellido = Apellido;
                        Central5.Controls.Add(CBR_ADMIN.IT.Optimizacion.Instance);
                        CBR_ADMIN.IT.Optimizacion.Instance.Dock = DockStyle.Fill;
                        CBR_ADMIN.IT.Optimizacion.Instance.BringToFront();


                    }
                    else
                    {
                        CBR_ADMIN.IT.Optimizacion.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }






        }

        private void toolStripMenuItem102_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (!Central5.Controls.Contains(IT.Otros.Instance))
                {
                    CBR_ADMIN.IT.Otros.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Otros.Instance.Apellido = Apellido;
                    Central5.Controls.Add(CBR_ADMIN.IT.Otros.Instance);
                    CBR_ADMIN.IT.Otros.Instance.Dock = DockStyle.Fill;
                    CBR_ADMIN.IT.Otros.Instance.BringToFront();


                }
                else
                {
                    CBR_ADMIN.IT.Otros.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);    
                    }
                    oNetDrive = null;

                    if (!Central5.Controls.Contains(IT.Otros.Instance))
                    {
                        CBR_ADMIN.IT.Otros.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Otros.Instance.Apellido = Apellido;
                        Central5.Controls.Add(CBR_ADMIN.IT.Otros.Instance);
                        CBR_ADMIN.IT.Otros.Instance.Dock = DockStyle.Fill;
                        CBR_ADMIN.IT.Otros.Instance.BringToFront();


                    }
                    else
                    {
                        CBR_ADMIN.IT.Otros.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }










        }

        private void toolStripMenuItem103_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;


                if (!Central5.Controls.Contains(IT.Servicio.Instance))
                {
                    CBR_ADMIN.IT.Servicio.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Servicio.Instance.Apellido = Apellido;
                    Central5.Controls.Add(CBR_ADMIN.IT.Servicio.Instance);
                    CBR_ADMIN.IT.Servicio.Instance.Dock = DockStyle.Fill;
                    CBR_ADMIN.IT.Servicio.Instance.BringToFront();


                }
                else
                {
                    CBR_ADMIN.IT.Servicio.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); 
                    }
                    oNetDrive = null;


                    if (!Central5.Controls.Contains(IT.Servicio.Instance))
                    {
                        CBR_ADMIN.IT.Servicio.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Servicio.Instance.Apellido = Apellido;
                        Central5.Controls.Add(CBR_ADMIN.IT.Servicio.Instance);
                        CBR_ADMIN.IT.Servicio.Instance.Dock = DockStyle.Fill;
                        CBR_ADMIN.IT.Servicio.Instance.BringToFront();


                    }
                    else
                    {
                        CBR_ADMIN.IT.Servicio.Instance.BringToFront();

                    }

                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }








        }

        private void toolStripMenuItem104_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2; if (!Central5.Controls.Contains(IT.Soporte.Instance))
                {
                    CBR_ADMIN.IT.Soporte.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Soporte.Instance.Apellido = Apellido;
                    Central5.Controls.Add(CBR_ADMIN.IT.Soporte.Instance);
                    CBR_ADMIN.IT.Soporte.Instance.Dock = DockStyle.Fill;
                    CBR_ADMIN.IT.Soporte.Instance.BringToFront();


                }
                else
                {
                    CBR_ADMIN.IT.Soporte.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); 
                    }
                    oNetDrive = null;
                    if (!Central5.Controls.Contains(IT.Soporte.Instance))
                    {
                        CBR_ADMIN.IT.Soporte.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Soporte.Instance.Apellido = Apellido;
                        Central5.Controls.Add(CBR_ADMIN.IT.Soporte.Instance);
                        CBR_ADMIN.IT.Soporte.Instance.Dock = DockStyle.Fill;
                        CBR_ADMIN.IT.Soporte.Instance.BringToFront();


                    }
                    else
                    {
                        CBR_ADMIN.IT.Soporte.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }









        }

        private void toolStripMenuItem106_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (!Central6.Controls.Contains(IT.Administracion.Instance))
                {
                    CBR_ADMIN.IT.Administracion.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Administracion.Instance.Apellido = Apellido;
                    Central6.Controls.Add(CBR_ADMIN.IT.Administracion.Instance);
                    CBR_ADMIN.IT.Administracion.Instance.Dock = DockStyle.Fill;
                    CBR_ADMIN.IT.Administracion.Instance.BringToFront();


                }
                else
                {
                    CBR_ADMIN.IT.Administracion.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {//  MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);
                    }
                    oNetDrive = null;
                    if (!Central6.Controls.Contains(IT.Administracion.Instance))
                    {
                        CBR_ADMIN.IT.Administracion.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Administracion.Instance.Apellido = Apellido;
                        Central6.Controls.Add(CBR_ADMIN.IT.Administracion.Instance);
                        CBR_ADMIN.IT.Administracion.Instance.Dock = DockStyle.Fill;
                        CBR_ADMIN.IT.Administracion.Instance.BringToFront();


                    }
                    else
                    {
                        CBR_ADMIN.IT.Administracion.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }



        }

        private void toolStripMenuItem107_Click(object sender, EventArgs e)
        {

            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (!Central6.Controls.Contains(IT.Aplicativo.Instance))
                {

                    CBR_ADMIN.IT.Aplicativo.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Aplicativo.Instance.Apellido = Apellido;
                    Central6.Controls.Add(IT.Aplicativo.Instance);
                    IT.Aplicativo.Instance.Dock = DockStyle.Fill;
                    IT.Aplicativo.Instance.BringToFront();


                }
                else
                {
                    IT.Administracion.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); 
                    }
                    oNetDrive = null;
                    if (!Central6.Controls.Contains(IT.Aplicativo.Instance))
                    {

                        CBR_ADMIN.IT.Aplicativo.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Aplicativo.Instance.Apellido = Apellido;
                        Central6.Controls.Add(IT.Aplicativo.Instance);
                        IT.Aplicativo.Instance.Dock = DockStyle.Fill;
                        IT.Aplicativo.Instance.BringToFront();


                    }
                    else
                    {
                        IT.Administracion.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }




        }

        private void toolStripMenuItem108_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;

                if (!Central6.Controls.Contains(IT.Optimizacion.Instance))
                {
                    CBR_ADMIN.IT.Optimizacion.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Optimizacion.Instance.Apellido = Apellido;
                    Central6.Controls.Add(CBR_ADMIN.IT.Optimizacion.Instance);
                    CBR_ADMIN.IT.Optimizacion.Instance.Dock = DockStyle.Fill;
                    CBR_ADMIN.IT.Optimizacion.Instance.BringToFront();


                }
                else
                {
                    CBR_ADMIN.IT.Optimizacion.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;

                    if (!Central6.Controls.Contains(IT.Optimizacion.Instance))
                    {
                        CBR_ADMIN.IT.Optimizacion.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Optimizacion.Instance.Apellido = Apellido;
                        Central6.Controls.Add(CBR_ADMIN.IT.Optimizacion.Instance);
                        CBR_ADMIN.IT.Optimizacion.Instance.Dock = DockStyle.Fill;
                        CBR_ADMIN.IT.Optimizacion.Instance.BringToFront();


                    }
                    else
                    {
                        CBR_ADMIN.IT.Optimizacion.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }




        }

        private void toolStripMenuItem109_Click(object sender, EventArgs e)
        {

            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (!Central6.Controls.Contains(IT.Otros.Instance))
                {
                    CBR_ADMIN.IT.Otros.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Otros.Instance.Apellido = Apellido;

                    Central6.Controls.Add(CBR_ADMIN.IT.Otros.Instance);
                    CBR_ADMIN.IT.Otros.Instance.Dock = DockStyle.Fill;
                    CBR_ADMIN.IT.Otros.Instance.BringToFront();


                }
                else
                {
                    CBR_ADMIN.IT.Otros.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;
                    if (!Central6.Controls.Contains(IT.Otros.Instance))
                    {
                        CBR_ADMIN.IT.Otros.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Otros.Instance.Apellido = Apellido;

                        Central6.Controls.Add(CBR_ADMIN.IT.Otros.Instance);
                        CBR_ADMIN.IT.Otros.Instance.Dock = DockStyle.Fill;
                        CBR_ADMIN.IT.Otros.Instance.BringToFront();


                    }
                    else
                    {
                        CBR_ADMIN.IT.Otros.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }



        }

        private void toolStripMenuItem110_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (!Central6.Controls.Contains(IT.Servicio.Instance))
                {
                    CBR_ADMIN.IT.Servicio.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Servicio.Instance.Apellido = Apellido;
                    Central6.Controls.Add(CBR_ADMIN.IT.Servicio.Instance);
                    CBR_ADMIN.IT.Servicio.Instance.Dock = DockStyle.Fill;
                    CBR_ADMIN.IT.Servicio.Instance.BringToFront();


                }
                else
                {
                    CBR_ADMIN.IT.Servicio.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); 
                    }
                    oNetDrive = null;
                    if (!Central6.Controls.Contains(IT.Servicio.Instance))
                    {
                        CBR_ADMIN.IT.Servicio.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Servicio.Instance.Apellido = Apellido;
                        Central6.Controls.Add(CBR_ADMIN.IT.Servicio.Instance);
                        CBR_ADMIN.IT.Servicio.Instance.Dock = DockStyle.Fill;
                        CBR_ADMIN.IT.Servicio.Instance.BringToFront();


                    }
                    else
                    {
                        CBR_ADMIN.IT.Servicio.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void toolStripMenuItem111_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (!Central6.Controls.Contains(IT.Soporte.Instance))
                {
                    CBR_ADMIN.IT.Soporte.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Soporte.Instance.Apellido = Apellido;

                    Central6.Controls.Add(CBR_ADMIN.IT.Soporte.Instance);
                    CBR_ADMIN.IT.Soporte.Instance.Dock = DockStyle.Fill;
                    CBR_ADMIN.IT.Soporte.Instance.BringToFront();


                }
                else
                {
                    CBR_ADMIN.IT.Soporte.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //     MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); 
                    }
                    oNetDrive = null;
                    if (!Central6.Controls.Contains(IT.Soporte.Instance))
                    {
                        CBR_ADMIN.IT.Soporte.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Soporte.Instance.Apellido = Apellido;

                        Central6.Controls.Add(CBR_ADMIN.IT.Soporte.Instance);
                        CBR_ADMIN.IT.Soporte.Instance.Dock = DockStyle.Fill;
                        CBR_ADMIN.IT.Soporte.Instance.BringToFront();


                    }
                    else
                    {
                        CBR_ADMIN.IT.Soporte.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void toolStripMenuItem113_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (!Central7.Controls.Contains(IT.Administracion.Instance))
                {
                    CBR_ADMIN.IT.Administracion.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Administracion.Instance.Apellido = Apellido;

                    Central7.Controls.Add(CBR_ADMIN.IT.Administracion.Instance);
                    CBR_ADMIN.IT.Administracion.Instance.Dock = DockStyle.Fill;
                    CBR_ADMIN.IT.Administracion.Instance.BringToFront();


                }
                else
                {
                    CBR_ADMIN.IT.Administracion.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); 
                    }
                    oNetDrive = null;

                    if (!Central7.Controls.Contains(IT.Administracion.Instance))
                    {
                        CBR_ADMIN.IT.Administracion.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Administracion.Instance.Apellido = Apellido;

                        Central7.Controls.Add(CBR_ADMIN.IT.Administracion.Instance);
                        CBR_ADMIN.IT.Administracion.Instance.Dock = DockStyle.Fill;
                        CBR_ADMIN.IT.Administracion.Instance.BringToFront();


                    }
                    else
                    {
                        CBR_ADMIN.IT.Administracion.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }



        }

        private void toolStripMenuItem114_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (!Central7.Controls.Contains(IT.Aplicativo.Instance))
                {
                    CBR_ADMIN.IT.Aplicativo.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Aplicativo.Instance.Apellido = Apellido;
                    Central7.Controls.Add(IT.Aplicativo.Instance);
                    IT.Aplicativo.Instance.Dock = DockStyle.Fill;
                    IT.Aplicativo.Instance.BringToFront();


                }
                else
                {
                    IT.Administracion.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); 
                    }
                    oNetDrive = null;
                    if (!Central7.Controls.Contains(IT.Aplicativo.Instance))
                    {
                        CBR_ADMIN.IT.Aplicativo.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Aplicativo.Instance.Apellido = Apellido;
                        Central7.Controls.Add(IT.Aplicativo.Instance);
                        IT.Aplicativo.Instance.Dock = DockStyle.Fill;
                        IT.Aplicativo.Instance.BringToFront();


                    }
                    else
                    {
                        IT.Administracion.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }




        }

        private void toolStripMenuItem115_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;

                if (!Central7.Controls.Contains(IT.Optimizacion.Instance))
                {

                    CBR_ADMIN.IT.Optimizacion.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Optimizacion.Instance.Apellido = Apellido;
                    Central7.Controls.Add(CBR_ADMIN.IT.Optimizacion.Instance);
                    CBR_ADMIN.IT.Optimizacion.Instance.Dock = DockStyle.Fill;
                    CBR_ADMIN.IT.Optimizacion.Instance.BringToFront();


                }
                else
                {
                    CBR_ADMIN.IT.Optimizacion.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;

                    if (!Central7.Controls.Contains(IT.Optimizacion.Instance))
                    {

                        CBR_ADMIN.IT.Optimizacion.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Optimizacion.Instance.Apellido = Apellido;
                        Central7.Controls.Add(CBR_ADMIN.IT.Optimizacion.Instance);
                        CBR_ADMIN.IT.Optimizacion.Instance.Dock = DockStyle.Fill;
                        CBR_ADMIN.IT.Optimizacion.Instance.BringToFront();


                    }
                    else
                    {
                        CBR_ADMIN.IT.Optimizacion.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void toolStripMenuItem116_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;

                if (!Central7.Controls.Contains(IT.Otros.Instance))
                {
                    CBR_ADMIN.IT.Otros.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Otros.Instance.Apellido = Apellido;
                    Central7.Controls.Add(CBR_ADMIN.IT.Otros.Instance);
                    CBR_ADMIN.IT.Otros.Instance.Dock = DockStyle.Fill;
                    CBR_ADMIN.IT.Otros.Instance.BringToFront();


                }
                else
                {
                    CBR_ADMIN.IT.Otros.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            // MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;

                    if (!Central7.Controls.Contains(IT.Otros.Instance))
                    {
                        CBR_ADMIN.IT.Otros.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Otros.Instance.Apellido = Apellido;
                        Central7.Controls.Add(CBR_ADMIN.IT.Otros.Instance);
                        CBR_ADMIN.IT.Otros.Instance.Dock = DockStyle.Fill;
                        CBR_ADMIN.IT.Otros.Instance.BringToFront();


                    }
                    else
                    {
                        CBR_ADMIN.IT.Otros.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }




        }

        private void toolStripMenuItem117_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (!Central7.Controls.Contains(IT.Servicio.Instance))
                {
                    CBR_ADMIN.IT.Servicio.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Servicio.Instance.Apellido = Apellido;
                    Central7.Controls.Add(CBR_ADMIN.IT.Servicio.Instance);
                    CBR_ADMIN.IT.Servicio.Instance.Dock = DockStyle.Fill;
                    CBR_ADMIN.IT.Servicio.Instance.BringToFront();


                }
                else
                {
                    CBR_ADMIN.IT.Servicio.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);    
                    }
                    oNetDrive = null;
                    if (!Central7.Controls.Contains(IT.Servicio.Instance))
                    {
                        CBR_ADMIN.IT.Servicio.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Servicio.Instance.Apellido = Apellido;
                        Central7.Controls.Add(CBR_ADMIN.IT.Servicio.Instance);
                        CBR_ADMIN.IT.Servicio.Instance.Dock = DockStyle.Fill;
                        CBR_ADMIN.IT.Servicio.Instance.BringToFront();


                    }
                    else
                    {
                        CBR_ADMIN.IT.Servicio.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }



        }

        private void toolStripMenuItem118_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (!Central7.Controls.Contains(IT.Soporte.Instance))
                {
                    CBR_ADMIN.IT.Soporte.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Soporte.Instance.Apellido = Apellido;
                    Central7.Controls.Add(CBR_ADMIN.IT.Soporte.Instance);
                    CBR_ADMIN.IT.Soporte.Instance.Dock = DockStyle.Fill;
                    CBR_ADMIN.IT.Soporte.Instance.BringToFront();


                }
                else
                {
                    CBR_ADMIN.IT.Soporte.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //     MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;
                    if (!Central7.Controls.Contains(IT.Soporte.Instance))
                    {
                        CBR_ADMIN.IT.Soporte.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Soporte.Instance.Apellido = Apellido;
                        Central7.Controls.Add(CBR_ADMIN.IT.Soporte.Instance);
                        CBR_ADMIN.IT.Soporte.Instance.Dock = DockStyle.Fill;
                        CBR_ADMIN.IT.Soporte.Instance.BringToFront();


                    }
                    else
                    {
                        CBR_ADMIN.IT.Soporte.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }



        }

        private void cierreToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            if (!Central8.Controls.Contains(CierreTickets.Instance))
            {

                Central8.Controls.Add(CierreTickets.Instance);
                CierreTickets.Instance.Dock = DockStyle.Fill;
                CierreTickets.Instance.BringToFront();


            }
            else
            {
                CierreTickets.Instance.BringToFront();

            }
        }

        private void seguimientoToolStripMenuItem8_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Departamento == "Asistente de direccion" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                {


                    if (!Central.Controls.Contains(Seguimiento_No_confor.Instance))
                    {
                        Ventana = "Pantalla: Seguimiento de No conformidad";
                        Seguimiento_No_confor.Instance.Departamento = Departamento;
                        Seguimiento_No_confor.Instance.Nombre = Nombre;
                        Seguimiento_No_confor.Instance.Apellido = Apellido;
                        Seguimiento_No_confor.Instance.NoConfor = NumeroConformidad;
                        Seguimiento_No_confor.Instance.Ventana1 = Ventana;
                        Central.Controls.Add(Seguimiento_No_confor.Instance);
                        Seguimiento_No_confor.Instance.Dock = DockStyle.Fill;
                        Seguimiento_No_confor.Instance.BringToFront();


                    }
                    else
                    {
                        Seguimiento_No_confor.Instance.BringToFront();
                        Ventana = "Pantalla: Seguimiento de No conformidad";
                    }
                }


                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;
                    if (Departamento == "Asistente de direccion" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                    {


                        if (!Central.Controls.Contains(Seguimiento_No_confor.Instance))
                        {
                            Ventana = "Pantalla: Seguimiento de No conformidad";
                            Seguimiento_No_confor.Instance.Departamento = Departamento;
                            Seguimiento_No_confor.Instance.Nombre = Nombre;
                            Seguimiento_No_confor.Instance.Apellido = Apellido;
                            Seguimiento_No_confor.Instance.NoConfor = NumeroConformidad;
                            Seguimiento_No_confor.Instance.Ventana1 = Ventana;
                            Central.Controls.Add(Seguimiento_No_confor.Instance);
                            Seguimiento_No_confor.Instance.Dock = DockStyle.Fill;
                            Seguimiento_No_confor.Instance.BringToFront();


                        }
                        else
                        {
                            Seguimiento_No_confor.Instance.BringToFront();
                            Ventana = "Pantalla: Seguimiento de No conformidad";
                        }
                    }


                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void actualizacionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Departamento == "Asistente de direccion" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                {
                    if (Permiso == "1" || Permiso == "2")
                    {

                        if (!Central.Controls.Contains(Actualizacion_NC.Instance))
                        {

                            Central.Controls.Add(Actualizacion_NC.Instance);
                            Actualizacion_NC.Instance.Dock = DockStyle.Fill;
                            Actualizacion_NC.Instance.BringToFront();


                        }
                        else
                        {
                            Actualizacion_NC.Instance.BringToFront();
                            Ventana = "Pantalla: Genera Oportunidad de Mejora";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;
                    if (Departamento == "Asistente de direccion" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                    {
                        if (Permiso == "1" || Permiso == "2")
                        {

                            if (!Central.Controls.Contains(Actualizacion_NC.Instance))
                            {

                                Central.Controls.Add(Actualizacion_NC.Instance);
                                Actualizacion_NC.Instance.Dock = DockStyle.Fill;
                                Actualizacion_NC.Instance.BringToFront();


                            }
                            else
                            {
                                Actualizacion_NC.Instance.BringToFront();
                                Ventana = "Pantalla: Genera Oportunidad de Mejora";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }
            ThreadStart proceso = new ThreadStart(MostrarCuadroInicio3); Thread hilo = new Thread(proceso);


        }

        private void Central7_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void Nom_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void toolStripMenuItem33_DoubleClick(object sender, EventArgs e)
        {

        }

        private void noConformidadToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            if (Departamento == "Administracion" || Departamento == "SuperAdmin" || Departamento == "Direccion") { }
            else { MessageBox.Show("No tiene acceso a este modulo"); }
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else { conexionvpn = 2; }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }
        }

        private void administracionToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            if (Departamento == "Administracion" || Departamento == "SuperAdmin" || Departamento == "Direccion") { }
            else { MessageBox.Show("No tiene acceso a este modulo"); }
        }

        private void comprasToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else { conexionvpn = 2; }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }
            if (Departamento == "Compras" || Departamento == "SuperAdmin" || Departamento == "Direccion")
            { }
            else
            {
                MessageBox.Show("No tiene acceso a este modulo");
            }

        }

        private void generacionToolStripMenuItem1_Click_1(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Departamento == "Asistente de direccion" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                {

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central.Controls.Contains(Genera_AM.Instance))
                        {
                            Ventana = "Pantalla: Genera Oportunidad de Mejora";
                            Genera_AM.Instance.Departamento = Departamento;
                            Genera_AM.Instance.Nombre = Nombre;
                            Genera_AM.Instance.Ventana1 = Ventana;
                            Genera_AM.Instance.NumeroAC = NumeroOport;
                            Central.Controls.Add(Genera_AM.Instance);
                            Genera_AM.Instance.Dock = DockStyle.Fill;
                            Genera_AM.Instance.BringToFront();


                        }
                        else
                        {
                            Genera_AM.Instance.BringToFront();
                            Ventana = "Pantalla: Genera Oportunidad de Mejora";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }

                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //     MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);   
                    }
                    oNetDrive = null;

                    if (Departamento == "Asistente de direccion" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                    {

                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                        {

                            if (!Central.Controls.Contains(Genera_AM.Instance))
                            {
                                Ventana = "Pantalla: Genera Oportunidad de Mejora";
                                Genera_AM.Instance.Departamento = Departamento;
                                Genera_AM.Instance.Nombre = Nombre;
                                Genera_AM.Instance.Ventana1 = Ventana;
                                Genera_AM.Instance.NumeroAC = NumeroOport;
                                Central.Controls.Add(Genera_AM.Instance);
                                Genera_AM.Instance.Dock = DockStyle.Fill;
                                Genera_AM.Instance.BringToFront();


                            }
                            else
                            {
                                Genera_AM.Instance.BringToFront();
                                Ventana = "Pantalla: Genera Oportunidad de Mejora";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                    }

                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void generaciónToolStripMenuItem4_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Departamento == "Cobranza" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                {


                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central.Controls.Contains(Genera_No_Confor.Instance))
                        {
                            Ventana = "Pantalla: Levantamiento de No conformidad";
                            Genera_No_Confor.Instance.Nombre = Nombre;

                            Genera_No_Confor.Instance.Ventana1 = Ventana;
                            //Genera_No_Confor.Instance.Email = Email;
                            Central.Controls.Add(Genera_No_Confor.Instance);
                            Genera_No_Confor.Instance.Dock = DockStyle.Fill;
                            Genera_No_Confor.Instance.BringToFront();



                        }
                        else
                        {
                            Genera_No_Confor.Instance.BringToFront();
                            Ventana = "Pantalla: Levantamiento de No conformidad";


                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");

                    }
                }

                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //    MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;
                    if (Departamento == "Cobranza" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                    {


                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                        {

                            if (!Central.Controls.Contains(Genera_No_Confor.Instance))
                            {
                                Ventana = "Pantalla: Levantamiento de No conformidad";
                                Genera_No_Confor.Instance.Nombre = Nombre;

                                Genera_No_Confor.Instance.Ventana1 = Ventana;
                                //Genera_No_Confor.Instance.Email = Email;
                                Central.Controls.Add(Genera_No_Confor.Instance);
                                Genera_No_Confor.Instance.Dock = DockStyle.Fill;
                                Genera_No_Confor.Instance.BringToFront();



                            }
                            else
                            {
                                Genera_No_Confor.Instance.BringToFront();
                                Ventana = "Pantalla: Levantamiento de No conformidad";


                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");

                        }
                    }

                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void seguimientoToolStripMenuItem6_Click_1(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Departamento == "Cobranza" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                {


                    if (!Central.Controls.Contains(Seguimiento_No_confor.Instance))
                    {
                        Ventana = "Pantalla: Seguimiento de No conformidad";
                        Seguimiento_No_confor.Instance.Departamento = Departamento;
                        Seguimiento_No_confor.Instance.Nombre = Nombre;
                        Seguimiento_No_confor.Instance.Apellido = Apellido;
                        Seguimiento_No_confor.Instance.NoConfor = NumeroConformidad;
                        Seguimiento_No_confor.Instance.Ventana1 = Ventana;
                        Central.Controls.Add(Seguimiento_No_confor.Instance);
                        Seguimiento_No_confor.Instance.Dock = DockStyle.Fill;
                        Seguimiento_No_confor.Instance.BringToFront();


                    }
                    else
                    {
                        Seguimiento_No_confor.Instance.BringToFront();
                        Ventana = "Pantalla: Seguimiento de No conformidad";
                    }
                }


                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {//  MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;
                    if (Departamento == "Cobranza" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                    {


                        if (!Central.Controls.Contains(Seguimiento_No_confor.Instance))
                        {
                            Ventana = "Pantalla: Seguimiento de No conformidad";
                            Seguimiento_No_confor.Instance.Departamento = Departamento;
                            Seguimiento_No_confor.Instance.Nombre = Nombre;
                            Seguimiento_No_confor.Instance.Apellido = Apellido;
                            Seguimiento_No_confor.Instance.NoConfor = NumeroConformidad;
                            Seguimiento_No_confor.Instance.Ventana1 = Ventana;
                            Central.Controls.Add(Seguimiento_No_confor.Instance);
                            Seguimiento_No_confor.Instance.Dock = DockStyle.Fill;
                            Seguimiento_No_confor.Instance.BringToFront();


                        }
                        else
                        {
                            Seguimiento_No_confor.Instance.BringToFront();
                            Ventana = "Pantalla: Seguimiento de No conformidad";
                        }
                    }


                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }

        }

        private void asistenteDeDireccionToolStripMenuItem_Click(object sender, EventArgs e)
        {

            if (Departamento == "Asistente de direccion" || Departamento == "SuperAdmin" || Departamento == "Direccion")
            { }
            else
            {
                MessageBox.Show("No tiene acceso a este modulo");
            }
        }



        private void cobranzaToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            if (Departamento == "Cobranza" || Departamento == "SuperAdmin" || Departamento == "Direccion")
            { }
            else
            {
                MessageBox.Show("No tiene acceso a este modulo");
            }
        }

        private void generalToolStripMenuItem8_Click_1(object sender, EventArgs e)
        {

            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central7.Controls.Contains(Consulta_No_Conf_general.Instance))
                    {
                        Ventana = "Pantalla: Consulta de No Conformidad";

                        Central7.Controls.Add(Consulta_No_Conf_general.Instance);
                        Consulta_No_Conf_general.Instance.Dock = DockStyle.Fill;
                        Consulta_No_Conf_general.Instance.BringToFront();


                    }
                    else
                    {
                        Consulta_No_Conf_general.Instance.BringToFront();
                        Ventana = "Pantalla: Consulta de No Conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;
                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central7.Controls.Contains(Consulta_No_Conf_general.Instance))
                        {
                            Ventana = "Pantalla: Consulta de No Conformidad";

                            Central7.Controls.Add(Consulta_No_Conf_general.Instance);
                            Consulta_No_Conf_general.Instance.Dock = DockStyle.Fill;
                            Consulta_No_Conf_general.Instance.BringToFront();


                        }
                        else
                        {
                            Consulta_No_Conf_general.Instance.BringToFront();
                            Ventana = "Pantalla: Consulta de No Conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void seguimientoToolStripMenuItem3_Click_1(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central7.Controls.Contains(Seguimiento_No_confor.Instance))
                    {
                        Ventana = "Pantalla: Seguimiento de No conformidad";
                        Seguimiento_No_confor.Instance.Departamento = Departamento;
                        Seguimiento_No_confor.Instance.Nombre = Nombre;
                        Seguimiento_No_confor.Instance.Apellido = Apellido;
                        Seguimiento_No_confor.Instance.NoConfor = NumeroConformidad;
                        Seguimiento_No_confor.Instance.Ventana1 = Ventana;
                        Central7.Controls.Add(Seguimiento_No_confor.Instance);
                        Seguimiento_No_confor.Instance.Dock = DockStyle.Fill;
                        Seguimiento_No_confor.Instance.BringToFront();


                    }
                    else
                    {
                        Seguimiento_No_confor.Instance.BringToFront();
                        Ventana = "Pantalla: Seguimiento de No conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //    MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);   
                    }
                    oNetDrive = null;
                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central7.Controls.Contains(Seguimiento_No_confor.Instance))
                        {
                            Ventana = "Pantalla: Seguimiento de No conformidad";
                            Seguimiento_No_confor.Instance.Departamento = Departamento;
                            Seguimiento_No_confor.Instance.Nombre = Nombre;
                            Seguimiento_No_confor.Instance.Apellido = Apellido;
                            Seguimiento_No_confor.Instance.NoConfor = NumeroConformidad;
                            Seguimiento_No_confor.Instance.Ventana1 = Ventana;
                            Central7.Controls.Add(Seguimiento_No_confor.Instance);
                            Seguimiento_No_confor.Instance.Dock = DockStyle.Fill;
                            Seguimiento_No_confor.Instance.BringToFront();


                        }
                        else
                        {
                            Seguimiento_No_confor.Instance.BringToFront();
                            Ventana = "Pantalla: Seguimiento de No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }



        }

        private void toolStripMenuItem14_Click(object sender, EventArgs e)
        {
            if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "AUXILIAR DE  PRODUCCION") { }
            else { MessageBox.Show("No tiene acceso a este modulo"); }
        }

        private void toolStripMenuItem76_Click_1(object sender, EventArgs e)
        {
            if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "Recepcion") { }
            else { MessageBox.Show("No tiene acceso a este modulo"); }
        }

        private void toolStripMenuItem137_Click(object sender, EventArgs e)
        {
            if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "Administrador TI") { }
            else { MessageBox.Show("No tiene acceso a este modulo"); }
        }

        private void toolStripMenuItem119_Click(object sender, EventArgs e)
        {
            if (Departamento == "Infraestructura" || Departamento == "SuperAdmin" || Departamento == "Direccion") { }
            else { MessageBox.Show("No tiene acceso a este modulo"); }
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();


        }

        private void Central8_Paint(object sender, PaintEventArgs e)
        {

        }

        private void toolStripMenuItem128_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;

                if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "Recepcion")
                {

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central8.Controls.Contains(Genera_No_Confor.Instance))
                        {
                            Ventana = "Pantalla: Levantamiento de No conformidad";
                            Genera_No_Confor.Instance.Nombre = Nombre;
                            Genera_No_Confor.Instance.Ventana1 = Ventana;

                            Central8.Controls.Add(Genera_No_Confor.Instance);
                            Genera_No_Confor.Instance.Dock = DockStyle.Fill;
                            Genera_No_Confor.Instance.BringToFront();


                        }
                        else
                        {
                            Genera_No_Confor.Instance.BringToFront();
                            Ventana = "Pantalla: Levantamiento de No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else { MessageBox.Show("No tiene acceso a este modulo"); }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);   
                    }
                    oNetDrive = null;


                    if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "Recepcion")
                    {

                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                        {

                            if (!Central8.Controls.Contains(Genera_No_Confor.Instance))
                            {
                                Ventana = "Pantalla: Levantamiento de No conformidad";
                                Genera_No_Confor.Instance.Nombre = Nombre;
                                Genera_No_Confor.Instance.Ventana1 = Ventana;

                                Central8.Controls.Add(Genera_No_Confor.Instance);
                                Genera_No_Confor.Instance.Dock = DockStyle.Fill;
                                Genera_No_Confor.Instance.BringToFront();


                            }
                            else
                            {
                                Genera_No_Confor.Instance.BringToFront();
                                Ventana = "Pantalla: Levantamiento de No conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                    }
                    else { MessageBox.Show("No tiene acceso a este modulo"); }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }

        }

        private void toolStripMenuItem129_Click(object sender, EventArgs e)
        {

            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;

                if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "Recepcion")
                {


                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central8.Controls.Contains(Consulta_No_Conf.Instance))
                        {
                            Ventana = "Pantalla: Consulta de No Conformidad";
                            Consulta_No_Conf.Instance.Departamento = Departamento;
                            Central8.Controls.Add(Consulta_No_Conf.Instance);
                            Consulta_No_Conf.Instance.Dock = DockStyle.Fill;
                            Consulta_No_Conf.Instance.BringToFront();


                        }
                        else
                        {
                            Consulta_No_Conf.Instance.BringToFront();
                            Ventana = "Pantalla: Consulta de No Conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else { MessageBox.Show("No tiene acceso a este modulo"); }

            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;


                    if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "Recepcion")
                    {


                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5") if (Permiso == "1" || Permiso == "2")
                            {

                                if (!Central8.Controls.Contains(Consulta_No_Conf.Instance))
                                {
                                    Ventana = "Pantalla: Consulta de No Conformidad";
                                    Consulta_No_Conf.Instance.Departamento = Departamento;
                                    Central8.Controls.Add(Consulta_No_Conf.Instance);
                                    Consulta_No_Conf.Instance.Dock = DockStyle.Fill;
                                    Consulta_No_Conf.Instance.BringToFront();


                                }
                                else
                                {
                                    Consulta_No_Conf.Instance.BringToFront();
                                    Ventana = "Pantalla: Consulta de No Conformidad";
                                }
                            }
                            else
                            {
                                MessageBox.Show("No tiene acceso a este modulo");
                            }
                    }
                    else { MessageBox.Show("No tiene acceso a este modulo"); }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void toolStripMenuItem131_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "Recepcion")
                {

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central8.Controls.Contains(Seguimiento_No_confor.Instance))
                        {
                            Ventana = "Pantalla: Seguimiento de No conformidad";
                            Seguimiento_No_confor.Instance.Departamento = Departamento;
                            Seguimiento_No_confor.Instance.Nombre = Nombre;
                            Seguimiento_No_confor.Instance.Apellido = Apellido;
                            Seguimiento_No_confor.Instance.NoConfor = NumeroConformidad;
                            Seguimiento_No_confor.Instance.Ventana1 = Ventana;
                            Central8.Controls.Add(Seguimiento_No_confor.Instance);
                            Seguimiento_No_confor.Instance.Dock = DockStyle.Fill;
                            Seguimiento_No_confor.Instance.BringToFront();


                        }
                        else
                        {
                            Seguimiento_No_confor.Instance.BringToFront();
                            Ventana = "Pantalla: Seguimiento de No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else { MessageBox.Show("No tiene acceso a este modulo"); }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //       MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); 
                    }
                    oNetDrive = null;

                    if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "Recepcion")
                    {

                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                        {

                            if (!Central8.Controls.Contains(Seguimiento_No_confor.Instance))
                            {
                                Ventana = "Pantalla: Seguimiento de No conformidad";
                                Seguimiento_No_confor.Instance.Departamento = Departamento;
                                Seguimiento_No_confor.Instance.Nombre = Nombre;
                                Seguimiento_No_confor.Instance.Apellido = Apellido;
                                Seguimiento_No_confor.Instance.NoConfor = NumeroConformidad;
                                Seguimiento_No_confor.Instance.Ventana1 = Ventana;
                                Central8.Controls.Add(Seguimiento_No_confor.Instance);
                                Seguimiento_No_confor.Instance.Dock = DockStyle.Fill;
                                Seguimiento_No_confor.Instance.BringToFront();


                            }
                            else
                            {
                                Seguimiento_No_confor.Instance.BringToFront();
                                Ventana = "Pantalla: Seguimiento de No conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                    }
                    else { MessageBox.Show("No tiene acceso a este modulo"); }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void toolStripMenuItem133_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "Recepcion")
                {

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central8.Controls.Contains(Genera_AM.Instance))
                        {
                            Ventana = "Pantalla: Genera Oportunidad de Mejora";
                            Genera_AM.Instance.Departamento = Departamento;
                            Genera_AM.Instance.Nombre = Nombre;
                            Genera_AM.Instance.Ventana1 = Ventana;
                            Genera_AM.Instance.NumeroAC = NumeroOport;
                            Central8.Controls.Add(Genera_AM.Instance);
                            Genera_AM.Instance.Dock = DockStyle.Fill;
                            Genera_AM.Instance.BringToFront();


                        }
                        else
                        {
                            Genera_AM.Instance.BringToFront();
                            Ventana = "Pantalla: Genera Oportunidad de Mejora";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else { MessageBox.Show("No tiene acceso a este modulo"); }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;



                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); 
                    }
                    oNetDrive = null;
                    if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "Recepcion")
                    {

                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                        {

                            if (!Central8.Controls.Contains(Genera_AM.Instance))
                            {
                                Ventana = "Pantalla: Genera Oportunidad de Mejora";
                                Genera_AM.Instance.Departamento = Departamento;
                                Genera_AM.Instance.Nombre = Nombre;
                                Genera_AM.Instance.Ventana1 = Ventana;
                                Genera_AM.Instance.NumeroAC = NumeroOport;
                                Central8.Controls.Add(Genera_AM.Instance);
                                Genera_AM.Instance.Dock = DockStyle.Fill;
                                Genera_AM.Instance.BringToFront();


                            }
                            else
                            {
                                Genera_AM.Instance.BringToFront();
                                Ventana = "Pantalla: Genera Oportunidad de Mejora";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                    }
                    else { MessageBox.Show("No tiene acceso a este modulo"); }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void toolStripMenuItem134_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "Recepcion")
                {

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central8.Controls.Contains(Seguimiento_AM.Instance))
                        {
                            Ventana = "Pantalla: Detalle No conformidad";
                            Seguimiento_AM.Instance.Nombre = Nombre;
                            Seguimiento_AM.Instance.Apellido = Apellido;
                            Seguimiento_AM.Instance.Departamento = Departamento;
                            Seguimiento_AM.Instance.Ventana1 = Ventana;
                            Central8.Controls.Add(Seguimiento_AM.Instance);
                            Seguimiento_AM.Instance.Dock = DockStyle.Fill;
                            Seguimiento_AM.Instance.BringToFront();


                        }
                        else
                        {
                            Seguimiento_AM.Instance.BringToFront();
                            Ventana = "Pantalla: Detalle No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else { MessageBox.Show("No tiene acceso a este modulo"); }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {

                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //     MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;
                    if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "Recepcion")
                    {

                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                        {

                            if (!Central8.Controls.Contains(Seguimiento_AM.Instance))
                            {
                                Ventana = "Pantalla: Detalle No conformidad";
                                Seguimiento_AM.Instance.Nombre = Nombre;
                                Seguimiento_AM.Instance.Apellido = Apellido;
                                Seguimiento_AM.Instance.Departamento = Departamento;
                                Seguimiento_AM.Instance.Ventana1 = Ventana;
                                Central8.Controls.Add(Seguimiento_AM.Instance);
                                Seguimiento_AM.Instance.Dock = DockStyle.Fill;
                                Seguimiento_AM.Instance.BringToFront();


                            }
                            else
                            {
                                Seguimiento_AM.Instance.BringToFront();
                                Ventana = "Pantalla: Detalle No conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                    }
                    else { MessageBox.Show("No tiene acceso a este modulo"); }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void toolStripMenuItem135_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "Recepcion")
                {

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central8.Controls.Contains(Consulta_OportunidadesMejora_General.Instance))
                        {
                            Ventana = "Pantalla: Detalle No conformidad";

                            Central8.Controls.Add(Consulta_OportunidadesMejora_General.Instance);
                            Consulta_OportunidadesMejora_General.Instance.Dock = DockStyle.Fill;
                            Consulta_OportunidadesMejora_General.Instance.BringToFront();


                        }
                        else
                        {
                            Consulta_OportunidadesMejora_General.Instance.BringToFront();
                            Ventana = "Pantalla: Detalle No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else { MessageBox.Show("No tiene acceso a este modulo"); }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);
                    }
                    oNetDrive = null;
                    if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "Recepcion")
                    {

                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                        {

                            if (!Central8.Controls.Contains(Consulta_OportunidadesMejora_General.Instance))
                            {
                                Ventana = "Pantalla: Detalle No conformidad";

                                Central8.Controls.Add(Consulta_OportunidadesMejora_General.Instance);
                                Consulta_OportunidadesMejora_General.Instance.Dock = DockStyle.Fill;
                                Consulta_OportunidadesMejora_General.Instance.BringToFront();


                            }
                            else
                            {
                                Consulta_OportunidadesMejora_General.Instance.BringToFront();
                                Ventana = "Pantalla: Detalle No conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                    }
                    else { MessageBox.Show("No tiene acceso a este modulo"); }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void toolStripMenuItem136_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }

            if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "Recepcion")
            {


                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central8.Controls.Contains(Consulta_Oportunidades_Mejora.Instance))
                    {
                        Ventana = "Pantalla: Detalle No conformidad";


                        Consulta_Oportunidades_Mejora.Instance.Nombre = Nombre;
                        Consulta_Oportunidades_Mejora.Instance.Departamento = Departamento;
                        Consulta_Oportunidades_Mejora.Instance.Ventana1 = Ventana;

                        Central8.Controls.Add(Consulta_Oportunidades_Mejora.Instance);
                        Consulta_Oportunidades_Mejora.Instance.Dock = DockStyle.Fill;
                        Consulta_Oportunidades_Mejora.Instance.BringToFront();



                    }
                    else
                    {
                        Consulta_OportunidadesMejora_General.Instance.BringToFront();
                        Ventana = "Pantalla: Detalle No conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }


            conexionvpn = 2;

            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;


                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);    
                    }
                    oNetDrive = null;

                    if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "Recepcion")
                    {


                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                        {

                            if (!Central8.Controls.Contains(Consulta_Oportunidades_Mejora.Instance))
                            {
                                Ventana = "Pantalla: Detalle No conformidad";


                                Consulta_Oportunidades_Mejora.Instance.Nombre = Nombre;
                                Consulta_Oportunidades_Mejora.Instance.Departamento = Departamento;
                                Consulta_Oportunidades_Mejora.Instance.Ventana1 = Ventana;

                                Central8.Controls.Add(Consulta_Oportunidades_Mejora.Instance);
                                Consulta_Oportunidades_Mejora.Instance.Dock = DockStyle.Fill;
                                Consulta_Oportunidades_Mejora.Instance.BringToFront();



                            }
                            else
                            {
                                Consulta_OportunidadesMejora_General.Instance.BringToFront();
                                Ventana = "Pantalla: Detalle No conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                    }
                    else { MessageBox.Show("No tiene acceso a este modulo"); }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void toolStripMenuItem139_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "Administrador TI")
                {

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central8.Controls.Contains(Genera_No_Confor.Instance))
                        {
                            Ventana = "Pantalla: Levantamiento de No conformidad";
                            Genera_No_Confor.Instance.Nombre = Nombre;
                            Genera_No_Confor.Instance.Ventana1 = Ventana;

                            Central8.Controls.Add(Genera_No_Confor.Instance);
                            Genera_No_Confor.Instance.Dock = DockStyle.Fill;
                            Genera_No_Confor.Instance.BringToFront();


                        }
                        else
                        {
                            Genera_No_Confor.Instance.BringToFront();
                            Ventana = "Pantalla: Levantamiento de No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else { MessageBox.Show("No tiene acceso a este modulo"); }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {//  MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); 
                    }
                    oNetDrive = null;
                    if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "Administrador TI")
                    {

                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                        {

                            if (!Central8.Controls.Contains(Genera_No_Confor.Instance))
                            {
                                Ventana = "Pantalla: Levantamiento de No conformidad";
                                Genera_No_Confor.Instance.Nombre = Nombre;
                                Genera_No_Confor.Instance.Ventana1 = Ventana;

                                Central8.Controls.Add(Genera_No_Confor.Instance);
                                Genera_No_Confor.Instance.Dock = DockStyle.Fill;
                                Genera_No_Confor.Instance.BringToFront();


                            }
                            else
                            {
                                Genera_No_Confor.Instance.BringToFront();
                                Ventana = "Pantalla: Levantamiento de No conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                    }
                    else { MessageBox.Show("No tiene acceso a este modulo"); }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }
        
    

     
        }

        private void toolStripMenuItem140_Click(object sender, EventArgs e)
        {

            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else { conexionvpn = 2;

                if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "Administrador TI")
                {


                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5") if (Permiso == "1" || Permiso == "2")
                        {

                            if (!Central8.Controls.Contains(Consulta_No_Conf.Instance))
                            {
                                Ventana = "Pantalla: Consulta de No Conformidad";
                                Consulta_No_Conf.Instance.Departamento = Departamento;
                                Central8.Controls.Add(Consulta_No_Conf.Instance);
                                Consulta_No_Conf.Instance.Dock = DockStyle.Fill;
                                Consulta_No_Conf.Instance.BringToFront();


                            }
                            else
                            {
                                Consulta_No_Conf.Instance.BringToFront();
                                Ventana = "Pantalla: Consulta de No Conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                }
                else { MessageBox.Show("No tiene acceso a este modulo"); }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                             oNetDrive1.Force = true;  oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                           // MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);    
                    }
                    oNetDrive = null;

                    if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "Administrador TI")
                    {


                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5") if (Permiso == "1" || Permiso == "2")
                            {

                                if (!Central8.Controls.Contains(Consulta_No_Conf.Instance))
                                {
                                    Ventana = "Pantalla: Consulta de No Conformidad";
                                    Consulta_No_Conf.Instance.Departamento = Departamento;
                                    Central8.Controls.Add(Consulta_No_Conf.Instance);
                                    Consulta_No_Conf.Instance.Dock = DockStyle.Fill;
                                    Consulta_No_Conf.Instance.BringToFront();


                                }
                                else
                                {
                                    Consulta_No_Conf.Instance.BringToFront();
                                    Ventana = "Pantalla: Consulta de No Conformidad";
                                }
                            }
                            else
                            {
                                MessageBox.Show("No tiene acceso a este modulo");
                            }
                    }
                    else { MessageBox.Show("No tiene acceso a este modulo"); }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }

        }

        private void toolStripMenuItem130_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else { conexionvpn = 2;
                if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "Recepcion")
                {

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5") 
                        {

                            if (!Central8.Controls.Contains(Consulta_No_Conf_general.Instance))
                            {
                                Ventana = "Pantalla: Consulta de No Conformidad";

                                Central8.Controls.Add(Consulta_No_Conf_general.Instance);
                                Consulta_No_Conf_general.Instance.Dock = DockStyle.Fill;
                                Consulta_No_Conf_general.Instance.BringToFront();


                            }
                            else
                            {
                                Consulta_No_Conf_general.Instance.BringToFront();
                                Ventana = "Pantalla: Consulta de No Conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                }
                else { MessageBox.Show("No tiene acceso a este modulo"); }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                             oNetDrive1.Force = true;  oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                          //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;
                    if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "Recepcion")
                    {

                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5") if (Permiso == "1" || Permiso == "2")
                            {

                                if (!Central8.Controls.Contains(Consulta_No_Conf_general.Instance))
                                {
                                    Ventana = "Pantalla: Consulta de No Conformidad";

                                    Central8.Controls.Add(Consulta_No_Conf_general.Instance);
                                    Consulta_No_Conf_general.Instance.Dock = DockStyle.Fill;
                                    Consulta_No_Conf_general.Instance.BringToFront();


                                }
                                else
                                {
                                    Consulta_No_Conf_general.Instance.BringToFront();
                                    Ventana = "Pantalla: Consulta de No Conformidad";
                                }
                            }
                            else
                            {
                                MessageBox.Show("No tiene acceso a este modulo");
                            }
                    }
                    else { MessageBox.Show("No tiene acceso a este modulo"); }

                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }

     
        }

        private void toolStripMenuItem141_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else { conexionvpn = 2;
                if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "Administrador TI")
                {

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5") 
                        {

                            if (!Central8.Controls.Contains(Consulta_No_Conf_general.Instance))
                            {
                                Ventana = "Pantalla: Consulta de No Conformidad";

                                Central8.Controls.Add(Consulta_No_Conf_general.Instance);
                                Consulta_No_Conf_general.Instance.Dock = DockStyle.Fill;
                                Consulta_No_Conf_general.Instance.BringToFront();


                            }
                            else
                            {
                                Consulta_No_Conf_general.Instance.BringToFront();
                                Ventana = "Pantalla: Consulta de No Conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                }
                else { MessageBox.Show("No tiene acceso a este modulo"); }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                             oNetDrive1.Force = true;  oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                         //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); 
                    }
                    oNetDrive = null;
                    if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "AUXILIAR DE  PRODUCCION")
                    {

                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                            {

                                if (!Central8.Controls.Contains(Consulta_No_Conf_general.Instance))
                                {
                                    Ventana = "Pantalla: Consulta de No Conformidad";

                                    Central8.Controls.Add(Consulta_No_Conf_general.Instance);
                                    Consulta_No_Conf_general.Instance.Dock = DockStyle.Fill;
                                    Consulta_No_Conf_general.Instance.BringToFront();


                                }
                                else
                                {
                                    Consulta_No_Conf_general.Instance.BringToFront();
                                    Ventana = "Pantalla: Consulta de No Conformidad";
                                }
                            }
                            else
                            {
                                MessageBox.Show("No tiene acceso a este modulo");
                            }
                    }
                    else { MessageBox.Show("No tiene acceso a este modulo"); }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }

    
        }

        private void toolStripMenuItem142_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else { conexionvpn = 2;
                if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "Administrador TI")
                {

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central8.Controls.Contains(Seguimiento_No_confor.Instance))
                        {
                            Ventana = "Pantalla: Seguimiento de No conformidad";
                            Seguimiento_No_confor.Instance.Departamento = Departamento;
                            Seguimiento_No_confor.Instance.Nombre = Nombre;
                            Seguimiento_No_confor.Instance.Apellido = Apellido;
                            Seguimiento_No_confor.Instance.NoConfor = NumeroConformidad;
                            Seguimiento_No_confor.Instance.Ventana1 = Ventana;
                            Central8.Controls.Add(Seguimiento_No_confor.Instance);
                            Seguimiento_No_confor.Instance.Dock = DockStyle.Fill;
                            Seguimiento_No_confor.Instance.BringToFront();


                        }
                        else
                        {
                            Seguimiento_No_confor.Instance.BringToFront();
                            Ventana = "Pantalla: Seguimiento de No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else { MessageBox.Show("No tiene acceso a este modulo"); }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                             oNetDrive1.Force = true;  oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                         //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;
                    if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "Administrador TI")
                    {

                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                        {

                            if (!Central8.Controls.Contains(Seguimiento_No_confor.Instance))
                            {
                                Ventana = "Pantalla: Seguimiento de No conformidad";
                                Seguimiento_No_confor.Instance.Departamento = Departamento;
                                Seguimiento_No_confor.Instance.Nombre = Nombre;
                                Seguimiento_No_confor.Instance.Apellido = Apellido;
                                Seguimiento_No_confor.Instance.NoConfor = NumeroConformidad;
                                Seguimiento_No_confor.Instance.Ventana1 = Ventana;
                                Central8.Controls.Add(Seguimiento_No_confor.Instance);
                                Seguimiento_No_confor.Instance.Dock = DockStyle.Fill;
                                Seguimiento_No_confor.Instance.BringToFront();


                            }
                            else
                            {
                                Seguimiento_No_confor.Instance.BringToFront();
                                Ventana = "Pantalla: Seguimiento de No conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                    }
                    else { MessageBox.Show("No tiene acceso a este modulo"); }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }

         
        }

        private void toolStripMenuItem144_Click(object sender, EventArgs e)
        {
            validacion();
            if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "Administrador TI")
            {

                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central8.Controls.Contains(Genera_AM.Instance))
                    {
                        Ventana = "Pantalla: Genera Oportunidad de Mejora";
                        Genera_AM.Instance.Departamento = Departamento;
                        Genera_AM.Instance.Nombre = Nombre;
                        Genera_AM.Instance.Ventana1 = Ventana;
                        Genera_AM.Instance.NumeroAC = NumeroOport;
                        Central8.Controls.Add(Genera_AM.Instance);
                        Genera_AM.Instance.Dock = DockStyle.Fill;
                        Genera_AM.Instance.BringToFront();


                    }
                    else
                    {
                        Genera_AM.Instance.BringToFront();
                        Ventana = "Pantalla: Genera Oportunidad de Mejora";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            else { MessageBox.Show("No tiene acceso a este modulo"); }
        }

        private void toolStripMenuItem145_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else { conexionvpn = 2;
                if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "Administrador TI")
                {

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central8.Controls.Contains(Seguimiento_AM.Instance))
                        {
                            Ventana = "Pantalla: Detalle No conformidad";
                            Seguimiento_AM.Instance.Nombre = Nombre;
                            Seguimiento_AM.Instance.Apellido = Apellido;
                            Seguimiento_AM.Instance.Departamento = Departamento;
                            Seguimiento_AM.Instance.Ventana1 = Ventana;
                            Central8.Controls.Add(Seguimiento_AM.Instance);
                            Seguimiento_AM.Instance.Dock = DockStyle.Fill;
                            Seguimiento_AM.Instance.BringToFront();


                        }
                        else
                        {
                            Seguimiento_AM.Instance.BringToFront();
                            Ventana = "Pantalla: Detalle No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else { MessageBox.Show("No tiene acceso a este modulo"); }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;
                    if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "Administrador TI")
                    {

                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                        {

                            if (!Central8.Controls.Contains(Seguimiento_AM.Instance))
                            {
                                Ventana = "Pantalla: Detalle No conformidad";
                                Seguimiento_AM.Instance.Nombre = Nombre;
                                Seguimiento_AM.Instance.Apellido = Apellido;
                                Seguimiento_AM.Instance.Departamento = Departamento;
                                Seguimiento_AM.Instance.Ventana1 = Ventana;
                                Central8.Controls.Add(Seguimiento_AM.Instance);
                                Seguimiento_AM.Instance.Dock = DockStyle.Fill;
                                Seguimiento_AM.Instance.BringToFront();


                            }
                            else
                            {
                                Seguimiento_AM.Instance.BringToFront();
                                Ventana = "Pantalla: Detalle No conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                    }
                    else { MessageBox.Show("No tiene acceso a este modulo"); }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }

         
        }

        private void toolStripMenuItem160_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else { conexionvpn = 2;
                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!central10.Controls.Contains(Genera_No_Confor.Instance))
                    {
                        Ventana = "Pantalla: Levantamiento de No conformidad";
                        Genera_No_Confor.Instance.Nombre = Nombre;
                        Genera_No_Confor.Instance.Ventana1 = Ventana;

                        central10.Controls.Add(Genera_No_Confor.Instance);
                        Genera_No_Confor.Instance.Dock = DockStyle.Fill;
                        Genera_No_Confor.Instance.BringToFront();


                    }
                    else
                    {
                        Genera_No_Confor.Instance.BringToFront();
                        Ventana = "Pantalla: Levantamiento de No conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }

            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {

                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                             oNetDrive1.Force = true;  oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                       //     MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;


                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!central10.Controls.Contains(Genera_No_Confor.Instance))
                        {
                            Ventana = "Pantalla: Levantamiento de No conformidad";
                            Genera_No_Confor.Instance.Nombre = Nombre;
                            Genera_No_Confor.Instance.Ventana1 = Ventana;

                            central10.Controls.Add(Genera_No_Confor.Instance);
                            Genera_No_Confor.Instance.Dock = DockStyle.Fill;
                            Genera_No_Confor.Instance.BringToFront();


                        }
                        else
                        {
                            Genera_No_Confor.Instance.BringToFront();
                            Ventana = "Pantalla: Levantamiento de No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }

                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }

    

        }

        private void toolStripMenuItem161_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else { conexionvpn = 2;
                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5") if (Permiso == "1" || Permiso == "2")
                    {

                        if (!central10.Controls.Contains(Consulta_No_Conf.Instance))
                        {
                            Ventana = "Pantalla: Consulta de No Conformidad";
                            Consulta_No_Conf.Instance.Departamento = Departamento;
                            central10.Controls.Add(Consulta_No_Conf.Instance);
                            Consulta_No_Conf.Instance.Dock = DockStyle.Fill;
                            Consulta_No_Conf.Instance.BringToFront();


                        }
                        else
                        {
                            Consulta_No_Conf.Instance.BringToFront();
                            Ventana = "Pantalla: Consulta de No Conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                             oNetDrive1.Force = true;  oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                            //MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;


                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);    
                    }
                    oNetDrive = null;
                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5") if (Permiso == "1" || Permiso == "2")
                        {

                            if (!central10.Controls.Contains(Consulta_No_Conf.Instance))
                            {
                                Ventana = "Pantalla: Consulta de No Conformidad";
                                Consulta_No_Conf.Instance.Departamento = Departamento;
                                central10.Controls.Add(Consulta_No_Conf.Instance);
                                Consulta_No_Conf.Instance.Dock = DockStyle.Fill;
                                Consulta_No_Conf.Instance.BringToFront();


                            }
                            else
                            {
                                Consulta_No_Conf.Instance.BringToFront();
                                Ventana = "Pantalla: Consulta de No Conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }

        
        }

        private void toolStripMenuItem162_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else { conexionvpn = 2;

                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5") if (Permiso == "1" || Permiso == "2")
                    {

                        if (!central10.Controls.Contains(Consulta_No_Conf_general.Instance))
                        {
                            Ventana = "Pantalla: Consulta de No Conformidad";

                            central10.Controls.Add(Consulta_No_Conf_general.Instance);
                            Consulta_No_Conf_general.Instance.Dock = DockStyle.Fill;
                            Consulta_No_Conf_general.Instance.BringToFront();


                        }
                        else
                        {
                            Consulta_No_Conf_general.Instance.BringToFront();
                            Ventana = "Pantalla: Consulta de No Conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                             oNetDrive1.Force = true;  oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                        //    MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5") if (Permiso == "1" || Permiso == "2")
                        {

                            if (!central10.Controls.Contains(Consulta_No_Conf_general.Instance))
                            {
                                Ventana = "Pantalla: Consulta de No Conformidad";

                                central10.Controls.Add(Consulta_No_Conf_general.Instance);
                                Consulta_No_Conf_general.Instance.Dock = DockStyle.Fill;
                                Consulta_No_Conf_general.Instance.BringToFront();


                            }
                            else
                            {
                                Consulta_No_Conf_general.Instance.BringToFront();
                                Ventana = "Pantalla: Consulta de No Conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }

                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }

            }

        private void toolStripMenuItem163_Click(object sender, EventArgs e)
        {
            if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
            {

                if (!central10.Controls.Contains(Seguimiento_No_confor.Instance))
                {
                    Ventana = "Pantalla: Seguimiento de No conformidad";
                    Seguimiento_No_confor.Instance.Departamento = Departamento;
                    Seguimiento_No_confor.Instance.Nombre = Nombre;
                    Seguimiento_No_confor.Instance.Apellido = Apellido;
                    Seguimiento_No_confor.Instance.NoConfor = NumeroConformidad;
                    Seguimiento_No_confor.Instance.Ventana1 = Ventana;
                    central10.Controls.Add(Seguimiento_No_confor.Instance);
                    Seguimiento_No_confor.Instance.Dock = DockStyle.Fill;
                    Seguimiento_No_confor.Instance.BringToFront();


                }
                else
                {
                    Seguimiento_No_confor.Instance.BringToFront();
                    Ventana = "Pantalla: Seguimiento de No conformidad";
                }
            }
        }

        private void toolStripMenuItem163_Click_1(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else { conexionvpn = 2;

                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!central10.Controls.Contains(Seguimiento_No_confor.Instance))
                    {
                        Ventana = "Pantalla: Seguimiento de No conformidad";
                        Seguimiento_No_confor.Instance.Departamento = Departamento;
                        Seguimiento_No_confor.Instance.Nombre = Nombre;
                        Seguimiento_No_confor.Instance.Apellido = Apellido;
                        Seguimiento_No_confor.Instance.NoConfor = NumeroConformidad;
                        Seguimiento_No_confor.Instance.Ventana1 = Ventana;
                        central10.Controls.Add(Seguimiento_No_confor.Instance);
                        Seguimiento_No_confor.Instance.Dock = DockStyle.Fill;
                        Seguimiento_No_confor.Instance.BringToFront();


                    }
                    else
                    {
                        Seguimiento_No_confor.Instance.BringToFront();
                        Ventana = "Pantalla: Seguimiento de No conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                             oNetDrive1.Force = true;  oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                       //     MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);    
                    }
                    oNetDrive = null;

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!central10.Controls.Contains(Seguimiento_No_confor.Instance))
                        {
                            Ventana = "Pantalla: Seguimiento de No conformidad";
                            Seguimiento_No_confor.Instance.Departamento = Departamento;
                            Seguimiento_No_confor.Instance.Nombre = Nombre;
                            Seguimiento_No_confor.Instance.Apellido = Apellido;
                            Seguimiento_No_confor.Instance.NoConfor = NumeroConformidad;
                            Seguimiento_No_confor.Instance.Ventana1 = Ventana;
                            central10.Controls.Add(Seguimiento_No_confor.Instance);
                            Seguimiento_No_confor.Instance.Dock = DockStyle.Fill;
                            Seguimiento_No_confor.Instance.BringToFront();


                        }
                        else
                        {
                            Seguimiento_No_confor.Instance.BringToFront();
                            Ventana = "Pantalla: Seguimiento de No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }

        }

        private void toolStripMenuItem176_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else { conexionvpn = 2;

                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!central10.Controls.Contains(Genera_AM.Instance))
                    {
                        Ventana = "Pantalla: Genera Oportunidad de Mejora";
                        Genera_AM.Instance.Departamento = Departamento;
                        Genera_AM.Instance.Nombre = Nombre;
                        Genera_AM.Instance.Ventana1 = Ventana;
                        Genera_AM.Instance.NumeroAC = NumeroOport;
                        central10.Controls.Add(Genera_AM.Instance);
                        Genera_AM.Instance.Dock = DockStyle.Fill;
                        Genera_AM.Instance.BringToFront();


                    }
                    else
                    {
                        Genera_AM.Instance.BringToFront();
                        Ventana = "Pantalla: Genera Oportunidad de Mejora";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {

                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                             oNetDrive1.Force = true;  oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                          //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!central10.Controls.Contains(Genera_AM.Instance))
                        {
                            Ventana = "Pantalla: Genera Oportunidad de Mejora";
                            Genera_AM.Instance.Departamento = Departamento;
                            Genera_AM.Instance.Nombre = Nombre;
                            Genera_AM.Instance.Ventana1 = Ventana;
                            Genera_AM.Instance.NumeroAC = NumeroOport;
                            central10.Controls.Add(Genera_AM.Instance);
                            Genera_AM.Instance.Dock = DockStyle.Fill;
                            Genera_AM.Instance.BringToFront();


                        }
                        else
                        {
                            Genera_AM.Instance.BringToFront();
                            Ventana = "Pantalla: Genera Oportunidad de Mejora";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }

                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }

        }

        private void toolStripMenuItem177_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else { conexionvpn = 2;
                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!central10.Controls.Contains(Seguimiento_AM.Instance))
                    {
                        Ventana = "Pantalla: Detalle No conformidad";
                        Seguimiento_AM.Instance.Nombre = Nombre;
                        Seguimiento_AM.Instance.Apellido = Apellido;
                        Seguimiento_AM.Instance.Departamento = Departamento;
                        Seguimiento_AM.Instance.Ventana1 = Ventana;
                        central10.Controls.Add(Seguimiento_AM.Instance);
                        Seguimiento_AM.Instance.Dock = DockStyle.Fill;
                        Seguimiento_AM.Instance.BringToFront();


                    }
                    else
                    {
                        Seguimiento_AM.Instance.BringToFront();
                        Ventana = "Pantalla: Detalle No conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                             oNetDrive1.Force = true;  oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                        //    MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!central10.Controls.Contains(Seguimiento_AM.Instance))
                        {
                            Ventana = "Pantalla: Detalle No conformidad";
                            Seguimiento_AM.Instance.Nombre = Nombre;
                            Seguimiento_AM.Instance.Apellido = Apellido;
                            Seguimiento_AM.Instance.Departamento = Departamento;
                            Seguimiento_AM.Instance.Ventana1 = Ventana;
                            central10.Controls.Add(Seguimiento_AM.Instance);
                            Seguimiento_AM.Instance.Dock = DockStyle.Fill;
                            Seguimiento_AM.Instance.BringToFront();


                        }
                        else
                        {
                            Seguimiento_AM.Instance.BringToFront();
                            Ventana = "Pantalla: Detalle No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }

 
        }

        private void toolStripMenuItem178_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else { conexionvpn = 2;

                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!central10.Controls.Contains(Consulta_OportunidadesMejora_General.Instance))
                    {
                        Ventana = "Pantalla: Detalle No conformidad";

                        central10.Controls.Add(Consulta_OportunidadesMejora_General.Instance);
                        Consulta_OportunidadesMejora_General.Instance.Dock = DockStyle.Fill;
                        Consulta_OportunidadesMejora_General.Instance.BringToFront();


                    }
                    else
                    {
                        Consulta_OportunidadesMejora_General.Instance.BringToFront();
                        Ventana = "Pantalla: Detalle No conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                             oNetDrive1.Force = true;  oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                           // MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!central10.Controls.Contains(Consulta_OportunidadesMejora_General.Instance))
                        {
                            Ventana = "Pantalla: Detalle No conformidad";

                            central10.Controls.Add(Consulta_OportunidadesMejora_General.Instance);
                            Consulta_OportunidadesMejora_General.Instance.Dock = DockStyle.Fill;
                            Consulta_OportunidadesMejora_General.Instance.BringToFront();


                        }
                        else
                        {
                            Consulta_OportunidadesMejora_General.Instance.BringToFront();
                            Ventana = "Pantalla: Detalle No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }

        }

        private void toolStripMenuItem179_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else { conexionvpn = 2;
                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!central10.Controls.Contains(Consulta_Oportunidades_Mejora.Instance))
                    {
                        Ventana = "Pantalla: Detalle No conformidad";


                        Consulta_Oportunidades_Mejora.Instance.Nombre = Nombre;
                        Consulta_Oportunidades_Mejora.Instance.Departamento = Departamento;
                        Consulta_Oportunidades_Mejora.Instance.Ventana1 = Ventana;

                        central10.Controls.Add(Consulta_Oportunidades_Mejora.Instance);
                        Consulta_Oportunidades_Mejora.Instance.Dock = DockStyle.Fill;
                        Consulta_Oportunidades_Mejora.Instance.BringToFront();



                    }
                    else
                    {
                        Consulta_OportunidadesMejora_General.Instance.BringToFront();
                        Ventana = "Pantalla: Detalle No conformidad";
                    }
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                             oNetDrive1.Force = true;  oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                       //     MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);   
                    }
                    oNetDrive = null;
                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!central10.Controls.Contains(Consulta_Oportunidades_Mejora.Instance))
                        {
                            Ventana = "Pantalla: Detalle No conformidad";


                            Consulta_Oportunidades_Mejora.Instance.Nombre = Nombre;
                            Consulta_Oportunidades_Mejora.Instance.Departamento = Departamento;
                            Consulta_Oportunidades_Mejora.Instance.Ventana1 = Ventana;

                            central10.Controls.Add(Consulta_Oportunidades_Mejora.Instance);
                            Consulta_Oportunidades_Mejora.Instance.Dock = DockStyle.Fill;
                            Consulta_Oportunidades_Mejora.Instance.BringToFront();



                        }
                        else
                        {
                            Consulta_OportunidadesMejora_General.Instance.BringToFront();
                            Ventana = "Pantalla: Detalle No conformidad";
                        }
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


            }

        private void toolStripMenuItem182_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else { conexionvpn = 2;
                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!central11.Controls.Contains(Genera_No_Confor.Instance))
                    {
                        Ventana = "Pantalla: Levantamiento de No conformidad";
                        Genera_No_Confor.Instance.Nombre = Nombre;
                        Genera_No_Confor.Instance.Ventana1 = Ventana;

                        central11.Controls.Add(Genera_No_Confor.Instance);
                        Genera_No_Confor.Instance.Dock = DockStyle.Fill;
                        Genera_No_Confor.Instance.BringToFront();


                    }
                    else
                    {
                        Genera_No_Confor.Instance.BringToFront();
                        Ventana = "Pantalla: Levantamiento de No conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                             oNetDrive1.Force = true;  oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                      //      MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;


                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);     
                    }
                    oNetDrive = null;
                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!central11.Controls.Contains(Genera_No_Confor.Instance))
                        {
                            Ventana = "Pantalla: Levantamiento de No conformidad";
                            Genera_No_Confor.Instance.Nombre = Nombre;
                            Genera_No_Confor.Instance.Ventana1 = Ventana;

                            central11.Controls.Add(Genera_No_Confor.Instance);
                            Genera_No_Confor.Instance.Dock = DockStyle.Fill;
                            Genera_No_Confor.Instance.BringToFront();


                        }
                        else
                        {
                            Genera_No_Confor.Instance.BringToFront();
                            Ventana = "Pantalla: Levantamiento de No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }

     
        }

        private void toolStripMenuItem183_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else { conexionvpn = 2;
                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5") if (Permiso == "1" || Permiso == "2")
                    {

                        if (!central11.Controls.Contains(Consulta_No_Conf.Instance))
                        {
                            Ventana = "Pantalla: Consulta de No Conformidad";
                            Consulta_No_Conf.Instance.Departamento = Departamento;
                            central11.Controls.Add(Consulta_No_Conf.Instance);
                            Consulta_No_Conf.Instance.Dock = DockStyle.Fill;
                            Consulta_No_Conf.Instance.BringToFront();


                        }
                        else
                        {
                            Consulta_No_Conf.Instance.BringToFront();
                            Ventana = "Pantalla: Consulta de No Conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                             oNetDrive1.Force = true;  oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                           // MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;




                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;
                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5") if (Permiso == "1" || Permiso == "2")
                        {

                            if (!central11.Controls.Contains(Consulta_No_Conf.Instance))
                            {
                                Ventana = "Pantalla: Consulta de No Conformidad";
                                Consulta_No_Conf.Instance.Departamento = Departamento;
                                central11.Controls.Add(Consulta_No_Conf.Instance);
                                Consulta_No_Conf.Instance.Dock = DockStyle.Fill;
                                Consulta_No_Conf.Instance.BringToFront();


                            }
                            else
                            {
                                Consulta_No_Conf.Instance.BringToFront();
                                Ventana = "Pantalla: Consulta de No Conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }

        
        }

        private void toolStripMenuItem184_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else { conexionvpn = 2;
                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5") if (Permiso == "1" || Permiso == "2")
                    {

                        if (!central11.Controls.Contains(Consulta_No_Conf_general.Instance))
                        {
                            Ventana = "Pantalla: Consulta de No Conformidad";

                            central11.Controls.Add(Consulta_No_Conf_general.Instance);
                            Consulta_No_Conf_general.Instance.Dock = DockStyle.Fill;
                            Consulta_No_Conf_general.Instance.BringToFront();


                        }
                        else
                        {
                            Consulta_No_Conf_general.Instance.BringToFront();
                            Ventana = "Pantalla: Consulta de No Conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {

                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                             oNetDrive1.Force = true;  oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                         //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;




                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);   
                    }
                    oNetDrive = null;
                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5") if (Permiso == "1" || Permiso == "2")
                        {

                            if (!central11.Controls.Contains(Consulta_No_Conf_general.Instance))
                            {
                                Ventana = "Pantalla: Consulta de No Conformidad";

                                central11.Controls.Add(Consulta_No_Conf_general.Instance);
                                Consulta_No_Conf_general.Instance.Dock = DockStyle.Fill;
                                Consulta_No_Conf_general.Instance.BringToFront();


                            }
                            else
                            {
                                Consulta_No_Conf_general.Instance.BringToFront();
                                Ventana = "Pantalla: Consulta de No Conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }

         
        }

        private void toolStripMenuItem185_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else { conexionvpn = 2;
                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!central11.Controls.Contains(Seguimiento_No_confor.Instance))
                    {
                        Ventana = "Pantalla: Seguimiento de No conformidad";
                        Seguimiento_No_confor.Instance.Departamento = Departamento;
                        Seguimiento_No_confor.Instance.Nombre = Nombre;
                        Seguimiento_No_confor.Instance.Apellido = Apellido;
                        Seguimiento_No_confor.Instance.NoConfor = NumeroConformidad;
                        Seguimiento_No_confor.Instance.Ventana1 = Ventana;
                        central11.Controls.Add(Seguimiento_No_confor.Instance);
                        Seguimiento_No_confor.Instance.Dock = DockStyle.Fill;
                        Seguimiento_No_confor.Instance.BringToFront();


                    }
                    else
                    {
                        Seguimiento_No_confor.Instance.BringToFront();
                        Ventana = "Pantalla: Seguimiento de No conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {

                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                             oNetDrive1.Force = true;  oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                       //     MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;


                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;
                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!central11.Controls.Contains(Seguimiento_No_confor.Instance))
                        {
                            Ventana = "Pantalla: Seguimiento de No conformidad";
                            Seguimiento_No_confor.Instance.Departamento = Departamento;
                            Seguimiento_No_confor.Instance.Nombre = Nombre;
                            Seguimiento_No_confor.Instance.Apellido = Apellido;
                            Seguimiento_No_confor.Instance.NoConfor = NumeroConformidad;
                            Seguimiento_No_confor.Instance.Ventana1 = Ventana;
                            central11.Controls.Add(Seguimiento_No_confor.Instance);
                            Seguimiento_No_confor.Instance.Dock = DockStyle.Fill;
                            Seguimiento_No_confor.Instance.BringToFront();


                        }
                        else
                        {
                            Seguimiento_No_confor.Instance.BringToFront();
                            Ventana = "Pantalla: Seguimiento de No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }

    
        }

        private void toolStripMenuItem199_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else { conexionvpn = 2; }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                             oNetDrive1.Force = true;  oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                         //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;



                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);   
                    }
                    oNetDrive = null;
                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!central11.Controls.Contains(Seguimiento_AM.Instance))
                        {
                            Ventana = "Pantalla: Detalle No conformidad";
                            Seguimiento_AM.Instance.Nombre = Nombre;
                            Seguimiento_AM.Instance.Apellido = Apellido;
                            Seguimiento_AM.Instance.Departamento = Departamento;
                            Seguimiento_AM.Instance.Ventana1 = Ventana;
                            central11.Controls.Add(Seguimiento_AM.Instance);
                            Seguimiento_AM.Instance.Dock = DockStyle.Fill;
                            Seguimiento_AM.Instance.BringToFront();


                        }
                        else
                        {
                            Seguimiento_AM.Instance.BringToFront();
                            Ventana = "Pantalla: Detalle No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }

                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }

     
        }

        private void toolStripMenuItem198_Click(object sender, EventArgs e)
        {
               if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else { conexionvpn = 2;
                if (!central11.Controls.Contains(Genera_AM.Instance))
                {
                    Ventana = "Pantalla: Genera Oportunidad de Mejora";
                    Genera_AM.Instance.Departamento = Departamento;
                    Genera_AM.Instance.Nombre = Nombre;
                    Genera_AM.Instance.Ventana1 = Ventana;
                    Genera_AM.Instance.NumeroAC = NumeroOport;
                    central11.Controls.Add(Genera_AM.Instance);
                    Genera_AM.Instance.Dock = DockStyle.Fill;
                    Genera_AM.Instance.BringToFront();


                }
                else
                {
                    Genera_AM.Instance.BringToFront();
                    Ventana = "Pantalla: Genera Oportunidad de Mejora";
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive1 = new NetworkDrive();
                    try
                    {
                        //set propertys
                         oNetDrive1.Force = true;  oNetDrive1.LocalDrive = "G";
                        oNetDrive1.ShareName = url.Text;
                        //match call to options provided
                        oNetDrive1.MapDrive();
                        conexionvpn = 1;
                        //update status
                    }
                    catch (Exception err)
                    {
                        //report error

                    //    MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                    }
                    oNetDrive1 = null;








                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        oNetDrive.LocalDrive = url.Text;
                        oNetDrive.UnMapDrive();

                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);
                    }
                    oNetDrive = null;
                    if (!central11.Controls.Contains(Genera_AM.Instance))
                    {
                        Ventana = "Pantalla: Genera Oportunidad de Mejora";
                        Genera_AM.Instance.Departamento = Departamento;
                        Genera_AM.Instance.Nombre = Nombre;
                        Genera_AM.Instance.Ventana1 = Ventana;
                        Genera_AM.Instance.NumeroAC = NumeroOport;
                        central11.Controls.Add(Genera_AM.Instance);
                        Genera_AM.Instance.Dock = DockStyle.Fill;
                        Genera_AM.Instance.BringToFront();


                    }
                    else
                    {
                        Genera_AM.Instance.BringToFront();
                        Ventana = "Pantalla: Genera Oportunidad de Mejora";
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }

      
        
          

}

        private void toolStripMenuItem200_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else { conexionvpn = 2;
                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!central11.Controls.Contains(Consulta_OportunidadesMejora_General.Instance))
                    {
                        Ventana = "Pantalla: Detalle No conformidad";

                        central11.Controls.Add(Consulta_OportunidadesMejora_General.Instance);
                        Consulta_OportunidadesMejora_General.Instance.Dock = DockStyle.Fill;
                        Consulta_OportunidadesMejora_General.Instance.BringToFront();


                    }
                    else
                    {
                        Consulta_OportunidadesMejora_General.Instance.BringToFront();
                        Ventana = "Pantalla: Detalle No conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                             oNetDrive1.Force = true;  oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                         //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;


                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {
                        //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;
                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!central11.Controls.Contains(Consulta_OportunidadesMejora_General.Instance))
                        {
                            Ventana = "Pantalla: Detalle No conformidad";

                            central11.Controls.Add(Consulta_OportunidadesMejora_General.Instance);
                            Consulta_OportunidadesMejora_General.Instance.Dock = DockStyle.Fill;
                            Consulta_OportunidadesMejora_General.Instance.BringToFront();


                        }
                        else
                        {
                            Consulta_OportunidadesMejora_General.Instance.BringToFront();
                            Ventana = "Pantalla: Detalle No conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }

  
        }

        private void toolStripMenuItem201_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else { conexionvpn = 2;
                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!central11.Controls.Contains(Consulta_Oportunidades_Mejora.Instance))
                    {
                        Ventana = "Pantalla: Detalle No conformidad";


                        Consulta_Oportunidades_Mejora.Instance.Nombre = Nombre;
                        Consulta_Oportunidades_Mejora.Instance.Departamento = Departamento;
                        Consulta_Oportunidades_Mejora.Instance.Ventana1 = Ventana;

                        central11.Controls.Add(Consulta_Oportunidades_Mejora.Instance);
                        Consulta_Oportunidades_Mejora.Instance.Dock = DockStyle.Fill;
                        Consulta_Oportunidades_Mejora.Instance.BringToFront();



                    }
                    else
                    {
                        Consulta_OportunidadesMejora_General.Instance.BringToFront();
                        Ventana = "Pantalla: Detalle No conformidad";
                    }
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {

                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                             oNetDrive1.Force = true;  oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                       //     MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;


                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;
                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!central11.Controls.Contains(Consulta_Oportunidades_Mejora.Instance))
                        {
                            Ventana = "Pantalla: Detalle No conformidad";


                            Consulta_Oportunidades_Mejora.Instance.Nombre = Nombre;
                            Consulta_Oportunidades_Mejora.Instance.Departamento = Departamento;
                            Consulta_Oportunidades_Mejora.Instance.Ventana1 = Ventana;

                            central11.Controls.Add(Consulta_Oportunidades_Mejora.Instance);
                            Consulta_Oportunidades_Mejora.Instance.Dock = DockStyle.Fill;
                            Consulta_Oportunidades_Mejora.Instance.BringToFront();



                        }
                        else
                        {
                            Consulta_OportunidadesMejora_General.Instance.BringToFront();
                            Ventana = "Pantalla: Detalle No conformidad";
                        }
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }

    
        }

        private void toolStripMenuItem164_Click(object sender, EventArgs e)
        {
            if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
            {

                if (!Central12.Controls.Contains(Sistema.Password.Instance))
                {
                    Ventana = "Pantalla: Detalle No conformidad";


                   CBR_ADMIN.Sistema.Password.Instance.Nombre = Nombre;
                    Sistema.Password.Instance.Apellido = Apellido;
                 

                    Central12.Controls.Add(Sistema.Password.Instance);
                    Sistema.Password.Instance.Dock = DockStyle.Fill;
                    Sistema.Password.Instance.BringToFront();



                }
                else
                {
                    Consulta_OportunidadesMejora_General.Instance.BringToFront();
                    Ventana = "Pantalla: Detalle No conformidad";
                }
            }
        }

        private void exploradorDeArchivosToolStripMenuItem_Click(object sender, EventArgs e)
        {
        
        }
   
        private void explorarProyectosToolStripMenuItem_Click(object sender, EventArgs e)
        {


            NetworkDrive oNetDrive = new NetworkDrive();

            try
            {
                //set propertys

                oNetDrive.LocalDrive = "G";
                oNetDrive.ShareName = url.Text;
                //match call to options provided

                oNetDrive.MapDrive();
                conexionvpn = 1;

                //update status

            }
            catch (Exception err)
            {
                //report error
//
           //     MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
            }
            oNetDrive = null;

            if (!Central8.Controls.Contains(windowsNavegador.Instance))
            {
                Ventana = "Pantalla: Consulta de No Conformidad";

              
         
                windowsNavegador.Instance.usuario = Nombre;
                windowsNavegador.Instance.emailusuario = Email;
                windowsNavegador.Instance.apellido = Apellido;
                windowsNavegador.Instance.departamento = Departamento;
                Central8.Controls.Add(windowsNavegador.Instance);
                windowsNavegador.Instance.Dock = DockStyle.Fill;
                windowsNavegador.Instance.BringToFront();
            }
            else
            {
                windowsNavegador.Instance.BringToFront();
                Ventana = "Pantalla: Consulta de No Conformidad";
            }
        }

        private void crearNuevoProyectoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            NetworkDrive oNetDrive = new NetworkDrive();

           
              Validacion_carpetas crear = new Validacion_carpetas();

                crear.Show();

 
          
        }

        private void toolStripMenuItem173_Click(object sender, EventArgs e)
        {
        
            Form2 conecta = new Form2();

            string envio="2";
            conecta.BringToFront();
            conecta.WindowState = FormWindowState.Normal;
            conecta.Close();
            Form2 conecta2 = new Form2();

            conecta2.recep = envio;
            conecta2.Show();






        }

        private void cambioDeLideresToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
            if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "Administrador TI")
            {

                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central8.Controls.Contains(Actualiza_lideres.Instance))
                    {
                        Ventana = "Pantalla: Levantamiento de No conformidad";


                        Central8.Controls.Add(Actualiza_lideres.Instance);
                        Actualiza_lideres.Instance.Dock = DockStyle.Fill;
                        Actualiza_lideres.Instance.BringToFront();


                    }
                    else
                    {
                        Actualiza_lideres.Instance.BringToFront();
                        Ventana = "Pantalla: Levantamiento de No conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            else { MessageBox.Show("No tiene acceso a este modulo"); }
        }

        private void toolStripMenuItem166_Click(object sender, EventArgs e)
        {
            NetworkDrive oNetDrive = new NetworkDrive();

            try
            {
                //set propertys

                oNetDrive.LocalDrive = "G";
                oNetDrive.ShareName = url.Text;
                //match call to options provided

                oNetDrive.MapDrive();
                conexionvpn = 1;

                //update status

            }
            catch (Exception err)
            {
                //report error

                //    MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
            }
            oNetDrive = null;




            if (!Central.Controls.Contains(windowsNavegador.Instance))
            {
                Ventana = "Pantalla: Consulta de No Conformidad";



                windowsNavegador.Instance.usuario = Nombre;
                windowsNavegador.Instance.emailusuario = Email;
                windowsNavegador.Instance.departamento = Departamento;
                windowsNavegador.Instance.apellido = Apellido;
                Central.Controls.Add(windowsNavegador.Instance);
                windowsNavegador.Instance.Dock = DockStyle.Fill;
                windowsNavegador.Instance.BringToFront();
            }
            else
            {
                windowsNavegador.Instance.BringToFront();
                Ventana = "Pantalla: Consulta de No Conformidad";
            }
        }


  
        private void toolStripMenuItem167_Click(object sender, EventArgs e)
        {
            Validacion_carpetas crear = new Validacion_carpetas();

            crear.Show();
        }

        private void ticketsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else { conexionvpn = 2; }         
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                { NetworkDrive oNetDrive = new NetworkDrive();

                try
                    {
                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        {StartInfo ={FileName = FolderPath + "\\VpnDisconnect.bat",WindowStyle = ProcessWindowStyle.Normal}};
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close(); }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;}
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }
        }

        private void toolStripMenuItem21_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else { conexionvpn = 2; }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {
                        //MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); 
                    }
                    oNetDrive = null;
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }
        }

        private void noConformidadesToolStripMenuItem5_Click_1(object sender, EventArgs e)
        {

        }

        private void toolStripMenuItem121_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (!Central8.Controls.Contains(IT.Administracion.Instance))
                {
                    CBR_ADMIN.IT.Administracion.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Administracion.Instance.Apellido = Apellido;

                    Central8.Controls.Add(CBR_ADMIN.IT.Administracion.Instance);
                    CBR_ADMIN.IT.Administracion.Instance.Dock = DockStyle.Fill;
                    CBR_ADMIN.IT.Administracion.Instance.BringToFront();


                }
                else
                {
                    CBR_ADMIN.IT.Administracion.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                             oNetDrive1.Force = true;  oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                           // MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  
                    }
                    oNetDrive = null;

                    if (!Central7.Controls.Contains(IT.Administracion.Instance))
                    {
                        CBR_ADMIN.IT.Administracion.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Administracion.Instance.Apellido = Apellido;

                        Central7.Controls.Add(CBR_ADMIN.IT.Administracion.Instance);
                        CBR_ADMIN.IT.Administracion.Instance.Dock = DockStyle.Fill;
                        CBR_ADMIN.IT.Administracion.Instance.BringToFront();


                    }
                    else
                    {
                        CBR_ADMIN.IT.Administracion.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void toolStripMenuItem122_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (!Central7.Controls.Contains(IT.Aplicativo.Instance))
                {
                    CBR_ADMIN.IT.Aplicativo.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Aplicativo.Instance.Apellido = Apellido;
                    Central7.Controls.Add(IT.Aplicativo.Instance);
                    IT.Aplicativo.Instance.Dock = DockStyle.Fill;
                    IT.Aplicativo.Instance.BringToFront();


                }
                else
                {
                    IT.Administracion.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                             oNetDrive1.Force = true;  oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                         //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); 
                    }
                    oNetDrive = null;
                    if (!Central8.Controls.Contains(IT.Aplicativo.Instance))
                    {
                        CBR_ADMIN.IT.Aplicativo.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Aplicativo.Instance.Apellido = Apellido;
                        Central8.Controls.Add(IT.Aplicativo.Instance);
                        IT.Aplicativo.Instance.Dock = DockStyle.Fill;
                        IT.Aplicativo.Instance.BringToFront();


                    }
                    else
                    {
                        IT.Administracion.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void toolStripMenuItem186_Click(object sender, EventArgs e)
        {

        }

        private void toolStripMenuItem190_Click(object sender, EventArgs e)
        {

            NetworkDrive oNetDrive = new NetworkDrive();

            try
            {
                //set propertys

                oNetDrive.LocalDrive = "G";
                oNetDrive.ShareName = url.Text;
                //match call to options provided

                oNetDrive.MapDrive();
                conexionvpn = 1;

                //update status

            }
            catch (Exception err)
            {
                //report error

                //    MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
            }
            oNetDrive = null;

            if (!Central4.Controls.Contains(windowsNavegador.Instance))
            {
                Ventana = "Pantalla: Consulta de No Conformidad";

                
       
                windowsNavegador.Instance.usuario = Nombre;
                windowsNavegador.Instance.emailusuario = Email;
                windowsNavegador.Instance.apellido = Apellido;
                windowsNavegador.Instance.departamento = Departamento;
                Central4.Controls.Add(windowsNavegador.Instance);
                windowsNavegador.Instance.Dock = DockStyle.Fill;
                windowsNavegador.Instance.BringToFront();
            }
            else
            {
                windowsNavegador.Instance.BringToFront();
                Ventana = "Pantalla: Consulta de No Conformidad";
            }
        }

        private void toolStripMenuItem195_Click(object sender, EventArgs e)
        {

            NetworkDrive oNetDrive = new NetworkDrive();

            try
            {
                //set propertys

                oNetDrive.LocalDrive = "G";
                oNetDrive.ShareName = url.Text;
                //match call to options provided

                oNetDrive.MapDrive();
                conexionvpn = 1;

                //update status

            }
            catch (Exception err)
            {
                //report error

            //    MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
            }
            oNetDrive = null;


            if (!Central5.Controls.Contains(windowsNavegador.Instance))
            {
                Ventana = "Pantalla: Consulta de No Conformidad";

            
             
                windowsNavegador.Instance.usuario = Nombre;
                windowsNavegador.Instance.emailusuario = Email;
                windowsNavegador.Instance.apellido = Apellido;
                windowsNavegador.Instance.departamento = Departamento;
                Central5.Controls.Add(windowsNavegador.Instance);
                windowsNavegador.Instance.Dock = DockStyle.Fill;
                windowsNavegador.Instance.BringToFront();
            }
            else
            {
                windowsNavegador.Instance.BringToFront();
                Ventana = "Pantalla: Consulta de No Conformidad";
            }
        }

        private void toolStripMenuItem204_Click(object sender, EventArgs e)
        {

            NetworkDrive oNetDrive = new NetworkDrive();

            try
            {
                //set propertys

                oNetDrive.LocalDrive = "G";
                oNetDrive.ShareName = url.Text;
                //match call to options provided

                oNetDrive.MapDrive();
                conexionvpn = 1;

                //update status

            }
            catch (Exception err)
            {
                //report error

                //    MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
            }
            oNetDrive = null;


            if (!Central6.Controls.Contains(windowsNavegador.Instance))
            {
                Ventana = "Pantalla: Consulta de No Conformidad";

          
          
                windowsNavegador.Instance.usuario = Nombre;
                windowsNavegador.Instance.emailusuario = Email;
                windowsNavegador.Instance.apellido = Apellido;
                windowsNavegador.Instance.departamento = Departamento;
                Central6.Controls.Add(windowsNavegador.Instance);
                windowsNavegador.Instance.Dock = DockStyle.Fill;
                windowsNavegador.Instance.BringToFront();
            }
            else
            {
                windowsNavegador.Instance.BringToFront();
                Ventana = "Pantalla: Consulta de No Conformidad";
            }
        }

        private void toolStripMenuItem208_Click(object sender, EventArgs e)
        {

            NetworkDrive oNetDrive = new NetworkDrive();

            try
            {
                //set propertys

                oNetDrive.LocalDrive = "G";
                oNetDrive.ShareName = url.Text;
                //match call to options provided

                oNetDrive.MapDrive();
                conexionvpn = 1;

                //update status

            }
            catch (Exception err)
            {
                //report error

      //          MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
            }
            oNetDrive = null;

            if (!Central7.Controls.Contains(windowsNavegador.Instance))
            {
                Ventana = "Pantalla: Consulta de No Conformidad";

          
              
                windowsNavegador.Instance.usuario = Nombre;
                windowsNavegador.Instance.emailusuario = Email;
                windowsNavegador.Instance.apellido = Apellido;
                windowsNavegador.Instance.departamento = Departamento;
       
                Central7.Controls.Add(windowsNavegador.Instance);
                windowsNavegador.Instance.Dock = DockStyle.Fill;
                windowsNavegador.Instance.BringToFront();
            }
            else
            {
                windowsNavegador.Instance.BringToFront();
                Ventana = "Pantalla: Consulta de No Conformidad";
            }
        }

        private void toolStripMenuItem212_Click(object sender, EventArgs e)
        {
            NetworkDrive oNetDrive = new NetworkDrive();

            try
            {
                //set propertys

                oNetDrive.LocalDrive = "G";
                oNetDrive.ShareName = url.Text;
                //match call to options provided

                oNetDrive.MapDrive();
                conexionvpn = 1;

                //update status

            }
            catch (Exception err)
            {
                //report error

             //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
            }
            oNetDrive = null;

            if (!central10.Controls.Contains(windowsNavegador.Instance))
            {
                Ventana = "Pantalla: Consulta de No Conformidad";

             
            
                windowsNavegador.Instance.usuario = Nombre;
                windowsNavegador.Instance.emailusuario = Email;
                windowsNavegador.Instance.apellido = Apellido;
                windowsNavegador.Instance.departamento = Departamento;
                central10.Controls.Add(windowsNavegador.Instance);
                windowsNavegador.Instance.Dock = DockStyle.Fill;
                windowsNavegador.Instance.BringToFront();
            }
            else
            {
                windowsNavegador.Instance.BringToFront();
                Ventana = "Pantalla: Consulta de No Conformidad";
            }


        }

        private void toolStripMenuItem216_Click(object sender, EventArgs e)
        {
            NetworkDrive oNetDrive = new NetworkDrive();

            try
            {
                //set propertys

                oNetDrive.LocalDrive = "G";
                oNetDrive.ShareName = url.Text;
                //match call to options provided

                oNetDrive.MapDrive();
                conexionvpn = 1;

                //update status

            }
            catch (Exception err)
            {
                //report error

              //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
            }
            oNetDrive = null;


            if (!central11.Controls.Contains(windowsNavegador.Instance))
            {
                Ventana = "Pantalla: Consulta de No Conformidad";

              
             
                windowsNavegador.Instance.usuario = Nombre;
                windowsNavegador.Instance.emailusuario = Email;
                windowsNavegador.Instance.apellido = Apellido;
                windowsNavegador.Instance.departamento = Departamento;
                central11.Controls.Add(windowsNavegador.Instance);
                windowsNavegador.Instance.Dock = DockStyle.Fill;
                windowsNavegador.Instance.BringToFront();
            }
            else
            {
                windowsNavegador.Instance.BringToFront();
                Ventana = "Pantalla: Consulta de No Conformidad";
            }
        }

        private void toolStripMenuItem123_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;

                if (!Central7.Controls.Contains(IT.Optimizacion.Instance))
                {

                    CBR_ADMIN.IT.Optimizacion.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Optimizacion.Instance.Apellido = Apellido;
                    Central7.Controls.Add(CBR_ADMIN.IT.Optimizacion.Instance);
                    CBR_ADMIN.IT.Optimizacion.Instance.Dock = DockStyle.Fill;
                    CBR_ADMIN.IT.Optimizacion.Instance.BringToFront();


                }
                else
                {
                    CBR_ADMIN.IT.Optimizacion.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                             oNetDrive1.Force = true;  oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                        //    MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;


                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);     }
                    oNetDrive = null;

                    if (!Central8.Controls.Contains(IT.Optimizacion.Instance))
                    {

                        CBR_ADMIN.IT.Optimizacion.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Optimizacion.Instance.Apellido = Apellido;
                        Central8.Controls.Add(CBR_ADMIN.IT.Optimizacion.Instance);
                        CBR_ADMIN.IT.Optimizacion.Instance.Dock = DockStyle.Fill;
                        CBR_ADMIN.IT.Optimizacion.Instance.BringToFront();


                    }
                    else
                    {
                        CBR_ADMIN.IT.Optimizacion.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }
        }

        private void toolStripMenuItem124_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;

                if (!Central7.Controls.Contains(IT.Otros.Instance))
                {
                    CBR_ADMIN.IT.Otros.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Otros.Instance.Apellido = Apellido;
                    Central7.Controls.Add(CBR_ADMIN.IT.Otros.Instance);
                    CBR_ADMIN.IT.Otros.Instance.Dock = DockStyle.Fill;
                    CBR_ADMIN.IT.Otros.Instance.BringToFront();


                }
                else
                {
                    CBR_ADMIN.IT.Otros.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                             oNetDrive1.Force = true;  oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                       //     MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;


                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {  MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);     }
                    oNetDrive = null;

                    if (!Central7.Controls.Contains(IT.Otros.Instance))
                    {
                        CBR_ADMIN.IT.Otros.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Otros.Instance.Apellido = Apellido;
                        Central8.Controls.Add(CBR_ADMIN.IT.Otros.Instance);
                        CBR_ADMIN.IT.Otros.Instance.Dock = DockStyle.Fill;
                        CBR_ADMIN.IT.Otros.Instance.BringToFront();


                    }
                    else
                    {
                        CBR_ADMIN.IT.Otros.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }

        }

        private void toolStripMenuItem125_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (!Central7.Controls.Contains(IT.Servicio.Instance))
                {
                    CBR_ADMIN.IT.Servicio.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Servicio.Instance.Apellido = Apellido;
                    Central7.Controls.Add(CBR_ADMIN.IT.Servicio.Instance);
                    CBR_ADMIN.IT.Servicio.Instance.Dock = DockStyle.Fill;
                    CBR_ADMIN.IT.Servicio.Instance.BringToFront();


                }
                else
                {
                    CBR_ADMIN.IT.Servicio.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {

                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                             oNetDrive1.Force = true;  oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                       //     MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); 
                    }
                    oNetDrive = null;
                    if (!Central8.Controls.Contains(IT.Servicio.Instance))
                    {
                        CBR_ADMIN.IT.Servicio.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Servicio.Instance.Apellido = Apellido;
                        Central8.Controls.Add(CBR_ADMIN.IT.Servicio.Instance);
                        CBR_ADMIN.IT.Servicio.Instance.Dock = DockStyle.Fill;
                        CBR_ADMIN.IT.Servicio.Instance.BringToFront();


                    }
                    else
                    {
                        CBR_ADMIN.IT.Servicio.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }

        }

        private void toolStripMenuItem126_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (!Central7.Controls.Contains(IT.Soporte.Instance))
                {
                    CBR_ADMIN.IT.Soporte.Instance.Nombre = Nombre;
                    CBR_ADMIN.IT.Soporte.Instance.Apellido = Apellido;
                    Central7.Controls.Add(CBR_ADMIN.IT.Soporte.Instance);
                    CBR_ADMIN.IT.Soporte.Instance.Dock = DockStyle.Fill;
                    CBR_ADMIN.IT.Soporte.Instance.BringToFront();


                }
                else
                {
                    CBR_ADMIN.IT.Soporte.Instance.BringToFront();

                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                             oNetDrive1.Force = true;  oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                          //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;


                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message); 
                    }
                    oNetDrive = null;
                    if (!Central8.Controls.Contains(IT.Soporte.Instance))
                    {
                        CBR_ADMIN.IT.Soporte.Instance.Nombre = Nombre;
                        CBR_ADMIN.IT.Soporte.Instance.Apellido = Apellido;
                        Central8.Controls.Add(CBR_ADMIN.IT.Soporte.Instance);
                        CBR_ADMIN.IT.Soporte.Instance.Dock = DockStyle.Fill;
                        CBR_ADMIN.IT.Soporte.Instance.BringToFront();


                    }
                    else
                    {
                        CBR_ADMIN.IT.Soporte.Instance.BringToFront();

                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }


        }

        private void toolStripMenuItem148_Click(object sender, EventArgs e)
        {

        }

        private void toolStripMenuItem186_Click_1(object sender, EventArgs e)
        {

            NetworkDrive oNetDrive = new NetworkDrive();

            try
            {
                //set propertys

                oNetDrive.LocalDrive = "G";
                oNetDrive.ShareName = url.Text;
                //match call to options provided

                oNetDrive.MapDrive();
                conexionvpn = 1;

                //update status

            }
            catch (Exception err)
            {
                //report error

             //   MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
            }
            oNetDrive = null;



            if (!Central3.Controls.Contains(windowsNavegador.Instance))
            {
                Ventana = "Pantalla: Consulta de No Conformidad";
         
              
    
                windowsNavegador.Instance.usuario = Nombre;
                windowsNavegador.Instance.emailusuario = Email;
                windowsNavegador.Instance.apellido = Apellido;
                windowsNavegador.Instance.departamento = Departamento;
                Central3.Controls.Add(windowsNavegador.Instance);
                windowsNavegador.Instance.Dock = DockStyle.Fill;
                windowsNavegador.Instance.BringToFront();
            }
            else
            {
                windowsNavegador.Instance.BringToFront();
                Ventana = "Pantalla: Consulta de No Conformidad";
            }
        }

        private void cambioDeLideresToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "Asistente de direccion")
            {

                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central.Controls.Contains(Actualiza_lideres.Instance))
                    {
                        Ventana = "Pantalla: Levantamiento de No conformidad";


                        Central.Controls.Add(Actualiza_lideres.Instance);
                        Actualiza_lideres.Instance.Dock = DockStyle.Fill;
                        Actualiza_lideres.Instance.BringToFront();


                    }
                    else
                    {
                        Actualiza_lideres.Instance.BringToFront();
                        Ventana = "Pantalla: Levantamiento de No conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            else { MessageBox.Show("No tiene acceso a este modulo"); }
        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void consultaToolStripMenuItem9_Click(object sender, EventArgs e)
        {
            if (Departamento == "Asistente de direccion" || Departamento == "SuperAdmin" || Departamento == "Direccion")
            {

                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central.Controls.Contains(Consulta_No_Conf.Instance))
                    {
                        Ventana = "Pantalla: Consulta de No Conformidad";
                        Consulta_No_Conf.Instance.Departamento = Departamento;
                        Central.Controls.Add(Consulta_No_Conf.Instance);
                        Consulta_No_Conf.Instance.Dock = DockStyle.Fill;
                        Consulta_No_Conf.Instance.BringToFront();


                    }
                    else
                    {
                        Consulta_No_Conf.Instance.BringToFront();
                        Ventana = "Pantalla: Consulta de No Conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }

            else
            {
                MessageBox.Show("No tiene acceso a este modulo");
            }
        }

        private void generalToolStripMenuItem9_Click(object sender, EventArgs e)
        {

            if (Departamento == "Asistente de direccion" || Departamento == "SuperAdmin" || Departamento == "Direccion")
            {

                if (Permiso == "1" || Permiso == "2")
                {

                    if (!Central.Controls.Contains(Consulta_No_Conf_general.Instance))
                    {
                        Ventana = "Pantalla: Consulta de No Conformidad";

                        Central.Controls.Add(Consulta_No_Conf_general.Instance);
                        Consulta_No_Conf_general.Instance.Dock = DockStyle.Fill;
                        Consulta_No_Conf_general.Instance.BringToFront();


                    }
                    else
                    {
                        Consulta_No_Conf_general.Instance.BringToFront();
                        Ventana = "Pantalla: Consulta de No Conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            else
            {
                MessageBox.Show("No tiene acceso a este modulo");
            }
        }

        private void Nombreprin_Click(object sender, EventArgs e)
        {

        }

        private void horra_Click(object sender, EventArgs e)
        {

        }

        private void toolStripMenuItem165_Click(object sender, EventArgs e)
        {

        }

        private void crearNuevoProyectoToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Validacion_carpetas crear = new Validacion_carpetas();

            crear.Show();
        }

        private void cambioDeLideresToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            if (Departamento == "SuperAdmin" || Departamento == "Direccion" || Puesto == "GERENTE DE GERENCIA COMERCIAL")
            {

                if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                {

                    if (!Central.Controls.Contains(Actualiza_lideres.Instance))
                    {
                        Ventana = "Pantalla: Levantamiento de No conformidad";


                        Central5.Controls.Add(Actualiza_lideres.Instance);
                        Actualiza_lideres.Instance.Dock = DockStyle.Fill;
                        Actualiza_lideres.Instance.BringToFront();


                    }
                    else
                    {
                        Actualiza_lideres.Instance.BringToFront();
                        Ventana = "Pantalla: Levantamiento de No conformidad";
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            else { MessageBox.Show("No tiene acceso a este modulo"); }
        }

        private void generaciónToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Departamento == "Compras" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                {


                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {


                        if (!Central.Controls.Contains(Genera_No_Confor.Instance))
                        {
                            Ventana = "Pantalla: Levantamiento de No conformidad";
                            Genera_No_Confor.Instance.Nombre = Nombre;
                            Genera_No_Confor.Instance.Ventana1 = Ventana;

                            Genera_No_Confor.Instance.Departamento = Departamento;
                            Central.Controls.Add(Genera_No_Confor.Instance);
                            Genera_No_Confor.Instance.Dock = DockStyle.Fill;
                            Genera_No_Confor.Instance.BringToFront();


                        }
                        else
                        {
                            Genera_No_Confor.Instance.BringToFront();
                            Ventana = "Pantalla: Levantamiento de No conformidad";

                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");

                    }
                }
                else { MessageBox.Show("No tiene acceso a este modulo"); }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                        //    MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();
                    }
                    catch (Exception err)
                    {//  MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);     
                    }
                    oNetDrive = null;
                    if (Departamento == "Compras" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                    {


                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                        {


                            if (!Central.Controls.Contains(Genera_No_Confor.Instance))
                            {
                                Ventana = "Pantalla: Levantamiento de No conformidad";
                                Genera_No_Confor.Instance.Nombre = Nombre;
                                Genera_No_Confor.Instance.Ventana1 = Ventana;

                                Genera_No_Confor.Instance.Departamento = Departamento;
                                Central.Controls.Add(Genera_No_Confor.Instance);
                                Genera_No_Confor.Instance.Dock = DockStyle.Fill;
                                Genera_No_Confor.Instance.BringToFront();


                            }
                            else
                            {
                                Genera_No_Confor.Instance.BringToFront();
                                Ventana = "Pantalla: Levantamiento de No conformidad";

                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");

                        }
                    }
                    else { MessageBox.Show("No tiene acceso a este modulo"); }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }
        }

        private void noConformidadesToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void toolStripMenuItem181_Click(object sender, EventArgs e)
        {

        }

        private void toolStripMenuItem168_Click(object sender, EventArgs e)
        {

        }

        private void generacionToolStripMenuItem2_Click(object sender, EventArgs e)
        {
      

                    if (Permiso == "1" || Permiso == "2" )
                    {

                        if (!central13.Controls.Contains(Genera_No_Confor.Instance))
                        {
                            Ventana = "Pantalla: Levantamiento de No conformidad";
                            Genera_No_Confor.Instance.Nombre = Nombre;

                            Genera_No_Confor.Instance.Ventana1 = Ventana;
                        //Genera_No_Confor.Instance.Email = Email;
                        central13.Controls.Add(Genera_No_Confor.Instance);
                            Genera_No_Confor.Instance.Dock = DockStyle.Fill;
                            Genera_No_Confor.Instance.BringToFront();



                        }
                        else
                        {
                            Genera_No_Confor.Instance.BringToFront();
                      


                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");

                    }
         
            
        
        }

        private void cierreToolStripMenuItem2_Click_1(object sender, EventArgs e)
        {
            if (Permiso == "1" || Permiso == "2")
            {

                if (!central13.Controls.Contains(Cierre_No_confor.Instance))
                {
                    Ventana = "Pantalla: Seguimiento de No conformidad";
                    central13.Controls.Add(Cierre_No_confor.Instance);
                    Cierre_No_confor.Instance.Dock = DockStyle.Fill;
                    Cierre_No_confor.Instance.BringToFront();


                }
                else
                {
                    Cierre_No_confor.Instance.BringToFront();
                    Ventana = "Pantalla: Seguimiento de No conformidad";
                }
            }
            else
            {
                MessageBox.Show("No tiene acceso a este modulo");

            }
        }

        private void detalleToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (Permiso == "1" || Permiso == "2")
            {

                if (!central13.Controls.Contains(Detalle.Instance))
                {
                    Ventana = "Pantalla: Detalle No conformidad";
                    central13.Controls.Add(Detalle.Instance);
                    Detalle.Instance.Dock = DockStyle.Fill;
                    Detalle.Instance.BringToFront();


                }
                else
                {
                    Detalle.Instance.BringToFront();
                    Detalle.Instance.Dock = DockStyle.Fill;
                    Ventana = "Pantalla: Detalle No conformidad";
                }
            }
            else
            {
                MessageBox.Show("No tiene acceso a este modulo");
            }
        }

        private void seguimientoToolStripMenuItem9_Click(object sender, EventArgs e)
        {
            if (Permiso == "1" || Permiso == "2")
            {

                if (!central13.Controls.Contains(Seguimiento_No_confor.Instance))
            {
                Ventana = "Pantalla: Seguimiento de No conformidad";
                Seguimiento_No_confor.Instance.Departamento = Departamento;
                Seguimiento_No_confor.Instance.Nombre = Nombre;
                Seguimiento_No_confor.Instance.Apellido = Apellido;
                Seguimiento_No_confor.Instance.NoConfor = NumeroConformidad;
                Seguimiento_No_confor.Instance.Ventana1 = Ventana;
                central13.Controls.Add(Seguimiento_No_confor.Instance);
                Seguimiento_No_confor.Instance.Dock = DockStyle.Fill;
                Seguimiento_No_confor.Instance.BringToFront();


            }
            else
            {
                Seguimiento_No_confor.Instance.BringToFront();
                Ventana = "Pantalla: Seguimiento de No conformidad";
            }
            }
            else
            {
                MessageBox.Show("No tiene acceso a este modulo");
            }
        }

        private void actualizacionToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (Permiso == "1" || Permiso == "2")
            {

                if (!central13.Controls.Contains(Actualizacion_NC.Instance))
                {

                    central13.Controls.Add(Actualizacion_NC.Instance);
                    Actualizacion_NC.Instance.Dock = DockStyle.Fill;
                    Actualizacion_NC.Instance.BringToFront();


                }
                else
                {
                    Actualizacion_NC.Instance.BringToFront();
                    Ventana = "Pantalla: Genera Oportunidad de Mejora";
                }
            }
            else
            {
                MessageBox.Show("No tiene acceso a este modulo");
            }
        }

        private void generalToolStripMenuItem10_Click(object sender, EventArgs e)
        {
            if (Permiso == "1" || Permiso == "2")
            {

                if (!central13.Controls.Contains(Consulta_No_Conf_general.Instance))
                {
                    Ventana = "Pantalla: Consulta de No Conformidad";

                    central13.Controls.Add(Consulta_No_Conf_general.Instance);
                    Consulta_No_Conf_general.Instance.Dock = DockStyle.Fill;
                    Consulta_No_Conf_general.Instance.BringToFront();


                }
                else
                {
                    Consulta_No_Conf_general.Instance.BringToFront();
                    Ventana = "Pantalla: Consulta de No Conformidad";
                }
            }
            else
            {
                MessageBox.Show("No tiene acceso a este modulo");
            }
        }

        private void Central_Paint(object sender, PaintEventArgs e)
        {

        }

        private void generacionToolStripMenuItem3_Click(object sender, EventArgs e)
        {

            if (Permiso == "1" || Permiso == "2" )
            {

                if (!central13.Controls.Contains(Genera_AM.Instance))
                {
                    Ventana = "Pantalla: Genera Oportunidad de Mejora";
                    Genera_AM.Instance.Departamento = Departamento;
                    Genera_AM.Instance.Nombre = Nombre;
                    Genera_AM.Instance.Ventana1 = Ventana;
                    Genera_AM.Instance.NumeroAC = NumeroOport;
                    central13.Controls.Add(Genera_AM.Instance);
                    Genera_AM.Instance.Dock = DockStyle.Fill;
                    Genera_AM.Instance.BringToFront();


                }
                else
                {
                    Genera_AM.Instance.BringToFront();
                    Ventana = "Pantalla: Genera Oportunidad de Mejora";
                }
            }
            else
            {
                MessageBox.Show("No tiene acceso a este modulo");
            }
        }

        private void seguimientoToolStripMenuItem10_Click(object sender, EventArgs e)
        {
            if (Permiso == "1" || Permiso == "2" )
            {

                if (!central13.Controls.Contains(Seguimiento_AM.Instance))
                {
                    Ventana = "Pantalla: Detalle No conformidad";
                    Seguimiento_AM.Instance.Nombre = Nombre;
                    Seguimiento_AM.Instance.Apellido = Apellido;
                    Seguimiento_AM.Instance.Departamento = Departamento;
                    Seguimiento_AM.Instance.Ventana1 = Ventana;
                    central13.Controls.Add(Seguimiento_AM.Instance);
                    Seguimiento_AM.Instance.Dock = DockStyle.Fill;
                    Seguimiento_AM.Instance.BringToFront();


                }
                else
                {
                    Seguimiento_AM.Instance.BringToFront();
                    Ventana = "Pantalla: Detalle No conformidad";
                }
            }
            else
            {
                MessageBox.Show("No tiene acceso a este modulo");
            }
        }

        private void cierreToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            if (Permiso == "1" || Permiso == "2")
            {

                if (!central13.Controls.Contains(Ciere_Accion_Mejora.Instance))
                {
                    Ventana = "Pantalla: Detalle No conformidad";
                    Ciere_Accion_Mejora.Instance.Nombre = Nombre;
                    Ciere_Accion_Mejora.Instance.Departamento = Departamento;
                    Ciere_Accion_Mejora.Instance.Ventana1 = Ventana;
                    central13.Controls.Add(Ciere_Accion_Mejora.Instance);
                    Ciere_Accion_Mejora.Instance.Dock = DockStyle.Fill;
                    Ciere_Accion_Mejora.Instance.BringToFront();


                }
                else
                {
                    Seguimiento_AM.Instance.BringToFront();
                    Ventana = "Pantalla: Detalle No conformidad";
                }
            }
            else
            {
                MessageBox.Show("No tiene acceso a este modulo");
            }
        }

        private void generalToolStripMenuItem11_Click(object sender, EventArgs e)
        {

            if (Permiso == "1" || Permiso == "2" )
            {

                if (!central13.Controls.Contains(Consulta_OportunidadesMejora_General.Instance))
                {
                    Ventana = "Pantalla: Detalle No conformidad";

                    central13.Controls.Add(Consulta_OportunidadesMejora_General.Instance);
                    Consulta_OportunidadesMejora_General.Instance.Dock = DockStyle.Fill;
                    Consulta_OportunidadesMejora_General.Instance.BringToFront();


                }
                else
                {
                    Consulta_OportunidadesMejora_General.Instance.BringToFront();
                    Ventana = "Pantalla: Detalle No conformidad";
                }
            }
            else
            {
                MessageBox.Show("No tiene acceso a este modulo");
            }
        }

        private void exploradorProyectosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MsgBoxUtil.HackMessageBox("Local ", "Remota");

            DialogResult Resultado;
            Resultado = MessageBox.Show("Seleccione el tipo de conexión", "Navegador de proyectos", MessageBoxButtons.YesNo);



            if (Resultado == DialogResult.Yes)
            {
                NetworkDrive oNetDrive = new NetworkDrive();

                try
                {
                    //set propertys

                    oNetDrive.LocalDrive = "G";
                    oNetDrive.ShareName = url.Text;
                    //match call to options provided

                    oNetDrive.MapDrive();
                    conexionvpn = 1;

                    //update status

                }
                catch (Exception err)
                {
                    //report error

                    //    MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                }
                oNetDrive = null;

            }
            else if (Resultado == DialogResult.No)
            {


                Form2 conecta = new Form2();

                string envio = "2";
                conecta.BringToFront();
                conecta.WindowState = FormWindowState.Normal;
                conecta.Close();

                Form2 conecta2 = new Form2();

                conecta2.recep = envio;
                conecta2.Show();
                conecta2.BringToFront();
                conecta2.WindowState = FormWindowState.Normal;

                NetworkDrive oNetDrive = new NetworkDrive();
                conexionvpn = 1;

            }



            if (!central13.Controls.Contains(windowsNavegador.Instance))
            {
                Ventana = "Pantalla: Consulta de No Conformidad";

           
                windowsNavegador.Instance.usuario = Nombre;
                windowsNavegador.Instance.emailusuario = Email;
                windowsNavegador.Instance.apellido = Apellido;
                windowsNavegador.Instance.departamento = Departamento;

                central13.Controls.Add(windowsNavegador.Instance);
                windowsNavegador.Instance.Dock = DockStyle.Fill;
                windowsNavegador.Instance.BringToFront();
            }
            else
            {
                windowsNavegador.Instance.BringToFront();
                Ventana = "Pantalla: Consulta de No Conformidad";
            }
        }

        private void crearNuevoProyectoToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            Validacion_carpetas crear = new Validacion_carpetas();

            crear.Show();
        }

        private void cambioDeLideresToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            if (Permiso == "1" || Permiso == "2" )
            {

                if (!central13.Controls.Contains(Actualiza_lideres.Instance))
                {
                    Ventana = "Pantalla: Levantamiento de No conformidad";


                    central13.Controls.Add(Actualiza_lideres.Instance);
                    Actualiza_lideres.Instance.Dock = DockStyle.Fill;
                    Actualiza_lideres.Instance.BringToFront();


                }
                else
                {
                    Actualiza_lideres.Instance.BringToFront();
                    Ventana = "Pantalla: Levantamiento de No conformidad";
                }
            }
            else
            {
                MessageBox.Show("No tiene acceso a este modulo");
            }
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {

        }

        private void cargaMasivoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Permiso == "1" || Permiso == "2")
            {

                if (!central13.Controls.Contains(CargaNCmasivo.Instance))
                {
                    Ventana = "Pantalla: Levantamiento de No conformidad";
                    MessageBox.Show("¡¡ADVERTENCIA!! Notifique al area de sistema antes de cargar archivo");

                    central13.Controls.Add(CargaNCmasivo.Instance);
                    CargaNCmasivo.Instance.Dock = DockStyle.Fill;
                    CargaNCmasivo.Instance.BringToFront();


                }
                else
                {
                    CargaNCmasivo.Instance.BringToFront();
                    Ventana = "Pantalla: Levantamiento de No conformidad";
                }
            }
            else
            {
                MessageBox.Show("No tiene acceso a este modulo");
            }
        }

        private void modousuario_Click(object sender, EventArgs e)
        {
            label6.Text = "Generales:";

         
            try
            {
                ////////////////////////////////////SE ABRE LA CONEXION/////////////////////////////////////////////////////////////////////////////////
                SqlConnection conexion = new SqlConnection(ObtenerCadena());
                conexion.Open();
                ////////////////////////////////////SE EJECUTA EL COMANDO HACIA LA VISTA Y SE HACE LA VALIDACION//////////////////////////////////////
                SqlCommand cmd = new SqlCommand(
                                                "select COUNT(*) AS conteo from [No_Conformidades] where  " +
                                                "  (Status ='Abierta')" +
                                                "and (Departamento = @Departamento)"
                                                , conexion);
                cmd.Parameters.AddWithValue("Responsable", Nombre);
                cmd.Parameters.AddWithValue("Departamento", Departamento);

                SqlCommand cmd1 = new SqlCommand(
                                    "select COUNT(*) AS conteo from [No_Conformidades] where (Departamento = @Departamento)  and (Status ='Cerrada')"
                                    , conexion);
                cmd1.Parameters.AddWithValue("Departamento", Departamento);


                SqlCommand cmd2 = new SqlCommand(
                                  "select COUNT([# NC]) AS conteo from [No_Conformidades] "
                                  , conexion);


                SqlCommand cmd3 = new SqlCommand(
              "select COUNT(*) AS conteo from [No_Conformidades] where " +
              "  (Status ='Verificando')" +
              "and (Departamento = @Departamento)"
              , conexion);
                cmd3.Parameters.AddWithValue("Responsable", Nombre);
                cmd3.Parameters.AddWithValue("Departamento", Departamento);


                SqlCommand cmd4 = new SqlCommand(
           "select COUNT(*) AS conteo from [Op_Mejora] "
           , conexion);


                SqlCommand cmd5 = new SqlCommand(
                                                "select COUNT(*) AS conteo from [Op_Mejora] where  " +
                                                "  (Status ='Abierta')" +
                                                "and (Departamento = @Departamento)"
                                                , conexion);

                cmd5.Parameters.AddWithValue("Departamento", Departamento);
                SqlCommand cmd6 = new SqlCommand(
                                    "select COUNT(*) AS conteo from [Op_Mejora] where  " +
                                    "  (Status ='Cerrada') " +
                                    "and (Departamento = @Departamento)"
                                    , conexion);

                cmd6.Parameters.AddWithValue("Departamento", Departamento);


                SqlCommand cmd7 = new SqlCommand(
                                    "select COUNT(*) AS conteo from [No_Conformidades] where  " +
                                    "  (Status ='Abierta')" +
                                    "and (Departamento = 'General')"
                                    , conexion);



                SqlCommand cmd8 = new SqlCommand(
                                    "select COUNT(*) AS conteo from [No_Conformidades] where (Departamento = 'General')  and (Status ='Cerrada')"
                                    , conexion);


                SqlCommand cmd9 = new SqlCommand(
                                    "select COUNT(*) AS conteo from [No_Conformidades] where (Departamento = 'General')  and  (Status ='Verificando')"
                                    , conexion);





                Int32 rows_count = Convert.ToInt32(cmd.ExecuteScalar());
                Int32 rows_count1 = Convert.ToInt32(cmd1.ExecuteScalar());
                Int32 rows_count2 = Convert.ToInt32(cmd3.ExecuteScalar());
                Int32 rows_count3 = Convert.ToInt32(cmd5.ExecuteScalar());
                Int32 rows_count4 = Convert.ToInt32(cmd6.ExecuteScalar());
                Int32 rows_count5 = Convert.ToInt32(cmd7.ExecuteScalar());
                Int32 rows_count6 = Convert.ToInt32(cmd8.ExecuteScalar());
                Int32 rows_count7 = Convert.ToInt32(cmd9.ExecuteScalar());

                NumeroConfo = Convert.ToInt32(cmd2.ExecuteScalar());
                NumeroOport = Convert.ToInt32(cmd4.ExecuteScalar());

                cmd.Dispose();
                cmd1.Dispose();
                cmd2.Dispose();
                cmd3.Dispose();
                cmd4.Dispose();
                cmd5.Dispose();
                cmd6.Dispose();
                cmd7.Dispose();
                cmd8.Dispose();
                cmd9.Dispose();

                conexion.Close();

                NumeroConformidad = NumeroConfo.ToString();
                Nc_Abiertas.Text = rows_count.ToString();
                NC_Concluidas.Text = rows_count1.ToString();
                NC_Verificacion.Text = rows_count2.ToString();
                OP_Abiertas.Text = rows_count3.ToString();
                Op_Cerradas.Text = rows_count4.ToString();
                Ncag.Text = rows_count5.ToString();
                Ncvg.Text = rows_count7.ToString();
                Nccg.Text = rows_count6.ToString();
                Nom.Text = Nombre;
                Ape.Text = Apellido;
            }
            catch (Exception d)
            {
                MessageBox.Show(d.Message);
            }
            finally
            {

            }

            Depart.Text = Departamento;
            Pues.Text = Puesto;
            correo.Text = Email;




        }

        private void pictureBox4_Click_1(object sender, EventArgs e)
        {
            label6.Text = "Totales";

            if (Cif1 == Var1 & Cif2 == Var2 & Cif3 == Var3 & Cif4 == Var4) { acceso1 = 0; }
            else if (Cif1 == Var1 & Cif2 == Var2 & Cif3 == Var3 & Cif4 != Var4) { acceso1 = 1; }
            else if (Cif1 == Var1 & Cif2 == Var2 & Cif3 != Var3 & Cif4 != Var4) { acceso1 = 2; }
            else if (Cif1 == Var1 & Cif2 != Var2 & Cif3 != Var3 & Cif4 != Var4) { acceso1 = 3; }
            else if (Cif1 != Var1 & Cif2 != Var2 & Cif3 != Var3 & Cif4 != Var4) { acceso1 = 4; }
            Nombreprin.Text = "Bienvenido  " + "" + Nombre + "" + Apellido + " a Cbr Administracion y Servicios";

            timer1.Enabled = true;
            try
            {
                ////////////////////////////////////SE ABRE LA CONEXION/////////////////////////////////////////////////////////////////////////////////
                SqlConnection conexion = new SqlConnection(ObtenerCadena());
                conexion.Open();
                ////////////////////////////////////SE EJECUTA EL COMANDO HACIA LA VISTA Y SE HACE LA VALIDACION//////////////////////////////////////
                SqlCommand cmd = new SqlCommand(
                                                "select COUNT(*) AS conteo from [No_Conformidades] where  " +
                                                "  (Status ='Abierta')" +
                                                "and (Departamento = @Departamento)"
                                                , conexion);
                cmd.Parameters.AddWithValue("Responsable", Nombre);
                cmd.Parameters.AddWithValue("Departamento", Departamento);

                SqlCommand cmd1 = new SqlCommand(
                                    "select COUNT(*) AS conteo from [No_Conformidades] where (Departamento = @Departamento)  and (Status ='Cerrada')"
                                    , conexion);
                cmd1.Parameters.AddWithValue("Departamento", Departamento);


                SqlCommand cmd2 = new SqlCommand(
                                  "select COUNT([# NC]) AS conteo from [No_Conformidades] "
                                  , conexion);


                SqlCommand cmd3 = new SqlCommand(
              "select COUNT(*) AS conteo from [No_Conformidades] where " +
              "  (Status ='Verificando')" +
              "and (Departamento = @Departamento)"
              , conexion);
                cmd3.Parameters.AddWithValue("Responsable", Nombre);
                cmd3.Parameters.AddWithValue("Departamento", Departamento);


                SqlCommand cmd4 = new SqlCommand(
           "select COUNT(*) AS conteo from [Op_Mejora] "
           , conexion);


                SqlCommand cmd5 = new SqlCommand(
                                                "select COUNT(*) AS conteo from [Op_Mejora] where  " +
                                                "  (Status ='Abierta')" +
                                                "and (Departamento = @Departamento)"
                                                , conexion);

                cmd5.Parameters.AddWithValue("Departamento", Departamento);
                SqlCommand cmd6 = new SqlCommand(
                                    "select COUNT(*) AS conteo from [Op_Mejora] where  " +
                                    "  (Status ='Cerrada') " +
                                    "and (Departamento = @Departamento)"
                                    , conexion);

                cmd6.Parameters.AddWithValue("Departamento", Departamento);


                SqlCommand cmd7 = new SqlCommand(
                                    "select COUNT(*) AS conteo from [No_Conformidades] where  " +
                                    "  (Status ='Abierta')"

                                    , conexion);



                SqlCommand cmd8 = new SqlCommand(
                                    "select COUNT(*) AS conteo from [No_Conformidades] where (Status ='Cerrada')"
                                    , conexion);


                SqlCommand cmd9 = new SqlCommand(
                                    "select COUNT(*) AS conteo from [No_Conformidades] where   (Status ='Verificando')"
                                    , conexion);





                Int32 rows_count = Convert.ToInt32(cmd.ExecuteScalar());
                Int32 rows_count1 = Convert.ToInt32(cmd1.ExecuteScalar());
                Int32 rows_count2 = Convert.ToInt32(cmd3.ExecuteScalar());
                Int32 rows_count3 = Convert.ToInt32(cmd5.ExecuteScalar());
                Int32 rows_count4 = Convert.ToInt32(cmd6.ExecuteScalar());
                Int32 rows_count5 = Convert.ToInt32(cmd7.ExecuteScalar());
                Int32 rows_count6 = Convert.ToInt32(cmd8.ExecuteScalar());
                Int32 rows_count7 = Convert.ToInt32(cmd9.ExecuteScalar());

                NumeroConfo = Convert.ToInt32(cmd2.ExecuteScalar());
                NumeroOport = Convert.ToInt32(cmd4.ExecuteScalar());

                cmd.Dispose();
                cmd1.Dispose();
                cmd2.Dispose();
                cmd3.Dispose();
                cmd4.Dispose();
                cmd5.Dispose();
                cmd6.Dispose();
                cmd7.Dispose();
                cmd8.Dispose();
                cmd9.Dispose();

                conexion.Close();

                NumeroConformidad = NumeroConfo.ToString();
                Nc_Abiertas.Text = rows_count.ToString();
                NC_Concluidas.Text = rows_count1.ToString();
                NC_Verificacion.Text = rows_count2.ToString();
                OP_Abiertas.Text = rows_count3.ToString();
                Op_Cerradas.Text = rows_count4.ToString();
                Ncag.Text = rows_count5.ToString();
                Ncvg.Text = rows_count7.ToString();
                Nccg.Text = rows_count6.ToString();
                Nom.Text = Nombre;
                Ape.Text = Apellido;
            }
            catch (Exception d)
            {
                MessageBox.Show(d.Message);
            }
            finally
            {

            }

            Depart.Text = Departamento;
            Pues.Text = Puesto;
            correo.Text = Email;




        }

        private void consultaToolStripMenuItem5_Click_1(object sender, EventArgs e)
        {
            if (conexionvpn == 1)
            {
                conexionvpn = 0;
            }
            else
            {
                conexionvpn = 2;
                if (Departamento == "Administracion" || Departamento == "SuperAdmin" || Departamento == "Direccion" || Departamento == "Compras")
                {

                    if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                    {

                        if (!Central.Controls.Contains(Consulta_No_Conf.Instance))
                        {
                            Ventana = "Pantalla: Consulta de No Conformidad";
                            Consulta_No_Conf.Instance.Departamento = Departamento;
                            Central.Controls.Add(Consulta_No_Conf.Instance);
                            Consulta_No_Conf.Instance.Dock = DockStyle.Fill;
                            Consulta_No_Conf.Instance.BringToFront();


                        }
                        else
                        {
                            Consulta_No_Conf.Instance.BringToFront();
                            Ventana = "Pantalla: Consulta de No Conformidad";
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else
                {
                    MessageBox.Show("No tiene acceso a este modulo");
                }
            }
            if (conexionvpn == 0)
            {
                DialogResult Resultado;
                Resultado = MessageBox.Show("Se desconectara del navegador ¿desea salir? ", "Confirmación", MessageBoxButtons.YesNo);
                if (Resultado == DialogResult.Yes)
                {
                    NetworkDrive oNetDrive = new NetworkDrive();

                    try
                    {
                        NetworkDrive oNetDrive1 = new NetworkDrive();
                        try
                        {
                            //set propertys
                            oNetDrive1.Force = true; oNetDrive1.LocalDrive = "G";
                            oNetDrive1.ShareName = url.Text;
                            //match call to options provided
                            oNetDrive1.MapDrive();
                            conexionvpn = 1;
                            //update status
                        }
                        catch (Exception err)
                        {
                            //report error

                          //  MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                            conexionvpn = 0;
                        }
                        oNetDrive1 = null;

                        oNetDrive.LocalDrive = "G";
                        oNetDrive.UnMapDrive();
                        File.WriteAllText(FolderPath + "\\VpnDisconnect.bat", "rasdial /d");
                        var newProcess = new Process
                        { StartInfo = { FileName = FolderPath + "\\VpnDisconnect.bat", WindowStyle = ProcessWindowStyle.Normal } };
                        newProcess.Start();
                        newProcess.WaitForExit();
                        Form fc = Application.OpenForms["Form2"];
                        if (fc != null) fc.Close();

                    }
                    catch (Exception err) { }
                    //   { // MessageBox.Show(this, "Cannot unmap drive!\nError: " + err.Message);  //   }
                    oNetDrive = null;
                    if (Departamento == "Compras" || Departamento == "SuperAdmin" || Departamento == "Direccion")
                    {

                        if (Permiso == "1" || Permiso == "2" || Permiso == "3" || Permiso == "4" || Permiso == "5")
                        {

                            if (!Central.Controls.Contains(Consulta_No_Conf.Instance))
                            {
                                Ventana = "Pantalla: Consulta de No Conformidad";
                                Consulta_No_Conf.Instance.Departamento = Departamento;
                                Central.Controls.Add(Consulta_No_Conf.Instance);
                                Consulta_No_Conf.Instance.Dock = DockStyle.Fill;
                                Consulta_No_Conf.Instance.BringToFront();


                            }
                            else
                            {
                                Consulta_No_Conf.Instance.BringToFront();
                                Ventana = "Pantalla: Consulta de No Conformidad";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No tiene acceso a este modulo");
                        }
                    }
                    else
                    {
                        MessageBox.Show("No tiene acceso a este modulo");
                    }
                }
                else if (Resultado == DialogResult.No)
                {
                    conexionvpn = 1;
                    MessageBox.Show("Operacion Cancelada");
                }

            }
        }

        private void consultaProyectosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CBR_ADMIN.Sistema.Folio_Proyectos crear1 = new CBR_ADMIN.Sistema.Folio_Proyectos();

            crear1.Show();
        }

        private void crearNuevaPropuestaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
                Crea_Propuesta crear2 = new Crea_Propuesta();

                crear2.Show();
            }

        private void crearNuevaRehabilitacionToolStripMenuItem_Click(object sender, EventArgs e)
        {

            Crea_Rehabilitacion crear2 = new Crea_Rehabilitacion();

            crear2.Show();
        }

        private void crearNuevaPropuestaToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Crea_Propuesta crear2 = new Crea_Propuesta();

            crear2.Show();
        }

        private void crearNuevaRehabilitacionToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Crea_Rehabilitacion crear2 = new Crea_Rehabilitacion();

            crear2.Show();
        }

        private void comprasToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            MsgBoxUtil.HackMessageBox("Local ", "Remota");

            DialogResult Resultado;
            Resultado = MessageBox.Show("Seleccione el tipo de conexión", "Navegador de proyectos", MessageBoxButtons.YesNo);

            string strCmdText;
            strCmdText = "/C attrib +h Y:\\pruebas";
            Process cop = System.Diagnostics.Process.Start("CMD.exe", strCmdText);
            cop.WaitForExit();
            if (Resultado == DialogResult.Yes)
            {
                NetworkDrive oNetDrive = new NetworkDrive();

                try
                {
                    //set propertys

                    oNetDrive.LocalDrive = "G";
                    oNetDrive.ShareName = url.Text;
                    //match call to options provided

                    oNetDrive.MapDrive();
                    conexionvpn = 1;

                    //update status

                }
                catch (Exception err)
                {
                    //report error

                    //    MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
                }
                oNetDrive = null;

            }
            else if (Resultado == DialogResult.No)
            {


                Form2 conecta = new Form2();

                string envio = "2";
                conecta.BringToFront();
                conecta.WindowState = FormWindowState.Normal;
                conecta.Close();

                Form2 conecta2 = new Form2();

                conecta2.recep = envio;
                conecta2.Show();
                conecta2.BringToFront();
                conecta2.WindowState = FormWindowState.Normal;

                NetworkDrive oNetDrive = new NetworkDrive();
                conexionvpn = 1;

            }



            if (!Central.Controls.Contains(windowsNavegador.Instance))
            {
                Ventana = "Pantalla: Consulta de No Conformidad";



                windowsNavegador.Instance.usuario = Nombre;
                windowsNavegador.Instance.emailusuario = Email;
                windowsNavegador.Instance.departamento = Departamento;
                windowsNavegador.Instance.apellido = Apellido;
                Central.Controls.Add(windowsNavegador.Instance);
                windowsNavegador.Instance.Dock = DockStyle.Fill;
                windowsNavegador.Instance.BringToFront();
            }
            else
            {
                windowsNavegador.Instance.BringToFront();
                Ventana = "Pantalla: Consulta de No Conformidad";
            }
        }

        private void asistenteDeDireccionToolStripMenuItem1_Click(object sender, EventArgs e)
        {
 

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
         
        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {

            CBR_ADMIN.Administracion.Avance genera = new CBR_ADMIN.Administracion.Avance();
            genera.Show();


        }

        private void pictureBox6_Click(object sender, EventArgs e)
        {
            Actualiza_porcentaje genera = new Actualiza_porcentaje();
            genera.Show();
        }

        private void generarValidadorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ValidadorProyecto crear = new ValidadorProyecto();

            crear.Show();
        }

        private void toolStripMenuItem171_Click(object sender, EventArgs e)
        {

            NetworkDrive oNetDrive = new NetworkDrive();

            try
            {
                //set propertys

                oNetDrive.LocalDrive = "G";
                oNetDrive.ShareName = url.Text;
                //match call to options provided

                oNetDrive.MapDrive();
                conexionvpn = 1;

                //update status

            }
            catch (Exception err)
            {
                //report error

           //     MessageBox.Show(this, "Cannot map drive!\nError: " + err.Message);
            }
            oNetDrive = null;



            if (!Central2.Controls.Contains(windowsNavegador.Instance))
            {
                Ventana = "Pantalla: Consulta de No Conformidad";

                
             
                windowsNavegador.Instance.usuario = Nombre;
                windowsNavegador.Instance.emailusuario = Email;
                windowsNavegador.Instance.apellido = Apellido;
                windowsNavegador.Instance.departamento = Departamento;
                Central2.Controls.Add(windowsNavegador.Instance);
                windowsNavegador.Instance.Dock = DockStyle.Fill;
                windowsNavegador.Instance.BringToFront();
            }
            else
            {
                windowsNavegador.Instance.BringToFront();
                Ventana = "Pantalla: Consulta de No Conformidad";
            }
        }

        private void Central5_Paint(object sender, PaintEventArgs e)
        {

        }

       

     

        private void btnminimizar_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
           
        }

        private void General_SizeChanged(object sender, EventArgs e)
        {
         if(this.WindowState == FormWindowState.Minimized)
            {
              //  notifyIcon1.Icon =
            }
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void maximizarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Normal;
        }

        private void notifyIcon1_BalloonTipClicked(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Normal;
        }



        private void recepcionToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            try
            {
                ////////////////////////////////////SE ABRE LA CONEXION/////////////////////////////////////////////////////////////////////////////////
                SqlConnection conexion = new SqlConnection(ObtenerCadena());
                conexion.Open();
                ////////////////////////////////////SE EJECUTA EL COMANDO HACIA LA VISTA Y SE HACE LA VALIDACION//////////////////////////////////////
                SqlCommand cmd = new SqlCommand(
                                                "select COUNT(*) AS conteo from [No_Conformidades] where  " +
                                                " (Status ='Abierta')" +
                                                "and (Departamento = @Departamento)"
                                                , conexion);
                cmd.Parameters.AddWithValue("Responsable", Nombre);
                cmd.Parameters.AddWithValue("Departamento", Departamento);
                SqlCommand cmd1 = new SqlCommand(
                                    "    select COUNT(*) AS conteo from[No_Conformidades] where(Departamento = @Departamento)  and(Status = 'Cerrada') or(Status = 'Verificando')"
                                    , conexion);

                cmd1.Parameters.AddWithValue("Responsable", Nombre);
                cmd1.Parameters.AddWithValue("Departamento", Departamento);



                SqlCommand cmd2 = new SqlCommand(
                                  "select COUNT([# NC]) AS conteo from [No_Conformidades] "
                                  , conexion);
                SqlCommand cmd3 = new SqlCommand(
                              "select COUNT(*) AS conteo from [No_Conformidades] where " +
                              "  (Status ='Verificando')" +
                              "and (Departamento = @Departamento)"
                              , conexion);
             
                cmd3.Parameters.AddWithValue("Departamento", Departamento);
                SqlCommand cmd4 = new SqlCommand(
                              "select COUNT(*) AS conteo from [Op_Mejora]  "
                              , conexion);
              

                SqlCommand cmd5 = new SqlCommand(
                                                "select COUNT(*) AS conteo from [Op_Mejora] where  " +
                                                "  (Status ='Abierta')" +
                                                "and (Departamento = @Departamento)"
                                                , conexion);

                cmd5.Parameters.AddWithValue("Departamento", Departamento);
                SqlCommand cmd6 = new SqlCommand(
                                    "select COUNT(*) AS conteo from [Op_Mejora] where  " +
                                    "  (Status ='Cerrada') " +
                                    "and (Departamento = @Departamento)"
                                    , conexion);

                cmd6.Parameters.AddWithValue("Departamento", Departamento);
                cmd3.Parameters.AddWithValue("Responsable", Nombre);
                

                Int32 rows_count = Convert.ToInt32(cmd.ExecuteScalar());
                Int32 rows_count1 = Convert.ToInt32(cmd1.ExecuteScalar());
                Int32 rows_count2 = Convert.ToInt32(cmd3.ExecuteScalar());
                Int32 rows_count3 = Convert.ToInt32(cmd5.ExecuteScalar());
                Int32 rows_count4 = Convert.ToInt32(cmd6.ExecuteScalar());
                NumeroOport = Convert.ToInt32(cmd3.ExecuteScalar());
                NumeroConfo = Convert.ToInt32(cmd4.ExecuteScalar());
                cmd.Dispose();
                cmd1.Dispose();
                cmd2.Dispose();
                cmd3.Dispose();
                cmd4.Dispose();
                conexion.Close();
                NumeroConformidad = NumeroConfo.ToString();
                Nc_Abiertas.Text = rows_count.ToString();
                NC_Concluidas.Text = rows_count1.ToString();
                NC_Verificacion.Text = rows_count2.ToString();
                OP_Abiertas.Text = rows_count3.ToString();
                Op_Cerradas.Text = rows_count4.ToString();
            }
            catch (Exception d)
            {
                MessageBox.Show(d.Message);
            }
            finally
            {

            }
        }
    }





    }
    
