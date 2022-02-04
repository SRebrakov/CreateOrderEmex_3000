using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.ServiceModel;
using System.ServiceModel.Channels;
using System.Text.RegularExpressions;
using System.Net;
using System.Threading;
using System.Data.SqlClient;
using System.Configuration;
using System.IO;
using NLog;
using System.Web;

using System.Net.Mail;
using System.Net.Mime;



namespace CreateOrderEmex
{

  
    public partial class MainForm : Form
    {
        public EmexReference.Customer customer = new EmexReference.Customer();
        public EmexReference.ServiceSoapClient soap = new EmexReference.ServiceSoapClient();

        static string MegaLogin = string.Empty;
        static string MegaPassword = string.Empty;
        static string ticketMega;
        static string SupplierLogo = string.Empty;
        SqlConnection myConnection;
        private static Logger logger ;
        bool bwork;
        static string Mode = string.Empty;
        static string headid = string.Empty;
               
        public MainForm()
        {
            InitializeComponent();

            FileInfo fi = new FileInfo(Application.ExecutablePath);
            string iniFile = string.Empty;
            logger = LogManager.GetCurrentClassLogger();

            iniFile = fi.DirectoryName +"\\"+ fi.Name.Substring(0,fi.Name.Length-4)+".ini"; 
            
            Ini.IniFile ifile = new Ini.IniFile(iniFile);

            FileInfo ss = new FileInfo(iniFile);

            if (ss.Exists)
            {
                
                
                tbLogin.Text = ifile.IniReadValue("Web Service", "Login");
                tbPassword.Text = ifile.IniReadValue("Web Service", "Password");

                tbServer.Text = ifile.IniReadValue("SQL Server", "Server");
                tbDatabase.Text = ifile.IniReadValue("SQL Server", "Database");
                tbsLogin.Text = ifile.IniReadValue("SQL Server", "Login");
                tbsPassowrd.Text = ifile.IniReadValue("SQL Server", "Password");
                SupplierLogo = ifile.IniReadValue("DopInfo", "SupplierLogo");
                Mode = ifile.IniReadValue("DopInfo", "Mode");

                if (SupplierLogo.Length == 0) {
                    logger.Log(LogLevel.Debug, "Unknown SupplierLogo");
                    MessageBox.Show("Unknown SupplierLogo");

                }
            }
            else {
              logger.Log(LogLevel.Debug, "No find ini file (" + iniFile + ")");
              MessageBox.Show("No find ini file ("+iniFile+")");
            }
            
            MegaLogin = tbLogin.Text;
            MegaPassword = tbPassword.Text;
        }

        
        private void CreateOrder()
        {
            bwork = true;
            bStart.Enabled = false;
            bSettings.Enabled = false;
            bExit.Enabled = false;
           
            int nstr = 0;
            EmexReference.BasketDetails[] BI;
            EmexReference.Balance BL = new EmexReference.Balance();
            EmexReference.partstobasket[] arr;
            //  EmexReference.CreateOrder_Result CreateOrder = new EmexReference.CreateOrder_Result();
            EmexReference.ListOfOrders OrderInfo = new EmexReference.ListOfOrders();
            System.Data.SqlClient.SqlCommand sqlcom;
            System.Data.SqlClient.SqlCommand sqlcomplete;
            System.Data.SqlClient.SqlCommand sqlproc;
            System.Data.SqlClient.SqlDataReader sqldr;
            System.Data.SqlClient.SqlDataReader drcomplete;
            DataTable SupplierHeder = new DataTable();
            DataTable Details = new DataTable();
            int IsAccepted =0;
            int kol = 0;
            int DetId = 0;
            string temstr = string.Empty;
            bool nextdetail = false;
            bool neadorder = false;
            int ResultCode=0;
            string ResultText = string.Empty;

            textBox1.Clear();
            logger.Log(LogLevel.Debug, "Start");
            textBox1.Text = "Start";
            sqlcomplete = new SqlCommand();
         
            while (bwork == true)
            {
  
                try
                {
                    String string_con = "Data Source=" + tbServer.Text + ";Initial Catalog=" + tbDatabase.Text + ";User ID=" + tbsLogin.Text + ";Password=" + tbsPassowrd.Text;
                    myConnection = new SqlConnection(string_con);
                    sqlcom = new SqlCommand();
                    sqlcomplete = new SqlCommand();
                    sqlcom.Connection = myConnection;
                    sqlcomplete.Connection = myConnection;
               
                    myConnection.Open();
 
                    sqlcom.CommandText = "exec dbo.p_SupplierOrder_Export_GetHeads @SupplierLogo = '" + SupplierLogo + "'";
                    sqldr = sqlcom.ExecuteReader();
                    SupplierHeder.Load(sqldr);

                    for (int i = 0; i < SupplierHeder.Rows.Count; i++)
                    {
                       
                        customer.UserName = "QXXX";
                        customer.Password = "qXQx";
                        customer.CustomerId = null;
                        customer.SubCustomerId = null;

                        customer = soap.Login(customer);

                        BI=soap.GetBasketDetails(customer);

                        soap.DeleteFromBasket(customer, BI);

                        myConnection.Close();
                        myConnection.Open();


                        headid = SupplierHeder.Rows[i][0].ToString();
                        sqlcom.CommandText = "exec dbo.p_SupplierOrder_Export_GetDetails @HeadId=" + SupplierHeder.Rows[i][0].ToString();
                        sqldr = sqlcom.ExecuteReader();
                        Details.Clear();
                        Details.Load(sqldr);

                        arr = new EmexReference.partstobasket[Details.Rows.Count];
                        kol = Details.Rows.Count;

                       

                        if (kol == 0) { continue; }

                        for (int k = 0; k < kol; k++)
                        {
                            arr[k] = new EmexReference.partstobasket();
                          //arr[k].Comments = Details.Rows[k]["Comment"].ToString(); //коментарий
                            arr[k].DetailNum = Details.Rows[k]["Articul"].ToString();  // номер детали
                            arr[k].Quantity = Convert.ToInt32(Details.Rows[k]["Quantity"]);  // кол-во
                            arr[k].MakeLogo = Details.Rows[k]["Brand"].ToString();
                            arr[k].CoeffMaxAgree = Convert.ToDouble(Details.Rows[k]["CoeffPriceAgree"].ToString());                                
                            arr[k].PriceLogo = Details.Rows[k]["PriceLogo"].ToString();
                            arr[k].CustomerSubId = Convert.ToInt64(Details.Rows[k]["ReferenceId"]);
                            arr[k].UploadedPrice = Convert.ToDouble(Details.Rows[k]["Price"]);
                            arr[k].Reference = Details.Rows[k]["InternalOrderDetailId"].ToString();
                            arr[k].DestinationLogo = Details.Rows[k]["DestinationLogo"].ToString();
                            arr[k].CustomerStickerData = Details.Rows[k]["StickerData"].ToString();
                         //  arr[k].bitONLY = Convert.ToBoolean(Details.Rows[k]["bitONLY"]);
                         //  arr[k].bitBrand = Convert.ToBoolean(Details.Rows[k]["bitBrand"]);

                        }
                        soap.InsertPartToBasket(customer, arr);
                        //   return;

                        // возможно понадобиться очистка BI
                        BI = soap.GetBasketDetails(customer);

                        // InvokeMegaService(() => { StoreApi.AllToOrderBasket(); });
                        for (int j = 0; j < BI.Count(); j++)
                        {
                            nextdetail = true;
                            
                          
                            if (nextdetail == false) { continue; }
                               
                                IsAccepted = 0;
                                sqlproc = new SqlCommand();
                                sqlproc.Connection = myConnection;
                                sqlproc.CommandType = CommandType.StoredProcedure;
                                sqlproc.CommandText = "dbo.p_SupplierOrder_ExportRec";
                                sqlproc.Parameters.AddWithValue("@InternalOrderDetailId", BI[j].CustomerSubId);
                                sqlproc.Parameters.AddWithValue("@Articul", BI[j].DetailNum);
                                sqlproc.Parameters.AddWithValue("@Brand", BI[j].MakeLogo);
                                sqlproc.Parameters.AddWithValue("@Quantity", BI[j].Quantity);
                                sqlproc.Parameters.AddWithValue("@Price", BI[j].Price);
                                sqlproc.Parameters.AddWithValue("@Comment", BI[j].Comments);
                                sqlproc.Parameters.AddWithValue("@IsAccepted", IsAccepted);
                                sqlproc.Parameters[6].Direction = ParameterDirection.Output;
                                sqldr = sqlproc.ExecuteReader();
                                IsAccepted = Convert.ToInt32(sqlproc.Parameters[6].Value.ToString());
                                sqldr.Close();
                                if (IsAccepted == 0)
                                {
                                    BI[j].bitConfirm = false;
                                    nstr++;
                                    temstr = nstr + "   " + BI[j].Price.ToString() + "    " + Details.Rows[DetId]["Price"].ToString() + "      " + Details.Rows[DetId]["CoeffPriceAgree"].ToString();
                                    textBox1.Text = textBox1.Text + "\r\n" + temstr;
                                    logger.Log(LogLevel.Debug, temstr);
                                    Application.DoEvents();
                                }
                                else
                                {
                                    BI[j].bitConfirm = true;
                                }
                        }

                        soap.UpdateBasketDetails(customer, BI);
                        neadorder = false;
                        for (int j = 0; j < BI.Count(); j++)
                        {
                            if (BI[j].bitConfirm == true) {
                                neadorder = true;
                                
                                break;
                            }
                        }

                        if (neadorder == false)
                        {
                            ResultText = "UpdateBasket not confirm details";
                            ResultCode = -10000;
                        }

                        
                        if (neadorder == true)
                        {

                            int CreateOrderRes= soap.CreateOrder(customer);
                           
                             /*                           
                            if (CreateOrder.OrderNumber.ToString().Length == 0) {
                                CreateOrder.OrderNumber = 0;
                            }
                             */
                        
                            if (CreateOrderRes != 0)
                            {
                                textBox1.Text = textBox1.Text + "\r\n Create order ok";
                                logger.Log(LogLevel.Debug, "Create order ok");
                                ResultText = "Create order ok";
                                ResultCode = 0;

                            }
                            else
                            {
                                ResultText = "Error CreateOrder";
                                ResultCode = CreateOrderRes;
                            }
                        }
                        else {
                             ResultText = "needorder = false.StoreApi.UpdateBasket not confirm details";
                             ResultCode = -4;
                        }

                        logger.Log(LogLevel.Debug, ResultText + ". ResultCode=" + ResultCode.ToString());
                        sqlcomplete.CommandText = "exec dbo.p_SupplierOrder_Export_Complete @HeadId=" + SupplierHeder.Rows[i][0].ToString() + ",@ResultCode = " + ResultCode.ToString() +",@ResultText='" + ResultText + "'";
                       // drcomplete = sqlcomplete.ExecuteReader();
                       // drcomplete.Close();

                        BI = soap.GetBasketDetails(customer);
                        soap.DeleteFromBasket(customer, BI);
                       
                    }
                    
                    myConnection.Close();
                }
                catch (SqlException exp)
                {
                  //  bwork = false;
                    logger.Log(LogLevel.Debug, exp.Message.ToString());
                    textBox1.Text = textBox1.Text + "\r\n Ошибка!!! см. логи.";
                }
                catch (System.ServiceModel.FaultException exp)
                {
                    string err;
                    err = ((System.ServiceModel.FaultException<System.ServiceModel.ExceptionDetail>)exp).Detail.InnerException.Message.ToString();
                    logger.Log(LogLevel.Debug, err);
                    textBox1.Text = textBox1.Text + "\r\n Ошибка!!! \r\n"+ err;

                    sqlcomplete.CommandText = "exec dbo.p_SupplierOrder_Export_Complete @HeadId=" + headid + ",@ResultCode = 1 ,@ResultText='Exception CreateOrder'";
                  //  drcomplete = sqlcomplete.ExecuteReader();
                  // drcomplete.Close();
                    myConnection.Close();
                 //   SendMail("Error CreateOrderEmex", err);
                }
                catch (Exception exp)
                {
                    //bwork = false; 
                    logger.Log(LogLevel.Debug, exp.Message.ToString());
                    textBox1.Text = textBox1.Text + "\r\n Ошибка!!! см. логи.";
                    try
                    {
                        if (headid.Length > 0)
                        {
                            sqlcomplete.CommandText = "exec dbo.p_SupplierOrder_Export_Complete @HeadId=" + headid + ",@ResultCode = 1 ,@ResultText='Exception CreateOrder'";
                        //    drcomplete = sqlcomplete.ExecuteReader();
                        //    drcomplete.Close();
                            myConnection.Close();
                        }
                    }
                    catch (Exception exp1)
                    {
                     logger.Log(LogLevel.Debug, exp1.Message.ToString());
                     textBox1.Text = textBox1.Text + "\r\n Ошибка!!! см. логи.";
                    }
                }
                if (Mode == "auto")
                {
                    logger.Log(LogLevel.Debug, "Finish");
                    this.Close();
                    return;
                }

                for (int k = 0; k < 10000; k++)
                {
                    Application.DoEvents();
                    if (bwork == false) break;
                     Thread.Sleep(60);
                }
                if (bwork == false) break;
            }
            logger.Log(LogLevel.Debug, "Finish");
            textBox1.Text = textBox1.Text + "\r\n Finish";
            bStart.Enabled = true;
            bExit.Enabled = true;
          
        }
       

      
      

        private void bConnect_Click(object sender, EventArgs e)
        {
            //BuildConnectionString(tbDatabase.Text, tbsLogin.Text, tbsPassowrd.Text);
        }

        private void bExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void bStop_Click(object sender, EventArgs e)
        {
            bwork = false;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            textBox1.SelectionStart = textBox1.Text.Length;
            textBox1.ScrollToCaret();
            textBox1.Refresh();
        }

        private void SendMail(string sSubject, string sBody) {


            try
            {
                SmtpClient Smtp = new SmtpClient("smtp.gmail.com", 587);
                Smtp.Credentials = new NetworkCredential("rebrakovsw", "gecnbvtyz");
                Smtp.EnableSsl = true;
                //Smtp.EnableSsl = false;

                //Формирование письма
                MailMessage Message = new MailMessage();
                Message.From = new MailAddress("rebrakovsw@gmail.com");
                Message.To.Add(new MailAddress("rebrakovsw@gmail.com"));
                Message.Subject = sSubject;
                Message.Body = sBody;

                //Прикрепляем файл
                /*
                string file = "C:\\file.zip";
                Attachment attach = new Attachment(file, MediaTypeNames.Application.Octet);

                // Добавляем информацию для файла
                ContentDisposition disposition = attach.ContentDisposition;
                disposition.CreationDate = System.IO.File.GetCreationTime(file);
                disposition.ModificationDate = System.IO.File.GetLastWriteTime(file);
                disposition.ReadDate = System.IO.File.GetLastAccessTime(file);

                Message.Attachments.Add(attach);
                */
                Smtp.Send(Message);//отправка
            }
            catch (Exception exp) {
                logger.Log(LogLevel.Debug, exp);
            
            }
        
        }

        private void bStart_Click(object sender, EventArgs e)
        {
               
            CreateOrder();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            if (Mode == "auto")
            {
                CreateOrder();
            }
        }

       

    }
}
