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
        static string MegaLogin = string.Empty;
        static string MegaPassword = string.Empty;
        static ServiceReference1.AuthenticationServiceClient authClient;
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
            ServiceReference2.BasketItem[] BI;
            ServiceReference2.BalanceInfo BL = new ServiceReference2.BalanceInfo();
            ServiceReference2.InsertToBasketItem[] arr;
            ServiceReference2.CreateOrder_Result CreateOrder = new ServiceReference2.CreateOrder_Result();
            ServiceReference2.OrderInfo OrderInfo = new ServiceReference2.OrderInfo();
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
          /*  int ord = 0;
            int kk = 50;
                        while (true) {
                            ord = ord + 1;
                            authClient = new ServiceReference1.AuthenticationServiceClient();

                            if (authClient.State == CommunicationState.Closed)
                                authClient.Open();

                            ticketMega = AuthenticateMega();

                             InvokeMegaService(() => { StoreApi.EmptyBasket(); });

                             arr = new ServiceReference2.InsertToBasketItem[kk];
                             for (int j = 0; j <kk; j++)
                             {
                                 arr[j] = new ServiceReference2.InsertToBasketItem();
                                 arr[j].Comments = "Coment test"; //коментарий
                                 arr[j].DetailNum = "38913AA101"; //Details.Rows[k]["Articul"].ToString();  // номер детали
                                 arr[j].Quantity = 1;  // кол-во
                                 arr[j].MakeLogo = "SU";///Details.Rows[k]["Brand"].ToString();
                                 arr[j].CoeffMaxAgree = 1;
                                 arr[0].PriceLogo = "EMIR";
                                 //  arr[k].CustomerSubId = Convert.ToInt64(Details.Rows[k]["InternalOrderDetailId"]);
                                 // arr[0].CustomerSubId = 0;
                                 // arr[0].UploadedPrice = 0;
                                 arr[j].Reference = "";
                                 arr[j].DestinationLogo = "EMEW";
                                 arr[j].CustomerStickerData = "";
                             }
                             InvokeMegaService(() => { StoreApi.InsertBasketItemsToBasket(arr); });

                             BI = new ServiceReference2.BasketItem[kk];
                             InvokeMegaService(() => { BI = StoreApi.GetBasket(); });
                             for (int j = 0; j < kk; j++)
                             {
                                 BI[j].bitConfirm = true;

                                 InvokeMegaService(() =>
                                 {
                                     StoreApi.UpdateBasket(
                                         BI[j].BasketItemId
                                         , BI[j].Quantity
                                         , BI[j].DestinationLogo
                                         , (bool)BI[j].bitONLY
                                         , (bool)BI[j].bitAGREE
                                         , (bool)BI[j].bitWait
                                         , (bool)BI[j].bitConfirm
                                         , (bool)BI[j].bitBrand
                                         , BI[j].Reference
                                         , (long)BI[j].CustomerSubId
                                         , BI[j].Comments
                                         , (decimal)BI[j].CoeffMaxAgree
                                         , (decimal)BI[j].UploadedPrice);
                                 });
                             }
                            InvokeMegaService(() => { CreateOrder = StoreApi.CreateOrder(); });
                            logger.Log(LogLevel.Debug, ord.ToString() +" "+CreateOrder.OrderNumber);
                          //   logger.Log(LogLevel.Debug, ord.ToString());
                            InvokeMegaService(() => { StoreApi.EmptyBasket(); });
                           // Thread.Sleep(100);

                            authClient.Logout();
                            authClient.Close();
                            

            
                        }*/
                        
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

                        authClient = new ServiceReference1.AuthenticationServiceClient();

                        if (authClient.State == CommunicationState.Closed)
                          authClient.Open();

                        ticketMega = AuthenticateMega();

                        InvokeMegaService(() => { StoreApi.EmptyBasket(); });

                      //  authClient.Open();

                        myConnection.Close();
                        myConnection.Open();


                        headid = SupplierHeder.Rows[i][0].ToString();
                        sqlcom.CommandText = "exec dbo.p_SupplierOrder_Export_GetDetails @HeadId=" + SupplierHeder.Rows[i][0].ToString();
                        sqldr = sqlcom.ExecuteReader();
                        Details.Clear();
                        Details.Load(sqldr);

                        arr = new ServiceReference2.InsertToBasketItem[Details.Rows.Count];
                        kol = Details.Rows.Count;

                        if (kol == 0) { continue; }

                        for (int k = 0; k < kol; k++)
                        {
                            arr[k] = new ServiceReference2.InsertToBasketItem();
                            arr[k].Comments = Details.Rows[k]["Comment"].ToString(); //коментарий
                            arr[k].DetailNum = Details.Rows[k]["Articul"].ToString();  // номер детали
                            arr[k].Quantity = Convert.ToInt32(Details.Rows[k]["Quantity"]);  // кол-во
                            arr[k].MakeLogo = Details.Rows[k]["Brand"].ToString();
                            arr[k].CoeffMaxAgree = Convert.ToDecimal(Details.Rows[k]["CoeffPriceAgree"].ToString());                                
                            arr[k].PriceLogo = Details.Rows[k]["PriceLogo"].ToString();
                          //  arr[k].CustomerSubId = Convert.ToInt64(Details.Rows[k]["InternalOrderDetailId"]);
                            arr[k].CustomerSubId = Convert.ToInt64(Details.Rows[k]["ReferenceId"]);
                            arr[k].UploadedPrice = Convert.ToDecimal(Details.Rows[k]["Price"]);
                            arr[k].Reference = Details.Rows[k]["InternalOrderDetailId"].ToString();
                            arr[k].DestinationLogo = Details.Rows[k]["DestinationLogo"].ToString();
                            arr[k].CustomerStickerData = Details.Rows[k]["StickerData"].ToString();
                            arr[k].bitONLY = Convert.ToBoolean(Details.Rows[k]["bitONLY"]);
                            arr[k].bitBrand = Convert.ToBoolean(Details.Rows[k]["bitBrand"]);

                        }
                        InvokeMegaService(() => { StoreApi.InsertBasketItemsToBasket(arr); });
                     //   return;
                        BI = new ServiceReference2.BasketItem[Details.Rows.Count];
                        InvokeMegaService(() => { BI = StoreApi.GetBasket(); });

                        // InvokeMegaService(() => { StoreApi.AllToOrderBasket(); });
                        for (int j = 0; j < BI.Count(); j++)
                        {
                            nextdetail = true;
                            
                          /* не помню для чего 27.12.2016
                            for (DetId = 0; DetId < kol + 1; DetId++)
                            {
                                if (DetId == kol)
                                {
                                    logger.Log(LogLevel.Debug, "Неизвестная деталь! DetailId = " + DetId.ToString());
                                    nextdetail = false;
                                    break;
                                }

                                if (BI[j].Reference == Details.Rows[DetId]["InternalOrderDetailId"].ToString())
                                    break;

                            }
                           */
                            if (nextdetail == false) { continue; }

                        /*    if (BI[j].Price > ((Convert.ToDecimal(Details.Rows[DetId]["Price"])) + (Convert.ToDecimal(Details.Rows[DetId]["CoeffPriceAgree"]))))
                            {
                            */
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

                           /* }
                            else
                            {
                                BI[j].bitConfirm = true;
                            }
                            */
                            if (BI[j].bitConfirm == true)
                            {
                               InvokeMegaService(() =>
                               {
                                   StoreApi.UpdateBasket(
                                         BI[j].BasketItemId
                                       , BI[j].Quantity
                                       , BI[j].DestinationLogo
                                       , (bool)BI[j].bitONLY
                                       , (bool)BI[j].bitAGREE
                                       , (bool)BI[j].bitWait
                                       , (bool)BI[j].bitConfirm
                                       , (bool)BI[j].bitBrand
                                       , BI[j].Reference
                                       , (long)BI[j].CustomerSubId
                                       , BI[j].Comments
                                       , (decimal)BI[j].CoeffMaxAgree
                                       , (decimal)BI[j].UploadedPrice);
                               });
                               temstr = "UpdateBasket:"
                                        + " BasketItemId=" +BI[j].BasketItemId.ToString()
                                        + " DetailNum=" + BI[j].DetailNum.ToString()
                                        + " MakeLogo=" + BI[j].MakeLogo.ToString()
                                        + " Quantity=" + BI[j].Quantity.ToString()
                                        + " DestinationLogo=" + BI[j].DestinationLogo.ToString()
                                        + " bitONLY=" + BI[j].bitONLY.ToString()
                                        + " bitAGREE=" + BI[j].bitAGREE.ToString()
                                        + " bitWait=" + BI[j].bitWait.ToString()
                                        + " bitConfirm=" + BI[j].bitConfirm.ToString()
                                        + " bitBrand=" + BI[j].bitBrand.ToString()
                                        + " Reference=" + BI[j].Reference.ToString()
                                        + " CustomerSubId=" + BI[j].CustomerSubId.ToString()
                                        + " Comments=" + BI[j].Comments.ToString()
                                        + " CoeffMaxAgree=" + BI[j].CoeffMaxAgree.ToString()
                                        + " UploadedPrice=" + BI[j].UploadedPrice.ToString();

                               textBox1.Text = textBox1.Text + "\r\n" + temstr;
                               logger.Log(LogLevel.Debug, temstr);
                               Application.DoEvents();

                            }


                        
                        }
                        neadorder = true;
                        for (int j = 0; j < BI.Count(); j++)
                        {
                            if (BI[j].bitConfirm == false) {
                                neadorder = false;
                                ResultText = "UpdateBasket not confirm details";
                                ResultCode = -10000;
                                break;
                            }
                        }

                        
                        if (neadorder == true)
                        {
                             
                            InvokeMegaService(() => { CreateOrder = StoreApi.CreateOrder(); });
                                                        
                            if (CreateOrder.OrderNumber.ToString().Length == 0) {
                                CreateOrder.OrderNumber = 0;
                            }

                        
                            if  (CreateOrder.OrderNumber > 0)
                            {
                                textBox1.Text = textBox1.Text + "\r\n OrderNumber = " + CreateOrder.OrderNumber.ToString();
                                logger.Log(LogLevel.Debug, "OrderNumber = " + CreateOrder.OrderNumber.ToString());
                                ResultText = "OrderNumber = " + CreateOrder.OrderNumber.ToString();
                                ResultCode = 0;

                                if (CreateOrder.PaymentId != null)
                                {
                                    logger.Log(LogLevel.Debug, "PaymentId = " + CreateOrder.PaymentId.ToString());
                                    textBox1.Text = textBox1.Text + "\r\n PaymentId = " + CreateOrder.PaymentId.ToString();
                                }
                            }
                            else
                            {
                                ResultText = "Error CreateOrder";
                                if (CreateOrder.OrderNumber == 0)
                                {
                                    ResultCode = -1000;
                                }
                                else
                                {
                                    ResultCode = Convert.ToInt32(CreateOrder.OrderNumber);
                                }
                            }
                        }
                        else {
                            if (ResultCode != -10000)
                            {
                                ResultText = "needorder = false.StoreApi.UpdateBasket not confirm details";
                                ResultCode = -4;
                            }
                        }

                        logger.Log(LogLevel.Debug, ResultText + ". ResultCode=" + ResultCode.ToString());
                        sqlcomplete.CommandText = "exec dbo.p_SupplierOrder_Export_Complete @HeadId=" + SupplierHeder.Rows[i][0].ToString() + ",@ResultCode = " + ResultCode.ToString() +",@ResultText='" + ResultText + "'";
                        drcomplete = sqlcomplete.ExecuteReader();
                        drcomplete.Close(); 

                        /*
                        InvokeMegaService(() => { OrderInfo = StoreApi.GetOrderInfo(); });
                        textBox1.Text = textBox1.Text + "\r\n OrderInfo *************************";
                        textBox1.Text = textBox1.Text + "\r\n Amount = " + OrderInfo.Amount.ToString();
                        textBox1.Text = textBox1.Text + "\r\n AmountOrder = " + OrderInfo.AmountOrder.ToString();
                        textBox1.Text = textBox1.Text + "\r\n Lines = " + OrderInfo.Lines.ToString();
                        textBox1.Text = textBox1.Text + "\r\n LinesOrder = " + OrderInfo.LinesOrder.ToString();
                        textBox1.Text = textBox1.Text + "\r\n NumNextOrder = " + OrderInfo.NumNextOrder.ToString();
                        textBox1.Text = textBox1.Text + "\r\n Quantity = " + OrderInfo.Quantity.ToString();
                        textBox1.Text = textBox1.Text + "\r\n QuantityOrder = " + OrderInfo.QuantityOrder.ToString();
                        textBox1.Text = textBox1.Text + "\r\n Weight = " + OrderInfo.Weight.ToString();
                        textBox1.Text = textBox1.Text + "\r\n WeightOrder = " + OrderInfo.WeightOrder.ToString();
                        textBox1.Text = textBox1.Text + "\r\n OrderInfo *************************";
                        */
                        InvokeMegaService(() => { StoreApi.EmptyBasket(); });
                        authClient.Close();
                    }
                    if (authClient!=null)
                    {
                        authClient.Close();
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
                    drcomplete = sqlcomplete.ExecuteReader();
                    drcomplete.Close();
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
                            drcomplete = sqlcomplete.ExecuteReader();
                            drcomplete.Close();
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

        private static string AuthenticateMega()
        {
            string ret = null;
            try
            {
                using (OperationContextScope scope = new OperationContextScope(authClient.InnerChannel))
                {
                    if (authClient.Login(MegaLogin, MegaPassword, "", true))
                    {
                        MessageProperties properties = OperationContext.Current.IncomingMessageProperties;
                        HttpResponseMessageProperty responseProperty =
                            (HttpResponseMessageProperty)properties[HttpResponseMessageProperty.Name];
                        string cookieHeader = responseProperty.Headers[HttpResponseHeader.SetCookie];
                        Match match = Regex.Match(cookieHeader, @".*\.ASPXAUTH=(\w+);.*");
                        ret = match.Groups[1].Value;
                    }
                }
              
            }
            catch (Exception exp) {
                MessageBox.Show(exp.Message.ToString());
            }
            return ret;
        }

        private static  ServiceReference2.StoreApiClient _storeApi;
        
        public static ServiceReference2.StoreApiClient StoreApi
        {
            get
            {
                if (null == _storeApi)
                {
                    _storeApi = new ServiceReference2.StoreApiClient();
                }

                return _storeApi;
            }
        }

        private static void InvokeMegaService(Action action)
        {
          
            using (OperationContextScope scope = new OperationContextScope(StoreApi.InnerChannel))
            {
                var prop = new HttpRequestMessageProperty();
                prop.Headers.Add(HttpRequestHeader.Cookie, String.Format(".ASPXAUTH={0}; Lang={1}", ticketMega, Thread.CurrentThread.CurrentCulture.Name));
                OperationContext.Current.OutgoingMessageProperties[HttpRequestMessageProperty.Name] = prop;

                action();
            }
          
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
