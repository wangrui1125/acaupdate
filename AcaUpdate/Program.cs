using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Data;
using System.Net;
using System.Data.OleDb;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
//using System.Net.Json;


//using Excel = Microsoft.Office.Interop.Excel; 

namespace AcaUpdate
{
    public struct CAuthorPaper
    {
        public string name;
        public string department;
        public string papername;
        public int paperlevel;
       
    }
    public struct CAuthor
    {
        //public int author_id;
        public string name;
        //public string department;
        public string department;
        public string acadepartm;
        public string departmentauto;
    }

    public struct CData
    {
        public string academicianname;
        public string name;
        public string department;
        public int leve1number;
        public int leve2number;
    }
    public struct Academincian {
        public String name;
        public String flage;
        public String url;
        public String detail1;
        public String detail2;
        public String imageurl;
    
    }
    class Program
    {
        static SqlConnection m_Connection;
        static List<CAuthorPaper> namesall = new List<CAuthorPaper>();
        static Dictionary<CAuthor, int> dictionarynames = new Dictionary<CAuthor, int>();
        static List<Academincian> academician = new List<Academincian>();
        static String URL2 = "http://www.cae.cn/cae/html/main/col48/column_48_1.html";
        static String URL1 = "http://www.casad.cas.cn/chnl/371/index.html";
        //static CAuthor[] dictionarynames=new CAuthor[];
        static void Main(string[] args)
        {
            connectDatabase();
            DataTable dt = LoadDataFromExcel("..\\data200500.xlsx");
            HashSet<string> acaold = new HashSet<string>();
            for (int row = 0; row < dt.Rows.Count; row++)
            {
                string author = dt.Rows[row][1].ToString() + dt.Rows[row][2].ToString();
                acaold.Add(author);
            }
            //List<Academincian> acas = searcheraca1(acaold);//搜索新的院士
            List<Academincian> acas2 = searcheraca2(acaold);
            writetodatabase(acas2, "2017");
            birthdayupdata("2017");//更新单位和生日
           
             //writedepartment("Y:\\linktest\\data200500.xlsx");
            //writedepartment("Y:\\linktest\\data501700.xlsx");
            int i = 1;

        }
        private static void connectDatabase()
        {
            string SqlConnectionString = "Data Source=166.111.7.152;Initial Catalog=dbpaper;uid=sa;pwd=kxmsql8!";
            m_Connection = new SqlConnection(SqlConnectionString);
            m_Connection.Open();
        }
        private static List<Academincian> searcheraca1(HashSet<string> acaold)
        {
            /**
             * 搜索中国科学院院士
             */
            List<Academincian> acas = new List<Academincian>(); 
            HttpWebRequest request;
            request = (HttpWebRequest)WebRequest.Create(URL1);
            request.Method = "POST"; //Post请求方式
            request.UserAgent = "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 5.1; .NET CLR 2.0.50727; .NET CLR 3.0.04506.648; .NET CLR 3.5.21022)";
            HttpWebResponse response;
            string m_Html = "";
            string sLine = "";
            try
            {
                Stream writer = request.GetRequestStream(); //获得请求流
                response = (HttpWebResponse)request.GetResponse(); //获得响应流
                Stream s;
                s = response.GetResponseStream();
                StreamReader objReader = new StreamReader(s, System.Text.Encoding.UTF8);
                int i = 0;
                while (sLine != null)
                {
                    i++;
                    sLine = objReader.ReadLine();
                    if (sLine != null)
                        m_Html += sLine;
                }
            }
            catch (WebException ex1)
            {
                Console.WriteLine(ex1);
                
            }
            catch (OutOfMemoryException ex2)
            {
                Console.WriteLine(ex2);
                
            }
            catch (IOException ex3)
            {
                Console.WriteLine(ex3);
            }

            MatchCollection matches =Regex.Matches(m_Html,@"(<span><a href=\"")(.*?)(</a>)");
            for (int i = 0; i < matches.Count; i++)
            {
                string temp=matches[i].Groups[2].Value;
                string url=temp.Substring(0,temp.IndexOf("\""));
                string name = temp.Substring(temp.IndexOf(">")+1);  

                Academincian aca=new Academincian();
                aca.name =name;
                aca.url = url;
                aca.flage = "中国科学院";
                if (!acaold.Contains("中国科学院" + name))
                {
                    string[] result = searcherdetail1(url);
                    aca.imageurl = result[0];
                    aca.detail1 = result[1];
                    aca.detail2 = result[2];                   
                    acas.Add(aca);
                    Console.WriteLine(name);
                }
            }


            /*
            HtmlAgilityPack.HtmlDocument m_Document = new HtmlAgilityPack.HtmlDocument();
            m_Document.LoadHtml(m_Html);
            HtmlNode em = m_Document.GetElementbyId("allNameBar");
            HtmlNodeCollection ems = em.ChildNodes;
            foreach (HtmlNode em2 in ems)
            {
                if (em2.Name=="dd")
                {
                    HtmlNodeCollection ems2 = em2.ChildNodes;
                    foreach (HtmlNode em3 in ems2)
                    {
                        if (em3.Name == "span")
                        {
                            HtmlNode em4=em3.FirstChild;
                            String name = em4.InnerText;                            
                            String url = em4.GetAttributeValue("href", "");
                            Console.WriteLine(name+":"+url);
                            if(url=="")
                            {
                                Console.WriteLine(name+"没有找到链接");
                            }
                           
                                Academincian aca=new Academincian();
                                aca.name =name;
                                aca.url = url;
                                aca.flage = "中国科学院";
                            
                        }
                    }
                }
            }
             * */
            return acas;
        }
        private static List<Academincian> searcheraca2(HashSet<string> acaold)
        {
            /**
            * 搜索中国工程院院士
            */
            List<Academincian> acas = new List<Academincian>(); 
            HttpWebRequest request;
            request = (HttpWebRequest)WebRequest.Create(URL2);            
            //request.Method = "GET"; //Post请求方式
            //request.UserAgent = "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 5.1; .NET CLR 2.0.50727; .NET CLR 3.0.04506.648; .NET CLR 3.5.21022)";
            HttpWebResponse response;
            string m_Html = "";
            string sLine = "";
            try
            {
               // Stream writer = request.GetRequestStream(); //获得请求流
                response = (HttpWebResponse)request.GetResponse(); //获得响应流
                Stream s;
                s = response.GetResponseStream();
                StreamReader objReader = new StreamReader(s, System.Text.Encoding.UTF8);
                int i = 0;
                while (sLine != null)
                {
                    i++;
                    sLine = objReader.ReadLine();

                    if (sLine != null)
                    {
                       
                        m_Html += sLine;
                    }
                }
            }
            catch (WebException ex1)
            {
                Console.WriteLine(ex1);

            }
            catch (OutOfMemoryException ex2)
            {
                Console.WriteLine(ex2);

            }
            catch (IOException ex3)
            {
                Console.WriteLine(ex3);
            }

            MatchCollection matches = Regex.Matches(m_Html, @"name_list([\s\S]*?)(</a></li>)");
            int cont = matches.Count;
            HashSet<String> urls = new HashSet<string>(); 
            for(int i=0;i<cont;i++)
            {
                 string temp=matches[i].Value;
                 Match match = Regex.Match(temp, @"(href=)([\s\S]*?)( t)");
                 string url = "http://www.cae.cn" + match.Value.Replace("href=\"", "").Replace("\" t", "").Replace("jump", "introduction");                 
                 if (urls.Add(url))
                 {
                     match = Regex.Match(temp, @"([\u4E00-\u9FFF]+)");
                     String name = match.Value;
                     if (!acaold.Contains("中国工程院" + name))
                     {
                         Academincian academician = new Academincian();
                         academician.name = name;
                         academician.url = url;
                         academician.flage = "中国工程院";
                         String[] result = new String[3];
                         result = searcherdetail2(url);
                         academician.imageurl = result[0];
                         academician.detail1 = result[1];
                         academician.detail2 = result[2];
                         Console.WriteLine(name);
                         acas.Add(academician);
                     }
                 }
            }
            return acas;
        }
        private static String[] searcherdetail1(String url)
        {
            /**
           * 搜索中国科学院院士详细
           */
            String[] result =new String [3];

			 HttpWebRequest request;
            request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = "POST"; //Post请求方式
            request.UserAgent = "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 5.1; .NET CLR 2.0.50727; .NET CLR 3.0.04506.648; .NET CLR 3.5.21022)";
            HttpWebResponse response;
            string m_Html = "";
            string sLine = "";
            try
            {
                Stream writer = request.GetRequestStream(); //获得请求流
                response = (HttpWebResponse)request.GetResponse(); //获得响应流
                Stream s;
                s = response.GetResponseStream();
                StreamReader objReader = new StreamReader(s, System.Text.Encoding.UTF8);
                int i = 0;
                while (sLine != null)
                {
                    i++;
                    sLine = objReader.ReadLine();
                    if (sLine != null)
                        m_Html += sLine;
                }
                
            }
            catch (WebException ex1)
            {
                Console.WriteLine(ex1);
                
            }
            catch (OutOfMemoryException ex2)
            {
                Console.WriteLine(ex2);
                
            }
            catch (IOException ex3)
            {
                Console.WriteLine(ex3);
            }
            Match match1 = Regex.Match(m_Html, @"(<img src=\"")(.*?)(\"")");
            result[0] = match1.Groups[2].Value;
            HtmlAgilityPack.HtmlDocument m_Document = new HtmlAgilityPack.HtmlDocument();
            m_Document.LoadHtml(m_Html);
            HtmlNode em = m_Document.GetElementbyId("zoom");
            result[1] = em.InnerText;
            int index = 0;HtmlNode detailnode=null;
            foreach (HtmlNode emtemp in em.ChildNodes)
            {
                if(index==3)
                {
                 detailnode=emtemp;   
                }
                index++;
            }

            try {
                if (detailnode.FirstChild.ChildNodes.Count < 1)
                {
                    result[1] = detailnode.InnerText;
                    result[2] = detailnode.NextSibling.InnerText;
                }
                else
                {
                    result[1] = detailnode.FirstChild.InnerText;
                    result[2] = detailnode.FirstChild.NextSibling.InnerText;
                }                               
            }
            catch(Exception e) 
            { ;}
            return result;

        }
        private static String[] searcherdetail2(String url)
        {
            /**
           * 搜索中国工程院院士详细
           */
            String[] result = new String[3];
		    HttpWebRequest request;
            request = (HttpWebRequest)WebRequest.Create(url);
            //request.Method = "POST"; //Post请求方式
            //request.UserAgent = "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 5.1; .NET CLR 2.0.50727; .NET CLR 3.0.04506.648; .NET CLR 3.5.21022)";
            HttpWebResponse response;
            string m_Html = "";
            string sLine = "";
            try
            {
               // Stream writer = request.GetRequestStream(); //获得请求流
                response = (HttpWebResponse)request.GetResponse(); //获得响应流
                Stream s;
                s = response.GetResponseStream();
                StreamReader objReader = new StreamReader(s, System.Text.Encoding.UTF8);
                int i = 0;
                while (sLine != null)
                {
                    i++;
                    sLine = objReader.ReadLine();
                    if (sLine != null)
                        m_Html += sLine;
                }
            }
            catch (WebException ex1)
            {
                Console.WriteLine(ex1);
                
            }
            catch (OutOfMemoryException ex2)
            {
                Console.WriteLine(ex2);
                
            }
            catch (IOException ex3)
            {
                Console.WriteLine(ex3);
            }
            //照片url
            Match match = Regex.Match(m_Html, @"(info_img)([\s\S]*?)(</div>)");
            string temp = match.Value.Replace("info_img", "").Replace("</div>", "");
            match = Regex.Match(temp, @"(src=)([\s\S]*?)(style)");
            temp ="http://www.cae.cn"+ match.Value.Replace("src=\"", "").Replace("\" style","");
            result[0] = temp;
            match = Regex.Match(m_Html, @"(intro)([\s\S]*?)(</div>)");
            temp=match.Value.Replace("intro\">\t\t","").Replace("</div>","");
            MatchCollection matches = Regex.Matches(temp, @"<p>([\s\S]*?)</p>");
            string detail = "";int j=1;
            for (int i = 0; i < matches.Count; i++)
            {                
                detail = matches[i].Groups[1].Value;
                if (detail.Length > 10)
                {
                    result[j++] = detail.Replace("&ensp;","");
                    
                }
                if (j > 2)
                {
                    break;
                }
            }
           
           
            return result;

        }

        private static void writetotxt(int[,] papernumber, int m, int n)
        {
            FileStream fs = new FileStream("C:\\ak.txt", FileMode.Create);

            for (int i = 0; i < m; i++)
            {
                for (int j = 0; j < n; j++)
                {
                    byte[] data = System.Text.Encoding.Default.GetBytes(papernumber[i,j].ToString()+" ");
                    //开始写入
                    fs.Write(data, 0, data.Length);
                }
                    byte[] dataline = System.Text.Encoding.Default.GetBytes("\n\r");
                //开始写入
                    fs.Write(dataline, 0, dataline.Length);
            }
            //清空缓冲区、关闭流
            fs.Flush();
            fs.Close();
            //获得字节数组
           
        }
        private static void writedepartment(String path)
        {
           
            DataTable dt =LoadDataFromExcel(path);
            List<CAuthor> authors = new List<CAuthor>();

            for (int row = 0; row < dt.Rows.Count; row++)
            {

                CAuthor author = new CAuthor();
                author.department = "";
                author.departmentauto = "";
                author.acadepartm = dt.Rows[row][1].ToString();
                author.name = dt.Rows[row][2].ToString();

                string tempdepartment = dt.Rows[row][4].ToString();
                if (tempdepartment.CompareTo("") != 0)
                {
                    string[] tempdepartments = tempdepartment.Split('；');
                    author.department = tempdepartments[0].Replace("（）", "").Replace("()", "");
                }

                tempdepartment = dt.Rows[row][3].ToString();
                Match match = Regex.Match(tempdepartment, @"([\u4E00-\u9FFF0-9]+)教授");
                if (match.Length > 0)
                {
                    author.departmentauto = match.Value.Replace("教授", "").Replace("现任", "");
                    
                }
                else
                {
                    match = Regex.Match(tempdepartment, @"([\u4E00-\u9FFF0-9]+)研究员");
                    if (match.Length > 0)
                    {
                        author.departmentauto = match.Value.Replace("研究员", "").Replace("现任","");
                    }
                }

                authors.Add(author);
            }

            foreach (CAuthor author in authors)
            {
                Console.WriteLine(author.name);
                string cc = String.Format("update Academincian set departement='{0}', departementauto='{3}'where name='{1}' and title='{2}'", author.department, author.name, author.acadepartm, author.departmentauto);

                SqlCommand m_Command = new SqlCommand(cc, m_Connection);
                try
                {
                    m_Command.ExecuteNonQuery();
                }
                catch (SqlException ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
        
        }
        private static DataTable LoadDataFromExcel(string Path)
        {
            string strConn = "Provider=Microsoft.Ace.OleDb.12.0;" + "data source=" + Path + ";Extended Properties='Excel 12.0; HDR=Yes; IMEX=1'";
            OleDbConnection conn = new OleDbConnection(strConn);
            conn.Open();
            string strExcel = "";
            OleDbDataAdapter myCommand = null;
            DataTable dt = null;
            strExcel = "select  * from [sheet1$]";
            myCommand = new OleDbDataAdapter(strExcel, strConn);
            dt = new DataTable();
            myCommand.Fill(dt);
            return dt;
        }
        private static void writetodatabase(List<Academincian> acas,string thisyear)
       {
           SqlCommand m_Command=null;
           foreach(Academincian aca in acas)
           {
                string sq=string.Format("insert into Academincian (name,url,detail1,detail2,imageurl,title,acayear) "+
                                        "values('{0}','{1}','{2}','{3}','{4}','{5}','{6}')",aca.name,aca.url,
                    aca.detail1,aca.detail2,aca.imageurl,aca.flage,thisyear);
                m_Command = new SqlCommand(sq, m_Connection);
                try
                {
                    m_Command.ExecuteNonQuery();
                }
                catch (SqlException ex)
                {
                    Console.WriteLine(ex.Message);
                }

            }


       }
        private static void birthdayupdata(string thisyear)
       {
           //connectDatabase();
           List<String> details = new List<string>();
           List<int> ids = new List<int>();
           List<String> birthdays = new List<string>();
           List<string> titles = new List<string>();


           string cc = "select aca_id,detail1,title from academincian where acayear like "+thisyear;
           SqlCommand m_Command = new SqlCommand(cc, m_Connection);
           SqlDataReader sdr;
           sdr = m_Command.ExecuteReader();
           while (sdr.Read())
           {
               String detail1; string title;
               int id;
               //newAuthor.author_id = sdr.GetInt32(sdr.GetOrdinal("id"));
               try
               {

                   detail1 = sdr.GetString(sdr.GetOrdinal("detail1"));
                   title = sdr.GetString(sdr.GetOrdinal("title"));
                   id = sdr.GetInt32(sdr.GetOrdinal("aca_id"));
                   details.Add(detail1);
                   titles.Add(title);
                   ids.Add(id);
               }
               catch (Exception e)
               {
                   Console.WriteLine(e.ToString());
               }
           }
           sdr.Close();
           int number=0;
           foreach (String detail in details)
           {
               
               //生日
               String birth = "";
               String acayear = "";
               if (titles[number].Equals("中国科学院"))
               {
                   Match match = Regex.Match(detail, "[0-9]{4}年[0-9]{1,2}月");
                   
                   if (match.Value == "")
                   {
                       match = Regex.Match(detail, "[0-9]{4}年生");
                   }
                    birth = match.Value;
                   Console.WriteLine(birth);
                   Match match2 = Regex.Match(detail, "[0-9]{4}年[\u4E00-\u9FFF]+中[\u4E00-\u9FFF]*科[\u4E00-\u9FFF]*院");
                   if (match2.Value.Length > 0)
                   {
                       int temp = match2.Value.IndexOf("年");
                       acayear = match2.Value.Substring(0, temp + 1);
                   }
               }
               else
               {
                   
                   Match match2 = Regex.Match(detail, "[0-9]{4}年[\u4E00-\u9FFF]+中国工程院");
                   if (match2.Value.Length > 0)
                   {
                       int temp = match2.Value.IndexOf("年");
                       acayear = match2.Value.Substring(0, temp + 1);
                   }
                   Match match = Regex.Match(detail, "[0-9]{4}年[0-9]{1,2}月");
                   birth = match.Value;
                   Console.WriteLine(birth);
                   
                   if (match.Value=="")
                   {
                       match = Regex.Match(detail, "[0-9]{4}.[0-9]{1,2}.[0-9]{1,2}");
                       if (match.Value == "")
                       {
                           match = Regex.Match(detail, "[0-9]{4}.[0-9]{1,2}");
                       }
                       if (match.Value.Length > 0)
                       {
                           String tempbirth = match.Value;
                           String[] biths = tempbirth.Split('.');
                           if (biths.Length == 2)
                           {
                               birth = biths[0] + "年" + biths[1] + "月";
                               Console.WriteLine(birth);

                           }
                           if (biths.Length == 3)
                           {
                               birth = biths[0] + "年" + biths[1] + "月" + biths[2] + "日";
                               Console.WriteLine(birth);

                           }
                       
                       }
                   }
                   birthdays.Add(birth);
                   
                  }

               //单位
               string departmentauto = detail;
               Match matchdepart = Regex.Match(departmentauto, @"([\u4E00-\u9FFF0-9]+)教授");
               if (matchdepart.Length > 0)
               {
                   departmentauto = matchdepart.Value.Replace("教授", "").Replace("现任", "");

               }
               else
               {
                   matchdepart = Regex.Match(departmentauto, @"([\u4E00-\u9FFF0-9]+)研究员");
                   if (matchdepart.Length > 0)
                   {
                       departmentauto = matchdepart.Value.Replace("研究员", "").Replace("现任", "");
                   }
               }


               string ccwr = String.Format("update academincian set birthday='{0}',departementauto='{2}' where aca_id={1}", birth, ids[number++], departmentauto);
               SqlCommand m_Commandwr = new SqlCommand(ccwr, m_Connection);
                           try
                           {
                               m_Commandwr.ExecuteNonQuery();
                           }
                           catch (SqlException ex)
                           {
                               Console.WriteLine(ex.Message);
                           }
            }
           m_Connection.Close();
       }
    }
}
