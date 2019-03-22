using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Text;
using System.Net;
using System.IO;
using System.Diagnostics;

namespace SimpleHttpRequest
{
    public class HttpJsonRequest
    {
        public string DoPost(string URL, string Data, string UserAgent = "Mozilla/5.0 (Windows NT 6.3; Trident/7.0; rv 11.0) like Gecko", CookieContainer Cookies = null)
        {
            HttpWebRequest JsonRequest = GetJSONWebRequest(URL, "POST", UserAgent, Cookies);
            byte[] DataToPost = Encoding.UTF8.GetBytes(Data);
            Stream RequestStream = JsonRequest.GetRequestStream();
            RequestStream.Write(DataToPost, 0, DataToPost.Length);
            RequestStream.Close();
            HttpWebResponse Response = (HttpWebResponse)JsonRequest.GetResponse();
            Encoding DataEncoding = Encoding.GetEncoding(Response.CharacterSet);
            return GetResponseAsString(Response, DataEncoding);
        }
        public string DoGet(string URL, string UserAgent = "Mozilla/5.0 (Windows NT 6.3; Trident/7.0; rv 11.0) like Gecko", CookieContainer Cookies = null)
        {
            HttpWebRequest JsonRequest = GetJSONWebRequest(URL, "GET", UserAgent, Cookies);
            HttpWebResponse Response = (HttpWebResponse)JsonRequest.GetResponse();
            Encoding DataEncoding = Encoding.GetEncoding(Response.CharacterSet);
            return GetResponseAsString(Response, DataEncoding);
        }
        public HttpWebRequest GetJSONWebRequest(string URL, string Method, string UserAgent, CookieContainer Cookies = null)
        {
            HttpWebRequest HttpRequest = (HttpWebRequest)HttpWebRequest.Create(URL);
            HttpRequest.ServicePoint.Expect100Continue = false;
            HttpRequest.ContentType = "application/json";
            HttpRequest.Method = Method;
            HttpRequest.KeepAlive = true;
            HttpRequest.UserAgent = UserAgent;
            HttpRequest.Timeout = 1000000;
            HttpRequest.CookieContainer = Cookies;
            HttpRequest.Proxy = null;
            return HttpRequest;
        }
        public string GetResponseAsString(HttpWebResponse Response, Encoding DataEncoding)
        {
            StringBuilder Result = new StringBuilder();
            Stream ResponseStream = null;
            StreamReader ResponseReader = null;
            try
            {
                // 以字符流的方式读取HTTP响应
                ResponseStream = Response.GetResponseStream();
                ResponseReader = new StreamReader(ResponseStream, DataEncoding);
                // 每次读取不大于256个字符，并写入字符串
                char[] buffer = new char[256];
                int readBytes = 0;
                while ((readBytes = ResponseReader.Read(buffer, 0, buffer.Length)) > 0)
                {
                    Result.Append(buffer, 0, readBytes);
                }
            }
            finally
            {
                // 释放资源
                if (ResponseReader != null) ResponseReader.Close();
                if (ResponseStream != null) ResponseStream.Close();
                if (Response != null) Response.Close();
            }
            return Result.ToString();
        }
        public CookieContainer CookieStringToContainer(string CookieData, CookieContainer Cookies, string DomainName)
        {
            string[] CookieStrings = CookieData.Split(';');
            string CookieString = null;
            int Equallength = 0; //=的位置 
            string CookieKey = null;
            string CookieValue = null;
            for (int i = 0; i < CookieStrings.Length; i++)
            {
                if (!string.IsNullOrEmpty(CookieStrings[i]))
                {
                    CookieString = CookieStrings[i];
                    Equallength = CookieString.IndexOf("=");
                    if (Equallength != -1)       //有可能cookie 无=，就直接一个cookiename；比如:a=3;ck;abc=; 
                    {
                        CookieKey = CookieString.Substring(0, Equallength).Trim();
                        //cookie= 
                        if (Equallength == CookieString.Length - 1)    //这种是等号后面无值，如：abc=; 
                        {
                            CookieValue = "";
                        }
                        else
                        {
                            CookieValue = CookieString.Substring(Equallength + 1, CookieString.Length - Equallength - 1).Trim();
                        }
                    }
                    else
                    {
                        CookieKey = CookieString.Trim();
                        CookieValue = "";
                    }
                    Cookies.Add(new Cookie(CookieKey, CookieValue, "", DomainName));
                }
            }

            return Cookies;
        }

    /*
        作者：liheao 
        来源：CSDN 
        原文：https://blog.csdn.net/liheao/article/details/80597109 
        版权声明：本文为博主原创文章，转载请附上博文链接！
    */
    }
}
