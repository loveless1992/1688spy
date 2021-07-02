using System;
using System.IO;
using System.Net;
using System.Text;
using System.Web;

namespace 爬爬爬
{
    public class HttpRequestUtil
    {
        public string GetPageHtml(string strURL)
        {
            //Uri url = new Uri(strURL, false);
            HttpWebRequest request;
            request = (HttpWebRequest)WebRequest.Create(strURL);
            request.Method = "POST"; //Post请求方式
            request.ContentType = "text/html;charset=utf-8"; //内容类型
            string paraUrlCoded = HttpUtility.UrlEncode(""); //参数经过URL编码
            byte[] payload;
            payload = System.Text.Encoding.GetEncoding("UTF-8").GetBytes(paraUrlCoded); //将URL编码后的字符串转化为字节
            request.ContentLength = payload.Length; //设置请求的ContentLength
            Stream writer = request.GetRequestStream(); //获得请求流
            writer.Write(payload, 0, payload.Length); //将请求参数写入流
            writer.Close(); //关闭请求流
            HttpWebResponse response;
            response = (HttpWebResponse)request.GetResponse(); //获得响应流
            Stream s;
            s = response.GetResponseStream();
            StreamReader objReader = new StreamReader(s, System.Text.Encoding.GetEncoding("UTF-8"));
            string HTML = "";
            string sLine = "";
            int i = 0;
            while (sLine != null)
            {
                i++;
                sLine = objReader.ReadLine();
                if (sLine != null)
                    HTML += sLine;
            }
            //HTML = HTML.Replace("&lt;","<");
            //HTML = HTML.Replace("&gt;",">");
            string html = HTML;
            return html;
        }

        public string GetPageHtml2(string strUrl)
        {
            WebClient MyWebClient = new WebClient();
            MyWebClient.Credentials = CredentialCache.DefaultCredentials; //获取或设置用于向Internet资源的请求进行身份验证的网络凭据
            Byte[] pageData = MyWebClient.DownloadData(strUrl); //从指定网站下载数据
            string pageHtml = Encoding.Default.GetString(pageData); //如果获取网站页面采用的是GB2312，则使用这句    
            //string pageHtml = Encoding.UTF8.GetString(pageData); //如果获取网站页面采用的是UTF-8，则使用这句
            return pageHtml;
        }
    }
}