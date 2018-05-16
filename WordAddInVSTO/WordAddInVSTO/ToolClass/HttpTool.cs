using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace WordAddInVSTO
{
    class HttpTool
    {
        /// <summary>
        /// 同步请求API
        /// </summary>
        /// <param name="url"></param>
        /// <param name="postData"></param>
        /// <returns></returns>
        public static HttpWebResponse RequesPost(string url, string postData)
        {
            Stream datastream = null;
            HttpWebResponse response = null;
            HttpWebRequest request = null;
            Encoding encoding = Encoding.UTF8;
            byte[] data = encoding.GetBytes(postData);
            // 设置参数
            request = WebRequest.Create(url) as HttpWebRequest;
            CookieContainer cookieContainer = new CookieContainer();
            request.CookieContainer = cookieContainer;
            request.AllowAutoRedirect = true;
            request.Method = "POST";
            request.ContentType = "application/json";
            request.ContentLength = data.Length;
            datastream = request.GetRequestStream();
            datastream.Write(data, 0, data.Length);
            datastream.Close();
            //发送请求并获取相应回应数据
            response = request.GetResponse() as HttpWebResponse;           

            return response;
        }

        /// <summary>
        /// 异步等待请求API
        /// </summary>
        /// <param name="url"></param>
        /// <param name="postData"></param>
        /// <returns></returns>
        public static async Task<HttpWebResponse> RequesPostAsync(string url, string postData)
        {
            Stream datastream = null;
            HttpWebRequest request = null;
            Encoding encoding = Encoding.UTF8;
            byte[] data = encoding.GetBytes(postData);
            // 设置参数
            request = WebRequest.Create(url) as HttpWebRequest;
            CookieContainer cookieContainer = new CookieContainer();
            request.CookieContainer = cookieContainer;
            request.AllowAutoRedirect = true;
            request.Method = "POST";
            request.ContentType = "application/json";
            request.ContentLength = data.Length;
            datastream = request.GetRequestStream();
            datastream.Write(data, 0, data.Length);
            datastream.Close();

            return await request.GetResponseAsync() as HttpWebResponse;
        }

        public static T GetResponseJson<T>(HttpWebResponse hwr) where T : new()
        {

            Stream responseStream = hwr.GetResponseStream();
            StreamReader sr = new StreamReader(responseStream, Encoding.UTF8);
            string result = sr.ReadToEnd();

            JavaScriptSerializer js = new JavaScriptSerializer();
            T t = js.Deserialize<T>(result);

            return t;
        }

        /// <summary>
        /// API 返回执行结果
        /// </summary>
        public class ResponseResult
        {
            /// <summary>
            /// API 执行结果状态码 2000 OK， 4000 及以上代表出错
            /// </summary>
            public int Code { get; set; }
            /// <summary>
            /// API 执行结果信息
            /// </summary>
            public string Message { get; set; }
        }
    }
}
