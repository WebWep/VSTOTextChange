using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Web;
using System.Web.Http;

namespace WordAPI.Controllers
{
    public class WordAPIController : ApiController
    {
        [HttpPost]
        public ResponseResult PostDocument(dynamic documentObj)
        {

            try
            {
                string _document = Convert.ToString(documentObj.document);

                if (string.IsNullOrEmpty(_document))
                {
                    return new ResponseResult { Code = 4002, Message = "内容不能为空！" };
                }

                System.Web.HttpServerUtility Server = System.Web.HttpContext.Current.Server;
                string all_path = Server.MapPath("~/Upload");
                if (Directory.Exists(all_path))
                {
                    WriteTxt(all_path + "/Document.txt", _document, false);
                }
                else
                {
                    Directory.CreateDirectory(all_path);
                    WriteTxt(all_path + "/Document.txt", _document, false);
                }

                return new ResponseResult { Code = 2000, Message = "操作成功！" };
            }
            catch (Exception ex)
            {
                return new ResponseResult { Code = 4001, Message = ex.Message };
            }
        }

        [HttpPost]
        public ResponseResult PostSentence(dynamic sentenceObj)
        {
            try
            {
                string _sentence = Convert.ToString(sentenceObj.sentence);
                if (string.IsNullOrEmpty(_sentence))
                {
                    return new ResponseResult { Code = 4002, Message = "内容不能为空！" };
                }

                System.Web.HttpServerUtility Server = System.Web.HttpContext.Current.Server;
                string all_path = Server.MapPath("~/Upload");
                if (Directory.Exists(all_path))
                {

                    WriteTxt(all_path + "/Sentence.txt", _sentence, true);
                }
                else
                {
                    Directory.CreateDirectory(all_path);
                    WriteTxt(all_path + "/Sentence.txt", _sentence, true);
                }
                return new ResponseResult { Code = 2000, Message = "操作成功！" };
            }
            catch (Exception ex)
            {
                return new ResponseResult { Code = 4004, Message = ex.Message };
            }
        }

        /// <summary>
        /// 向 TxT 文件写入数据
        /// 文件不存在创建
        /// 文件存在追加内容
        /// </summary>
        /// <param name="path"></param>
        /// <param name="content"></param>
        /// <param name="append">是否追加</param>
        protected void WriteTxt(string path, string content, bool append)
        {
            System.IO.StreamWriter sw = new System.IO.StreamWriter(path, append, Encoding.UTF8);
            try
            {
                sw.WriteLine(content);
                sw.Close();
            }
            catch (Exception ex)
            {
                throw new Exception("写入文件出错", ex);
            }
            finally
            {
                sw.Close();
            }
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
            public string Message { get; set; }
        }

    }
}
