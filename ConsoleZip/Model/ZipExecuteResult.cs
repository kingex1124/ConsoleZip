using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleZip
{
    public class ZipExecuteResult
    {
        public bool IsSuccessed { get; set; }

        public string Message { get; set; }

        public ZipExecuteResult()
        {

        }

        public ZipExecuteResult(bool isSuccessed, string message)
        {
            IsSuccessed = isSuccessed;
            Message = message;
        }

        public static ZipExecuteResult Ok()
        {
            return new ZipExecuteResult { IsSuccessed = true };
        }

        public static ZipExecuteResult Fail(string errMsg)
        {
            return new ZipExecuteResult { IsSuccessed = false, Message = errMsg };
        }
    }

    public class ZipExecuteResult<T>: ZipExecuteResult
    {
        public T Data { get; set; }

        public ZipExecuteResult()
        {
            Data = default(T);
        }

        public static ZipExecuteResult<T> Ok(T data)
        {
            return new ZipExecuteResult<T> { IsSuccessed = true, Data = data };
        }

        public static ZipExecuteResult<T> Ok(T data, string msg)
        {
            return new ZipExecuteResult<T> { IsSuccessed = true, Message = msg, Data = data };
        }

        public static ZipExecuteResult<T> Fail(string errMsg)
        {
            return new ZipExecuteResult<T> { IsSuccessed = false, Message = errMsg };
        }
    }
}
