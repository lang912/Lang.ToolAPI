using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lang.ToolBiz.Model
{
    public class Result
    {
        /// <summary>
        /// true 表示处理成功
        /// </summary>
        public bool Success { get { return Code == 200; } private set { } }

        /// <summary>
        /// 失败时的理由
        /// </summary>
        public string Msg { get; set; }

        /// <summary>
        /// code=200 表示操作成功,业务操作中，仅对code和msg赋值即可
        /// </summary>
        public int Code { get; set; }
    }

    public class Result<T>: Result
    {
        /// <summary>
        /// 返回的数据
        /// </summary>
        public T Data { get; set; }
    }
}
