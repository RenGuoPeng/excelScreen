using System;
using System.Collections.Generic;
using System.Threading;

namespace excelScreen
{
    public class TaskFactor
    {
        /// <summary>
        /// 请求参数对象
        /// </summary>
        public class TaskPara
        {
            /// <summary>
            /// 自定义的参数数组
            /// </summary>
            public object[] Paras { get; set; }

            public List<string> Checknum { get; set; }
            /// <summary>
            /// 回调方法
            /// </summary>
            public Action<object> Callback { get; set; }

            public void Invoke(object obj)
            {
                try
                {
                    Callback?.Invoke(obj);
                }
                catch
                {
                    //ignore
                }
            }
        }

        /// <summary>
        /// 类型转换
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="anonymous">需要转换的对象</param>
        /// <param name="anonymousType">转换后的匿名类</param>
        /// <returns></returns>
        public static T CastAnonymous<T>(object anonymous, T anonymousType)
        {
            return (T)anonymous;
        }

        /// <summary>
        /// 发起一个异步请求
        /// </summary>
        /// <param name="target">函数回掉</param>
        /// <param name="obj">请求发送的参数</param>
        public static void NewTask(Action<object> target, TaskPara obj)
        {
            ParameterizedThreadStart parStart = new ParameterizedThreadStart(target);
            Thread myThread = new Thread(parStart);
            myThread.IsBackground = true;
            myThread.Start(obj);
        }
    }
}