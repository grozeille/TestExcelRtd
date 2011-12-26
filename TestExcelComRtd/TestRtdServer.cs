using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace TestExcelComRtd
{
    [ComVisible(true)]
    [Guid("181ef23c-1304-47d1-80ba-762409839ec6")]
    public class TestRtdServer : IRtdServer
    {
        private static readonly IDictionary<int, SpotListener> topics = new Dictionary<int, SpotListener>();

        private static IRTDUpdateEvent callback = null;

        public object ConnectData(int topicId, ref Array strings, ref bool newValues)
        {
            var isin = strings.GetValue(0) as string;

            LoggingWindow.WriteLine("Connecting: {0} {1}", topicId, isin);

            lock (topics)
            {
                // create a listener for this topic
                var listener = new SpotListener(isin);
                listener.OnUpdate += (sender, args) =>
                {
                    try
                    {
                        if (callback != null)
                        {
                            callback.UpdateNotify();
                        }
                    }
                    catch (COMException comex)
                    {
                        LoggingWindow.WriteLine("Unable to notify Excel: {0}", comex.ToString());
                    }
                };
                listener.Start();

                topics.Add(topicId, listener);
            }

            return "WAIT";
        }

        public void DisconnectData(int topicID)
        {
            LoggingWindow.WriteLine("Disconnecting: {0}", topicID);

            lock (topics)
            {
                var listener = topics[topicID];
                listener.Stop();
                topics.Remove(topicID);
            }
        }

        public int Heartbeat()
        {
            return 1;
        }

        public Array RefreshData(ref int topicCount)
        {
            object[,] result;
            lock (topics)
            {
                // create a list of topic-value pair, if we have a value
                List<Tuple<int, double>> data = new List<Tuple<int, double>>();
                foreach (var t in topics)
                {
                    var last = t.Value.GetLast();
                    if (last != null)
                    {
                        data.Add(new Tuple<int, double>(t.Key, last.Value));
                    }
                }

                // convert to an array for excel
                topicCount = data.Count;
                result = new object[2, topicCount];
                int cpt = 0;
                foreach (var d in data)
                {
                    result[0, cpt] = d.Item1;
                    result[1, cpt] = d.Item2;
                    cpt++;
                }
            }

            return result;
        }

        public int ServerStart(IRTDUpdateEvent CallbackObject)
        {
            LoggingWindow.WriteLine("Starting server");

            callback = CallbackObject;

            return 1;
        }

        public void ServerTerminate()
        {
            LoggingWindow.WriteLine("Stopping server");

            callback = null;
        }

    }
}
