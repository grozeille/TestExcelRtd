using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace TestExcelComRtd
{
    public class SpotListener
    {
        private readonly Queue<double> last = new Queue<double>(1);
        
        public string ISIN { get; private set; }
        
        private readonly object startedLock = new object();

        private bool started = false;

        public SpotListener(string isin)
        {
            ISIN = isin;
        }
        
        public double? GetLast()
        {
            lock (this.last)
            {
                if (this.last.Count > 0)
                {
                    return this.last.Dequeue();
                }
                else
                {
                    return null;
                }
            }
        }

        public void Start()
        {
            started = true;
            new Thread(new ThreadStart(() =>
                {
                    var rand = new Random();
                    var value = 19.27;

                    while (true)
                    {
                        // check if we should stop the thread
                        lock (startedLock)
                        {
                            if (!started)
                            {
                                break;
                            }
                        }

                        // simulate an update of the spot every seconds
                        Thread.Sleep(5 * 1000);
                        value += ((rand.Next(3)-1) * rand.NextDouble());

                        // keep in memory the last value
                        lock (this.last)
                        {
                            this.last.Enqueue(value);
                        }

                        // notify the new value
                        if (OnUpdate != null)
                        {
                            OnUpdate(null, new SpotUpdateEventArgs(this.ISIN, value));
                        }
                    }
                })).Start();
        }

        public void Stop()
        {
            lock (startedLock)
            {
                started = false;
            }
        }

        public event EventHandler<SpotUpdateEventArgs> OnUpdate;
    }
}
