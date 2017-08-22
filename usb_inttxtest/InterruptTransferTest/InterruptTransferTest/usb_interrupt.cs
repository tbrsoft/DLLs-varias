using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using System.Windows.Forms;

using PVOID = System.IntPtr;
using DWORD = System.UInt32;

namespace usb_api
{
    public delegate void NewUSBIntDataEventHandler(NewUSBIntDataEventArgs e);

    /// <summary>
    /// This class extends the usb_interface class to add support for asynchronous reading of EP2.
    /// </summary>
    unsafe public class USBInterruptInterface : usb_interface
    {
        protected string in_pipe_async = "\\MCHP_EP2_ASYNC";

        protected void OpenIntPipe()
        {
            DWORD selection = 0; // Selects the device to connect to, in this example it is assumed you will only have one device per vid_pid connected.
            myInPipe = _MPUSBOpen(selection, vid_pid_norm, in_pipe_async, 1, 0);
        }

        public uint TryReceive(out byte[] outarray)
        {
            byte* receive_buf = stackalloc byte[64];
            uint rxlen;
            uint rval = ReceiveIntPacket(receive_buf, &rxlen);
            if (rval != 1)
            {
                outarray = null;
                return rval;
            }
            outarray = new byte[rxlen];
            for (int i = 0; i < rxlen; i++)
            {
                outarray[i] = receive_buf[i];
            }
            return 1;
        }

        public USBInterruptInterface()
        {
            OpenIntPipe();
        }

        ~USBInterruptInterface()
        {
            if (myInPipe != null)
                _MPUSBClose(myInPipe);
        }

        public void Dispose()
        {
            if (myInPipe != null)
                _MPUSBClose(myInPipe);
            System.GC.SuppressFinalize(this);
        }

        protected DWORD ReceiveIntPacket(byte* ReceiveData, DWORD* ReceiveLength)
        {
            uint ReceiveDelay = 0; //check buffer and return immediately
            DWORD RxLen = (DWORD)64;
            if (_MPUSBReadInt(myInPipe, (void*)ReceiveData, RxLen, &RxLen, ReceiveDelay) == 1)
            {
                *ReceiveLength = RxLen;
                return 1;   // Success
            }
            return 0;  //Failure
        }
    }
    
    /// <summary>
    /// Event args for the EventNewUSBIntData event.  Received packet is in the byte array 'newdata'.
    /// </summary>
    public class NewUSBIntDataEventArgs : EventArgs
    {
        public NewUSBIntDataEventArgs(byte[] data)
        {
            newdata = data;
        }
        public byte[] newdata;
    }    
    
    /// <summary>
    /// This control is a simple manager for running the polling of an asynchronous PIC USB endpoint
    /// in a separate execution thread to take a load off of the GUI and main program.  The 
    /// EventNewUSBIntData event is fired when new data is received, with the packet in a byte array
    /// in the event args.
    /// </summary>
    public class PICAsyncManager:UserControl
    {
        public delegate void CallBack(byte[] data);
        public event NewUSBIntDataEventHandler EventNewUSBIntData;

        Thread IntThread;

        private void NewUSBIntData(byte[] data)
        {
            CallBack d = new CallBack(FireEventNewUSBIntData);
            this.Invoke(d, new object[] { data });
        }

        private void FireEventNewUSBIntData(byte[] data)
        {
            NewUSBIntDataEventArgs newDataArgs = new NewUSBIntDataEventArgs(data);
            EventNewUSBIntData(new NewUSBIntDataEventArgs(data));
        }

        public bool IsRunning
        {
            get
            {
                if (IntThread == null)
                    return false;
                return IntThread.IsAlive;
            }
        }

        public void Start()
        {
            if (IntThread == null || !IntThread.IsAlive)
            {
                PICAsync pa = new PICAsync(new CallBack(NewUSBIntData));
                IntThread = new Thread(new ThreadStart(pa.IntWatch));
                IntThread.IsBackground = true;
                IntThread.Start();
            }
        }

        public void Stop()
        {
            if(IntThread!=null)
                IntThread.Abort();
        }

        ~PICAsyncManager()
        {
            if(IntThread!=null)
                IntThread.Abort();
        }

        /// <summary>
        /// This class performs the polling of the asynchronous endpoint buffer.
        /// An instance of this will be run in a separate execution thread and throw packets back
        /// to the main thread via the callback method.
        /// </summary>
        class PICAsync
        {
            private CallBack CallBackMethod;

            public PICAsync(CallBack cb)
            {
                CallBackMethod = cb;
            }

            public void IntWatch()
            {
                USBInterruptInterface usbi = new USBInterruptInterface();
                byte[] outarray;

                try
                {
                    while (true)
                    {
                        while (usbi.TryReceive(out outarray) == 1) //grab all existing packets from buffer
                        {
                            CallBackMethod(outarray);
                        }
                        Thread.Sleep(10); //Let's conserve a little bit of CPU time...
                    }
                }
                catch (ThreadAbortException exc)
                {
                    usbi.Dispose();
                }
            }
        }
    }    
}
