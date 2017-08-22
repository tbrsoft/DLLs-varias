using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;

using PVOID = System.IntPtr;
using DWORD = System.UInt32;

namespace usb_api
{
	unsafe public class usb_interface
	{
		#region  String Definitions of Pipes and VID_PID
		public string vid_pid_boot= "vid_04d8&pid_000b";    // Bootloader vid_pid ID
		public string vid_pid_norm= "vid_04d8&pid_000c";

		protected string out_pipe= "\\MCHP_EP1"; // Define End Points
		protected string in_pipe= "\\MCHP_EP1";
		#endregion

		#region Imported DLL functions from mpusbapi.dll
		[DllImport("mpusbapi.dll")]
		protected static extern DWORD _MPUSBGetDLLVersion();
		[DllImport("mpusbapi.dll")]
		protected static extern DWORD _MPUSBGetDeviceCount(string pVID_PID);
		[DllImport("mpusbapi.dll")]
		protected static extern void* _MPUSBOpen(DWORD instance,string pVID_PID,string pEP,DWORD dwDir,DWORD dwReserved);
		[DllImport("mpusbapi.dll")]
		protected static extern DWORD _MPUSBRead(void* handle,void* pData,DWORD dwLen,DWORD* pLength,DWORD dwMilliseconds);
		[DllImport("mpusbapi.dll")]
        protected static extern DWORD _MPUSBWrite(void* handle, void* pData, DWORD dwLen, DWORD* pLength, DWORD dwMilliseconds);
		[DllImport("mpusbapi.dll")]
        protected static extern DWORD _MPUSBReadInt(void* handle, void* pData, DWORD dwLen, DWORD* pLength, DWORD dwMilliseconds);
		[DllImport("mpusbapi.dll")]
        protected static extern bool _MPUSBClose(void* handle);
		#endregion

		protected void* myOutPipe;
		protected void* myInPipe;

        public enum Commands : byte
        {
            UPDATE_LED = 0x32,
            RESET = 0xFF
        }

		protected virtual void OpenPipes()
		{
			DWORD selection=0; // Selects the device to connect to, in this example it is assumed you will only have one device per vid_pid connected.

			myOutPipe = _MPUSBOpen(selection,vid_pid_norm,out_pipe,0,0);
			myInPipe = _MPUSBOpen(selection,vid_pid_norm,in_pipe,1,0);
		}
		protected void ClosePipes()
		{
			_MPUSBClose(myOutPipe);
			_MPUSBClose(myInPipe);
		}

		protected virtual DWORD SendReceivePacket(byte* SendData, DWORD SendLength, byte* ReceiveData, DWORD *ReceiveLength)
		{
			uint SendDelay=1000;
			uint ReceiveDelay=1000;

			DWORD SentDataLength;
			DWORD ExpectedReceiveLength = *ReceiveLength;

			OpenPipes();

				if(_MPUSBWrite(myOutPipe,(void*)SendData,SendLength,&SentDataLength,SendDelay)==1)
				{

					if(_MPUSBRead(myInPipe,(void*)ReceiveData, ExpectedReceiveLength,ReceiveLength,ReceiveDelay)==1)
					{
						if(*ReceiveLength == ExpectedReceiveLength)
						{
							ClosePipes();
							return 1;   // Success!
						}
						else if(*ReceiveLength < ExpectedReceiveLength)
						{
							ClosePipes();
							return 2;   // Partially failed, incorrect receive length
						}
					}
				}
			ClosePipes();
			return 0;  // Operation Failed
		}

		public DWORD GetDLLVersion()
		{
			return _MPUSBGetDLLVersion();
		}
		public DWORD GetDeviceCount(string Vid_Pid)
		{
			return _MPUSBGetDeviceCount(Vid_Pid);
		}

		public int UpdateLED(uint led, bool State)
		{
			// The default demo firmware application has a defined application
			// level protocol.
			// To set the LED's, the host must send the UPDATE_LED
			// command which is defined as 0x32, followed by the LED to update,
			// then the state to chance the LED to.
			//
			// i.e. <UPDATE_LED><0x01><0x01>
			//
			// Would activate LED 1
			//
			// The receive buffer size must be equal to or larger than the maximum
			// endpoint size it is communicating with. In this case, it is set to 64 bytes.

			byte* send_buf=stackalloc byte[64];
			byte* receive_buf=stackalloc byte[64];

			DWORD RecvLength=3;
			send_buf[0] = 0x32;			//Command for LED Status  
			send_buf[1] = (byte)led;
			send_buf[2] = (byte)(State?1:0);
	
			if (SendReceivePacket(send_buf,3,receive_buf,&RecvLength) == 1)
			{
				if (RecvLength == 1 && receive_buf[0] == 0x32)
				{	
					return 0;
				}
				else
				{
					return 2;
				}
			}
			else
			{	
				return 1;
			}
		}

        /// <summary>
        /// Simple method for sending/receiving USB data from the PIC.
        /// </summary>
        /// <param name="Command">Command byte</param>
        /// <param name="rxlength">Expected length of received packet</param>
        /// <param name="data">Array of bytes to send after command - pass null if not applicable</param>
        /// <param name="dataout">Array of received bytes</param>
        /// <returns></returns>
        public uint EasyCommand(byte Command, int rxlength, byte[] data, out byte[] dataout)
        {
            byte* send_buf = stackalloc byte[64];
            byte* receive_buf = stackalloc byte[64];
            DWORD RecvLength;
            uint rval;

            RecvLength = (DWORD)rxlength;
            send_buf[0] = Command;
            if (data != null)
            {
                for (int i = 0; i < data.Length; i++)
                { send_buf[i + 1] = data[i]; }
            }
            else
                data = new byte[] { }; //just set to empty array so .Length member is valid at 0.
            rval = SendReceivePacket(send_buf, 1 + (uint)data.Length, receive_buf, &RecvLength);
            if (rval != 1)
            { dataout = null; return rval; }
            dataout = new byte[rxlength];
            for (int i = 0; i < rxlength; i++)
            { dataout[i] = receive_buf[i]; }
            return 1;
        }
	}
}
