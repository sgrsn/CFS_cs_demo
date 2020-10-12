using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using System.IO.Ports;
using System.Timers;

using System.Runtime.InteropServices; // DLL Import
class CFSAPI
{
    [DllImport("CfsUsb.dll", EntryPoint = "Initialize")]
    public static extern void Initialize();

	[DllImport("CfsUsb.dll", EntryPoint = "Finalize")]
	public static extern void func_Finalize();

    [DllImport("CfsUsb.dll", EntryPoint = "PortOpen")]
    public static extern bool PortOpen(int portNo);

    [DllImport("CfsUsb.dll", EntryPoint = "PortClose")]
    public static extern void PortClose(int portNo);

    [DllImport("CfsUsb.dll", EntryPoint = "SetSerialMode")]
    public static extern bool SetSerialMode(int portNo, bool mode);

    [DllImport("CfsUsb.dll", EntryPoint = "GetSerialData")]
    public static extern bool GetSerialData(int portNo, out double data, out char status);

    [DllImport("CfsUsb.dll", EntryPoint = "GetLatestData")]
    public static extern bool GetLatestData(int portNo, out double data, out char status);

    [DllImport("CfsUsb.dll", EntryPoint = "GetSensorLimit")]
    public static extern bool GetSensorLimit(int portNo, out double Limit);

    [DllImport("CfsUsb.dll", EntryPoint = "GetSensorInfo")]
    public static extern bool GetSensorInfo(int portNo, out char SerialNo);
}

namespace CFS_cs_demo
{
	/// <summary>
	/// MainWindow.xaml の相互作用ロジック
	/// </summary>
	/// 
	public partial class MainWindow : Window
    {
		private long cnt;
		private int portNo = 7;
		private char[] SerialNo = new char[9];
		private char Status;
		private double[] Limit = new double[6];
		private double[] Data = new double[6];
		private double Fx, Fy, Fz, Mx, My, Mz;

		private SerialPort port;
		const byte DLE = 0x10;
		const byte STX = 0x02;
		const byte ETX = 0x03;
		const byte NAK = 0x15;

		private System.Timers.Timer update_timer;
		private const int tick_receive = 1;
		
		private void button_Click(object sender, RoutedEventArgs e)
        {
			port.Open();
			port.DtrEnable = true;
			port.RtsEnable = true;
			//port.DataReceived += new SerialDataReceivedEventHandler(aDataReceivedHandler);

			Console.WriteLine("Connected");

			GetCFSLimit();
		}

		private void button1_Click(object sender, RoutedEventArgs e)
        {
			//GetCFSData();
			//StartCFSData();

			update_timer = new System.Timers.Timer(tick_receive);
			update_timer.Elapsed += update;
			update_timer.AutoReset = true;
			update_timer.Enabled = true;
		}

		private void button2_Click(object sender, RoutedEventArgs e)
		{
			//StopCFSData();

			update_timer.Stop();
			update_timer.Dispose();
		}

		private void update(Object source, ElapsedEventArgs e)
        {
			GetCFSData();
		}

		public MainWindow()
		{
			InitializeComponent();
			//demo();

			port = new SerialPort("COM7", 460800, Parity.None, 8, StopBits.One);

			Console.WriteLine("Start");
		}

		private void GetCFSInformation()
		{
			byte[] command = { 0x04, 0xFF, 0x2A, 0x00 };
			byte[] read_buffer = new byte[45];

			Command2CFS(command, read_buffer);

			for (int i = 0; i < read_buffer.Length; i++)
			{
				Console.WriteLine(read_buffer[i]);
			}
		}

		private void GetCFSLimit()
		{
			byte[] command = { 0x04, 0xFF, 0x2B, 0x00 };
			byte[] read_buffer = new byte[32];
			int result = Command2CFS(command, read_buffer);
			/*for (int i = 0; i < read_buffer.Length; i++)
			{
				Console.WriteLine(read_buffer[i]);
			}*/
			if (result == 0)
			{
				Limit[0] = BitConverter.ToSingle(read_buffer, 4);
				Limit[1] = BitConverter.ToSingle(read_buffer, 8);
				Limit[2] = BitConverter.ToSingle(read_buffer, 12);
				Limit[3] = BitConverter.ToSingle(read_buffer, 16);
				Limit[4] = BitConverter.ToSingle(read_buffer, 20);
				Limit[5] = BitConverter.ToSingle(read_buffer, 24);
				Console.WriteLine("LimitFx:{0}, LimitFy:{1}, LimitFz:{2}", Limit[0], Limit[1], Limit[2]);
				Console.WriteLine("LimitMx:{0}, LimitMy:{1}, LimitMz:{2}", Limit[3], Limit[4], Limit[5]);
			}
		}

		private void GetCFSData()
        {
			byte[] command = { 0x04, 0xFF, 0x30, 0x00 };
			byte[] read_buffer = new byte[30];
			int result = Command2CFS(command, read_buffer);
			/*for (int i = 0; i < read_buffer.Length; i++)
			{
				Console.WriteLine(read_buffer[i]);
			}*/
			if (result == 0)
			{
                int fx = BitConverter.ToInt16(read_buffer, 4);
				int fy = BitConverter.ToInt16(read_buffer, 6);
				int fz = BitConverter.ToInt16(read_buffer, 8);
				int mx = BitConverter.ToInt16(read_buffer, 10);
				int my = BitConverter.ToInt16(read_buffer, 12);
				int mz = BitConverter.ToInt16(read_buffer, 14);

				Fx = Limit[0] / 10000 * fx;
				Fy = Limit[1] / 10000 * fy;
				Fz = Limit[2] / 10000 * fz;
				Mx = Limit[0] / 10000 * mx;
				My = Limit[1] / 10000 * my;
				Mz = Limit[2] / 10000 * mz;

				Console.WriteLine("Fx:{0}, Fy:{1}, Fz:{2}", Fx, Fy, Fz);
				Console.WriteLine("Mx:{0}, My:{1}, Mz:{2}", Mx, My, Mz);
			}
		}

		private void StartCFSData()
		{
			byte[] command = { 0x04, 0xFF, 0x32, 0x00 };
			byte[] read_buffer = new byte[10];
			int result = Command2CFS(command, read_buffer);
			for (int i = 0; i < read_buffer.Length; i++)
			{
				Console.WriteLine(read_buffer[i]);
			}
		}

		private void StopCFSData()
		{
			byte[] command = { 0x04, 0xFF, 0x33, 0x00 };
			byte[] read_buffer = new byte[10];
			int result = Command2CFS(command, read_buffer);
			for (int i = 0; i < read_buffer.Length; i++)
			{
				Console.WriteLine(read_buffer[i]);
			}
		}

        private void GetCFSDataUntilStop()
        {
			byte[] read_buffer = new byte[22];
			int result = ReadCFS(read_buffer);

			if (result == 0)
			{
				int fx = BitConverter.ToInt16(read_buffer, 4);
				int fy = BitConverter.ToInt16(read_buffer, 6);
				int fz = BitConverter.ToInt16(read_buffer, 8);
				int mx = BitConverter.ToInt16(read_buffer, 10);
				int my = BitConverter.ToInt16(read_buffer, 12);
				int mz = BitConverter.ToInt16(read_buffer, 14);

				Fx = Limit[0] / 10000 * fx;
				Fy = Limit[1] / 10000 * fy;
				Fz = Limit[2] / 10000 * fz;
				Mx = Limit[0] / 10000 * mx;
				My = Limit[1] / 10000 * my;
				Mz = Limit[2] / 10000 * mz;

				Console.WriteLine("Fx:{0}, Fy:{1}, Fz:{2}", Fx, Fy, Fz);
				Console.WriteLine("Mx:{0}, My:{1}, Mz:{2}", Mx, My, Mz);
			}
		}

		private byte CalculateBCC(byte[] command)
		{
			byte BCC = 0;
			for (int i = 0; i < command.Length; i++)
			{
				BCC = (byte)(BCC ^ command[i]);
			}
			BCC = (byte)(BCC ^ ETX);    // ETX 直前の DLE を含まない
			return BCC;
		}

		private int Command2CFS(byte[] command, byte[] read_buffer)
        {
			byte BCC = CalculateBCC(command);
			byte[] START = { DLE, STX };
			byte[] END = { DLE, ETX, BCC };
			byte[] buffer = new byte[START.Length + command.Length + END.Length];
			for(int i = 0; i < START.Length; i++)
            {
				buffer[i] = START[i];
            }
			for (int i = 0; i < command.Length; i++)
			{
				buffer[START.Length + i] = command[i];
			}
			for (int i = 0; i < END.Length; i++)
			{
				buffer[START.Length + command.Length + i] = END[i];
			}

			port.Write(buffer, 0, buffer.Length);
			/*for(int i = 0; i < buffer.Length; i++)
            {
				WriteOneByteData(buffer[i]);
			}*/

			//port.Read(read_buffer, 0, read_buffer.Length);

			int result = ReadCFS(read_buffer);

			return result;
		}

		// BCC計算対象範囲まで格納
		private int ReadCFS(byte[] read_buffer)
        {
			int read_byte = port.ReadByte();

			// 最初がDLEでなかったら失敗
			if (read_byte != DLE)
            {
				Console.WriteLine("received {0}, non HEAD byte:{1}", read_byte, DLE);
				return -1;
            }

			else
			{
				read_byte = port.ReadByte();
				if (read_byte == NAK)
				{
					// 否定応答
					Console.WriteLine("received NAK:{0}", NAK);
					return -1;
				}
                else if(read_byte == STX)
                {
					byte BCC = 0;
					int cnt = 0;
					while(true)
                    {
						read_byte = (byte)port.ReadByte();
						if(read_byte == DLE)
                        {
							read_byte = (byte)port.ReadByte();
							if(read_byte == ETX)
							{
								read_buffer[cnt] = (byte)read_byte;
								BCC = (byte)(BCC ^ read_buffer[cnt]);
								byte read_bcc = (byte)port.ReadByte();
								if (BCC == read_bcc)
								{
									return 0;
								}
								else
								{
									// BCCエラー
									Console.WriteLine("BCC eroor, received {0}, but in the calculations {1}", read_bcc, BCC);
									return -1;
								}
							}
							else if(read_byte == 0x10)
                            {
								read_buffer[cnt] = (byte)read_byte;
								BCC = (byte)(BCC ^ read_buffer[cnt]);
								cnt++;
							}
						}
                        else
                        {
							read_buffer[cnt] = (byte)read_byte;
							BCC = (byte)(BCC ^ read_buffer[cnt]);
							cnt++;
						}
					}
				}
                else
                {
					// ここはもうありえないけど
					Console.WriteLine("non STX, received {0}", read_byte);
					return -1;
                }
			}

			return 0;
		}

		public void demo()
        {
			// ＤＬＬの初期化処理
			CFSAPI.Initialize();

			// ポートオープン
			if (CFSAPI.PortOpen(portNo) == true)
			{
				// センサ定格確認
				if (CFSAPI.GetSensorLimit(portNo, out Limit[0]) == false)
				{
					Console.WriteLine("センサ定格確認ができません。");
				}
				// シリアルNo確認
				if (CFSAPI.GetSensorInfo(portNo, out SerialNo[0]) == false)
				{
					Console.WriteLine("シリアルNoが取得できません。");
				}
				// ハンドシェイクによる読込
				// 最新データ読込
				// ※センサからは定格を10000としてデータが出力されてくる
				if (CFSAPI.GetLatestData(portNo, out Data[0], out Status) == true)
				{
					Fx = Limit[0] / 10000 * Data[0];                                // Fxの値
					Fy = Limit[1] / 10000 * Data[1];                                // Fyの値
					Fz = Limit[2] / 10000 * Data[2];                                // Fzの値
					Mx = Limit[3] / 10000 * Data[3];                                // Mxの値
					My = Limit[4] / 10000 * Data[4];                                // Myの値
					Mz = Limit[5] / 10000 * Data[5];                                // Mzの値

					Console.WriteLine("GetLastData\n");
					Console.WriteLine("Fx:{0:f} Fy:{1:f} Fz:{2:f} Mx:{3:f} My:{4:f} Mz:{5:f}\n", Fx, Fy, Fz, Mx, My, Mz);
				}
				else
				{
					Console.WriteLine("最新データ取得に失敗しました。");
				}

				// 連続読込 
				// 連続データ読込モードに移行
				if (CFSAPI.SetSerialMode(portNo, true) == true)
				{
					// 10000回連続読込
					cnt = 0;
					Console.WriteLine("GetSerialData\n");
					while (cnt < 1000)
					{
						// データ取得
						if (CFSAPI.GetSerialData(portNo, out Data[0], out Status) == true)
						{
							Fx = Limit[0] / 10000 * Data[0];                        // Fxの値
							Fy = Limit[1] / 10000 * Data[1];                        // Fyの値
							Fz = Limit[2] / 10000 * Data[2];                        // Fzの値
							Mx = Limit[3] / 10000 * Data[3];                        // Mxの値
							My = Limit[4] / 10000 * Data[4];                        // Myの値
							Mz = Limit[5] / 10000 * Data[5];                        // Mzの値

							cnt++;

							Console.WriteLine("LimitFx:{0:f} LimitFy:{1:f} LimitFz:{2:f}\n", Limit[0], Limit[1], Limit[2]);
							Console.WriteLine("Fx:{0:f} Fy:{1:f} Fz:{2:f} Mx:{3:f} My:{4:f} Mz:{5:f}             \r", Fx, Fy, Fz, Mx, My, Mz);

						}
					}

					// 連続データ読込モードを停止
					if (CFSAPI.SetSerialMode(portNo, false) == false)
					{
						Console.WriteLine("連続読込モードを停止できません。");
					}
				}
				else
				{
					Console.WriteLine("連続読込モードに移行できません。");
				}

				// ポートクローズ
				CFSAPI.PortClose(portNo);
			}
			else
			{
				Console.WriteLine("回線がオープンできません。");
			}

			// ＤＬＬの終了処理
			//CFSAPI.func_Finalize();

			// ＤＬＬの解放
			//FreeLibrary(hDll);

			Console.WriteLine("\n完了");
			Console.ReadLine();
		}
    }
}
