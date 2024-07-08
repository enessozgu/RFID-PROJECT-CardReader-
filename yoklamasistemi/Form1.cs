using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using OfficeOpenXml;

namespace yoklamasistemi
{
	public partial class Form1 : Form
	{
		[DllImport("winscard.dll")]
		public static extern int SCardEstablishContext(uint dwScope, IntPtr pvReserved1, IntPtr pvReserved2, out IntPtr phContext);

		[DllImport("winscard.dll")]
		public static extern int SCardReleaseContext(IntPtr hContext);

		[DllImport("winscard.dll")]
		public static extern int SCardListReaders(IntPtr hContext, string mszGroups, byte[] mszReaders, ref int pcchReaders);

		[DllImport("winscard.dll")]
		public static extern int SCardConnect(IntPtr hContext, string szReader, uint dwShareMode, uint dwPreferredProtocols, out IntPtr phCard, out IntPtr pdwActiveProtocol);

		[DllImport("winscard.dll")]
		public static extern int SCardDisconnect(IntPtr hCard, int dwDisposition);

		[DllImport("winscard.dll")]
		public static extern int SCardTransmit(IntPtr hCard, IntPtr pioSendPci, byte[] pbSendBuffer, int cbSendLength, IntPtr pioRecvPci, byte[] pbRecvBuffer, ref int pcbRecvLength);

		private IntPtr hContext;
		private IntPtr hCard;

		private static string excelFilePath = @"C:\Users\Enes\asas.xlsx";

		public Form1()
		{
			InitializeComponent();
		}

		private void Form1_Load(object sender, EventArgs e)
		{
			InitializeReader();
		}

		public static class ExcelHelper
		{
			static ExcelHelper()
			{
				// LicenseContext'ı belirterek EPPlus lisans bağlamını ayarlayın
				ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Ticari olmayan kullanım için
			}

			public static void AddIDToExcel(string id)
			{
				try
				{
					FileInfo file = new FileInfo(excelFilePath);
					if (!file.Exists)
					{
						using (ExcelPackage package = new ExcelPackage(file))
						{
							package.Save();
						}
					}

					using (ExcelPackage package = new ExcelPackage(file))
					{
						ExcelWorksheet worksheet = package.Workbook.Worksheets.Count == 0
							? package.Workbook.Worksheets.Add("Kimlikler")
							: package.Workbook.Worksheets[0];

						int rowCount = worksheet.Dimension?.Rows ?? 0;
						worksheet.Cells[rowCount + 1, 1].Value = id;

						package.Save();
					}
				}
				catch (IOException ex)
				{
					MessageBox.Show($"Excel dosyasına yazma hatası: {ex.Message}\n\nDosya başka bir program tarafından kullanılıyor olabilir. Lütfen dosyayı kapatıp tekrar deneyin.");
				}
				catch (Exception ex)
				{
					MessageBox.Show($"Excel dosyasına yazma hatası: {ex.Message}");
				}
			}
		}

		private void InitializeReader()
		{
			int result = SCardEstablishContext(2, IntPtr.Zero, IntPtr.Zero, out hContext);
			if (result != 0)
			{
				MessageBox.Show($"Okuyucu bağlantı hatası, hata kodu: {result}");
				return;
			}

			int pcchReaders = 0;
			result = SCardListReaders(hContext, null, null, ref pcchReaders);
			if (result != 0)
			{
				MessageBox.Show($"Okuyucu listeleme hatası, hata kodu: {result}");
				return;
			}

			byte[] readersList = new byte[pcchReaders];
			result = SCardListReaders(hContext, null, readersList, ref pcchReaders);
			if (result != 0)
			{
				MessageBox.Show($"Okuyucu listeleme hatası, hata kodu: {result}");
				return;
			}

			string readerName = System.Text.Encoding.ASCII.GetString(readersList).TrimEnd('\0');
			MessageBox.Show($"Bulunan okuyucu: {readerName}");

			IntPtr activeProtocol;
			result = SCardConnect(hContext, readerName, 2, 3, out hCard, out activeProtocol);
			if (result != 0)
			{
				MessageBox.Show($"Okuyucuya bağlantı hatası, hata kodu: {result}");
				return;
			}

			MessageBox.Show("Okuyucuya başarıyla bağlanıldı!");
		}

		private void buttonReadCard_Click(object sender, EventArgs e)
		{
			byte[] receiveBuffer = new byte[256];
			int receiveLength = receiveBuffer.Length;
			byte[] sendBuffer = new byte[] { 0xFF, 0xCA, 0x00, 0x00, 0x00 };

			int result = SCardTransmit(hCard, IntPtr.Zero, sendBuffer, sendBuffer.Length, IntPtr.Zero, receiveBuffer, ref receiveLength);
			if (result == 0)
			{
				string cardUID = BitConverter.ToString(receiveBuffer, 0, receiveLength).Replace("-", "");
				if (!listBox1.Items.Contains(cardUID))
				{
					listBox1.Items.Add(cardUID);
					ExcelHelper.AddIDToExcel(cardUID);
				}
			}
			else
			{
				MessageBox.Show($"Kart okuma hatası, hata kodu: {result}");
			}
		}

		private void buttonAddID_Click(object sender, EventArgs e)
		{
			string inputID = textBox1.Text.Trim();
			if (!string.IsNullOrEmpty(inputID) && !listBox1.Items.Contains(inputID))
			{
				listBox1.Items.Add(inputID);
				ExcelHelper.AddIDToExcel(inputID);
			}
			textBox1.Clear();
		}
	}
}
