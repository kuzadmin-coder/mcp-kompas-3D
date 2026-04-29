using Kompas6API5;

using System;
using System.IO;
using Microsoft.Win32;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using KAPITypes;

namespace Steps.NET
{
	// Класс Step1 - Самая простая библиотека на C#
  [ClassInterface(ClassInterfaceType.AutoDual)]
 	public class Step1
	{
		// Имя библиотеки
		[return: MarshalAs(UnmanagedType.BStr)] public string GetLibraryName()
		{
			return "Step1 - Самая простая библиотека на C#";
		}
		

		// Головная функция библиотеки
		public void ExternalRunCommand([In] short command, [In] short mode, [In, MarshalAs(UnmanagedType.IDispatch)] object kompas_)
		{
			KompasObject kompas = (KompasObject) kompas_;
			kompas.ksMessage("Привет!");
		}

    // Номер ключа защиты  short  или int16
    public short GetProtectNumber()
    {
      return 111;
    }

    #region COM Registration
    // Эта функция выполняется при регистрации класса для COM
    // Она добавляет в ветку реестра компонента раздел Kompas_Library,
    // который сигнализирует о том, что класс является приложением Компас,
    // а также заменяет имя InprocServer32 на полное, с указанием пути.
    // Все это делается для того, чтобы иметь возможность подключить
    // библиотеку на вкладке ActiveX.
    [ComRegisterFunction]
		public static void RegisterKompasLib(Type t)
		{
			try
			{
				RegistryKey regKey = Registry.LocalMachine;
				string keyName = @"SOFTWARE\Classes\CLSID\{" + t.GUID.ToString() + "}";
				regKey = regKey.OpenSubKey(keyName, true);
				regKey.CreateSubKey("Kompas_Library");
				regKey = regKey.OpenSubKey("InprocServer32", true);
				regKey.SetValue(null, System.Environment.GetFolderPath(Environment.SpecialFolder.System) + @"\mscoree.dll");
				regKey.Close();
			}
			catch (Exception ex)
			{
				MessageBox.Show(string.Format("При регистрации класса для COM-Interop произошла ошибка:\n{0}", ex));
			}
		}
		
		// Эта функция удаляет раздел Kompas_Library из реестра
		[ComUnregisterFunction]
		public static void UnregisterKompasLib(Type t)
		{
			RegistryKey regKey = Registry.LocalMachine;
			string keyName = @"SOFTWARE\Classes\CLSID\{" + t.GUID.ToString() + "}";
			RegistryKey subKey = regKey.OpenSubKey(keyName, true);
			subKey.DeleteSubKey("Kompas_Library");
			subKey.Close();
		}
		#endregion
	}
}
