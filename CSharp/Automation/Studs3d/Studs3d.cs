
using Kompas6API5;
using KompasAPI7;


using System;
using System.Resources;
using Microsoft.Win32;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Kompas6Constants;

namespace Steps.NET
{
	// Класс Studs3d - Крепежный элемент на C#
  [ClassInterface(ClassInterfaceType.AutoDual)]
	public class Studs3d
	{
		// Имя библиотеки
		[return: MarshalAs(UnmanagedType.BStr)] public string GetLibraryName()
		{
			return "Studs3d - Крепежный элемент на C#";
		}
		

		// Головная функция библиотеки
		public void ExternalRunCommand([In] short command, [In] short mode, [In, MarshalAs(UnmanagedType.IDispatch)] object kompas_)
		{
			Kompas.Instance.KompasObject = (KompasObject)kompas_;
			Kompas.Instance.Document3D = (ksDocument3D)Kompas.Instance.KompasObject.ActiveDocument3D();
			if (Kompas.Instance.KompasObject != null 
				&& Kompas.Instance.Document3D != null 
				&& !Kompas.Instance.Document3D.IsDetail())
			{
				Pin par = new Pin();
				if (par != null)
					par.Draw3D(Kompas.Instance.Document3D);
				else
				{
					ResourceManager rm = new ResourceManager(this.GetType());
					Kompas.Instance.KompasObject.ksError(rm.GetString("IDS_3DDOCERROR"));
				}	
				if (Kompas.Instance.KompasObject.ksReturnResult() == (int)ErrorType.etError10)	//  10  "Ошибка! Вырожденный объект"
					Kompas.Instance.KompasObject.ksResultNULL();
			}
		}


		// Формирование меню библиотеки
		[return: MarshalAs(UnmanagedType.BStr)] public string ExternalMenuItem(short number, ref short itemType, ref short command)
		{
			string result = string.Empty;
			itemType = 1; // "MENUITEM"
			switch (number)
			{
				case 1:
					result = "Шпильки";
					command = 1;
					break;
				case 2:
					command = -1;
					itemType = 3; // "ENDMENU"
					break;
			}
			return result;
		}


		// HMODULE ресурсного модуля
		public IntPtr ExternalGetResourceModule()
		{
			return Marshal.GetHINSTANCE(Assembly.GetExecutingAssembly().GetLoadedModules()[0]);;
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
