
using Kompas6API5;
using KompasAPI7;

using System;
using System.IO;
using Microsoft.Win32;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Steps.NET
{
	// Класс Gayka - Конструкторский элемент
	public class Gayka
	{
		// Имя библиотеки
		[return: MarshalAs(UnmanagedType.BStr)] public string GetLibraryName()
		{
			return "Gayka - Конструкторский элемент";
		}
		

		// Головная функция библиотеки
		public void ExternalRunCommand([In] short command, [In] short mode, [In, MarshalAs(UnmanagedType.IDispatch)] object kompas_)
		{
			Kompas.Instance.KompasObject = (KompasObject)kompas_;
			Kompas.Instance.Document2D = (ksDocument2D)Kompas.Instance.KompasObject.ActiveDocument2D();
			Kompas.Instance.KompasApp = (_Application)Kompas.Instance.KompasObject.ksGetApplication7();

			switch (command)
			{
				case 1:
					GaykaObj gayka = new GaykaObj();
					gayka.Draw();
					break;
			}
		}

    //-------------------------------------------------------------------------------
    // Получить идентификаторы инструментальных и компактных панелей
    // ---
    public bool FillContextPanel([In, MarshalAs(UnmanagedType.IDispatch)] object _contextPanel)  // Индекс панели
    {
      bool res = false;
			IContextPanel contextPanel = (IContextPanel)_contextPanel;
      if (contextPanel != null)
      {
        contextPanel.Fill("GAYKA5915_CSharp_BAR");
        res = true;
      }
      return res;
    }



    // Формирование меню библиотеки
    [return: MarshalAs(UnmanagedType.BStr)] public string ExternalMenuItem(short number, ref short itemType, ref short command)
		{
			string result = string.Empty;
			itemType = 1; // "MENUITEM"
			switch (number)
			{
				case 1:
					result = "Гайки ГОСТ 5915-70";
					command = 1;
					break;
				case 2:
					command = -1;
					itemType = 3; // "ENDMENU"
					break;
			}
			return result;
		}


		// HINSTANCE модуля с ресурсами
		public object ExternalGetResourceModule()
		{
			return Assembly.GetExecutingAssembly().Location;
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
