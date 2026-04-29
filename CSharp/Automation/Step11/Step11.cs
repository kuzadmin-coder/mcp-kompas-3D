using Kompas6API5;

using System;
using Microsoft.Win32;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Kompas6Constants;
using KAPITypes;

using reference = System.Int32;

namespace Steps.NET
{
	// Класс Step11 - Пример запроса командного окна
  [ClassInterface(ClassInterfaceType.AutoDual)]
	public class Step11
	{
		private KompasObject kompas;

		// Имя библиотеки
		[return: MarshalAs(UnmanagedType.BStr)] public object GetLibraryName()
		{
			return "Step11 - Пример запроса командного окна";
		}
		

		// Головная функция библиотеки
		public void ExternalRunCommand([In] short command, [In] short mode, [In, MarshalAs(UnmanagedType.IDispatch)] object kompas_)
		{
			kompas = (KompasObject) kompas_;
			if (kompas != null)
			{
				ksDocument2D doc = (ksDocument2D)kompas.Document2D();
				ksRequestInfo info = (ksRequestInfo)kompas.GetParamStruct((short)StructType2DEnum.ko_RequestInfo);
				if (doc != null && info != null) 
				{
					info.Init();
					info.menuId = 3002;
					info.title = "Дерево команд";

					info.SetCallBackCm("CALLBACKPROCCOMMANDWINDOW", 0, this);
					int cmd = doc.ksCommandWindow(info);
					kompas.ksMessage(string.Format("Выбрана команда {0}", cmd));
				}
			}
		}


		// Формирование меню библиотеки
		[return: MarshalAs(UnmanagedType.BStr)]	public string ExternalMenuItem1(short number, ref short itemType, ref short command)
		{
			string result = string.Empty;	//По уполчанию - пустая строка
			itemType = 1;					//MENUITEM

			switch (number)
			{
				case 1:
					result = "Пример запроса командного окна";
					command = 1;
					break;
				case 2:
					itemType = 3;			//ENDMENU
					command = -1;
					break;
			}

			return result;
		}
	

		// HINSTANCE модуля с ресурсами
		public object ExternalGetResourceModule()
		{
			return Assembly.GetExecutingAssembly().Location;
		}


		// Функция обратной связи
		// 0 - плоскость, 1 - цилиндр, 2 - ось - состав collection
		public int CALLBACKPROCCOMMANDWINDOW(int comm, [In][MarshalAs(UnmanagedType.LPStruct)]object rInfo)
		{
			kompas.ksMessage(string.Format("Выполняется команда {0}", comm));
			// заменяем заголовок окна в зависимости от выполненной команды
			// аналогично можно заменить и состав дерева команд
			ksRequestInfo info = null;
			info = (ksRequestInfo)rInfo;
			if (info != null) 
			{
				switch (comm) 
				{
					case 1	: info.title = "1"; break;
					default	: info.title = "2"; break;
				}
			}
			// возвращаемый результат определяет, должна ли система продолжать
			// запрашивать команду :
			// TRUE - продолжать
			// FALSE - завержить работу с окном
			return comm == 2213 ? 0 : 1;
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
