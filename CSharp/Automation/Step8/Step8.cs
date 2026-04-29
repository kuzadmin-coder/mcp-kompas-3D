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
	// Класс Step8 - Работа с атрибутами
  [ClassInterface(ClassInterfaceType.AutoDual)]
	public class Step8
	{
		private KompasObject kompas;
		private ksDocument2D doc;
		private ksAttributeObject attr;

		// Имя библиотеки
		[return: MarshalAs(UnmanagedType.BStr)] public string GetLibraryName()
		{
			return "Step8 - Работа с атрибутами";
		}
		

		// Головная функция библиотеки
		public void ExternalRunCommand([In] short command, [In] short mode, [In, MarshalAs(UnmanagedType.IDispatch)] object kompas_)
		{
			kompas = (KompasObject) kompas_;
			
			if (kompas  != null)
			{
				doc = (ksDocument2D)kompas.ActiveDocument2D();
				attr = (ksAttributeObject)kompas.GetAttributeObject();
				if (doc  != null && doc.reference  != 0 && attr  != null)
				{
					switch (command)
					{
						case 1  : FuncAttrType();			break; // создать тип атрибута
						case 2  : DelTypeAttr();			break; // удалить  тип атрибута
						case 3  : ShowTypeAttr();			break; // получить  тип атрибута
						case 4  : ChangeType();				break; // заменить  тип атрибута
						case 5  : NewAttr();				break; // создать атрибут определенного типа
						case 6  : DelObjAttr();				break; // удалить атрибут
						case 7  : ReadObjAttr();			break; // считать атрибут
						case 8  : ShowObjAttr();			break; // просмотреть атрибут
						case 9  : ShowLib();				break; // просмотреть библиотеку
						case 10 : ShowType();				break; // просмотреть тип
						case 11 : WalkFromObjWithAttr();	break; // просмотреть атрибут
					}
				}
			}
		}


		// Формирование меню библиотеки
		[return: MarshalAs(UnmanagedType.BStr)]	public string ExternalMenuItem(short number, ref short itemType, ref short command)
		{
			string result = string.Empty;	//По уполчанию - пустая строка
			itemType = 1;					//MENUITEM

			switch (number)
			{
				case 1:
					result = "Создать тип атрибута";
					command = 1;
					break;
				case 2:
					result = "Удалить тип атрибута";
					command = 2;
					break;
				case 3:
					result = "Получить тип атрибута";
					command = 3;
					break;
				case 4:
					result = "Заменить тип атрибута";
					command = 4;
					break;
				case 5:
					result = "Создать атрибут определенного типа";
					command = 5;
					break;
				case 6:
					result = "Удалить атрибут";
					command = 6;
					break;
				case 7:
					result = "Считать атрибут";
					command = 7;
					break;
				case 8:
					result = "Просмотреть атрибуты объекта";
					command = 8;
					break;
				case 9:
					result = "Просмотреть библиотеку";
					command = 9;
					break;
				case 10:
					result = "Просмотреть тип";
					command = 10;
					break;
				case 11:
					result = "Просмотреть атрибут";
					command = 11;
					break;
				case 12:
					itemType = 3;			//ENDMENU
					command = -1;
					break;
			}

			return result;
		}


		void ShowCol(ksColumnInfoParam par, int iCol, bool fl)
		{
			string buf = string.Empty;
			string s = string.Empty;;
			if (fl)
				s = "структура";

			//выдадим поля колонки не указатели
			buf = string.Format("{0} i = {1} header = {2} type = {3} def = {4} flagEnum = {5}",
				s, iCol, par.header, par.type, par.def, par.flagEnum);
			kompas.ksMessage(buf);
			if (par.type == ldefin2d.RECORD_ATTR_TYPE)
			{
				// структура
				ksDynamicArray pCol = (ksDynamicArray)par.GetColumns();
				if (pCol  != null)
				{
					ShowColumns(pCol, true);
					pCol.ksDeleteArray();
				}
			}
		}

		void ShowColumns(ksDynamicArray pCol, bool fl)
		{
			ksColumnInfoParam par = (ksColumnInfoParam)kompas.GetParamStruct((short)StructType2DEnum.ko_ColumnInfoParam);
			if (par != null) 
			{
				par.Init();
				int n = pCol.ksGetArrayCount();

				for (int i = 0; i < n; i++) 
				{
					if (pCol.ksGetArrayItem(i, par) != 1)
						kompas.ksMessageBoxResult();  // проверяем результат работы нашей функции
					else
						ShowCol(par,i, fl);
				}
			}
		}

		// Создание типа аттрибута
		void FuncAttrType() 
		{
			ksAttributeTypeParam type = (ksAttributeTypeParam)kompas.GetParamStruct((short)StructType2DEnum.ko_AttributeType);
			ksColumnInfoParam col = (ksColumnInfoParam)kompas.GetParamStruct((short)StructType2DEnum.ko_ColumnInfoParam);
			if (type != null && col != null)
			{
				type.Init();
				col.Init();
				type.header = "double_str_long";	// заголовoк-комментарий типа
				type.rowsCount = 1;					// кол-во строк в таблице
				type.flagVisible = true;            // видимый, невидимый   в таблице
				type.password = string.Empty;       // пароль, если не пустая строка  - защищает от несанкционированного изменения типа
				type.key1 = 10;
				type.key2 = 20;
				type.key3 = 30;
				type.key4 = 0;
				ksDynamicArray arr = (ksDynamicArray)type.GetColumns();
				if (arr != null)
				{
					// первая колонка  структуры
					col.header = "double";					// заголовoк-комментарий столбца
					col.type = ldefin2d.DOUBLE_ATTR_TYPE;	// тип данных в столбце - см.ниже
					col.key = 0;							// дополнительный признак, который позволит отличить две переменные с одинаковым типом
					col.def = "123456789";					// значение по умолчанию
					col.flagEnum = false;					// флаг включающий режим, когда значение поля атрибута
					arr.ksAddArrayItem(-1, col);

					// вторая колонка структуры
					col.header = "str";						// заголовoк-комментарий столбца
					col.type = ldefin2d.STRING_ATTR_TYPE;	// тип данных в столбце - см.ниже
					col.key = 0;							// дополнительный признак, который позволит отличить две переменные с одинаковым типом
					col.def = "string";						// значение по умолчанию
					col.flagEnum = false;					// флаг включающий режим, когда значение поля атрибута
					arr.ksAddArrayItem(-1, col);

					// третья колонка структуры
					col.header = "long";					// заголовoк-комментарий столбца
					col.type = ldefin2d.LINT_ATTR_TYPE;		// тип данных в столбце - см.ниже
					col.key = 0;							// дополнительный признак, который позволит отличить две переменные с одинаковым типом
					col.def = "1000000";					// значение по умолчанию
					col.flagEnum = false;					// флаг включающий режим, когда значение поля атрибута
					arr.ksAddArrayItem(-1, col);
				}
				string nameFile = string.Empty;
				//запросить имя библиотеки
				nameFile = kompas.ksChoiceFile("*.lat", null, false);
				//создать тип атрибута
				double numbType = attr.ksCreateAttrType(type,	// информация о типе атрибута
					nameFile);									// имя библиотеки типов атрибутов
				if (numbType > 1)
				{
					string buf = string.Format("numbType = {0}", numbType);
					kompas.ksMessage(buf);
				}
				else
					kompas.ksMessageBoxResult();  // проверяем результат работы нашей функции

				//удалим  массив колонок
				arr.ksDeleteArray();
			}
		}

		void DelTypeAttr()
		{
			double numb = 0;
			int j;
			string password = string.Empty;
			//запросить имя библиотеки
			string nameFile = string.Empty;
			nameFile = kompas.ksChoiceFile("*.lat", null, false);
			do
			{
				j = kompas.ksReadDouble("Ввести номер типа атрибута", 1000, 0, 1e12, ref numb);
				if (j != 0)
				{
					password = kompas.ksReadString("Ввести пароль типа атрибута", "");
					if (attr.ksDeleteAttrType(numb, nameFile, password) != 1)
						kompas.ksMessageBoxResult();  // проверяем результат работы нашей функции
				}
			}
			while(j != 0);
		}

		// Получить тип атрибута
		void ShowTypeAttr()
		{
			double numb = 0;
			//запросить имя библиотеки
			string nameFile = string.Empty;
			nameFile = kompas.ksChoiceFile("*.lat", null, false);

			ksAttributeTypeParam type = (ksAttributeTypeParam)kompas.GetParamStruct((short)StructType2DEnum.ko_AttributeType);
			if (type != null)
			{
				type.Init();
				string buf = string.Empty;
				//		attrType.columns = CreateArray(ATTR_COLUMN_ARR, 0);

				do
				{
					numb = attr.ksChoiceAttrTypes(nameFile);
					if (numb != 0)
					{
						if (attr.ksGetAttrType(numb, nameFile, type) != 1)
							kompas.ksMessageBoxResult();	// проверяем результат работы нашей функции
						else
						{
							buf = string.Format("key1 = {0} key2 = {1} key3 = {2} key4 = {3}",
								type.key1, type.key2, type.key3, type.key4);
							kompas.ksMessage(buf);
							buf = string.Format("header = {0} rowsCount = {1} flagVisible = {2} password = {3}",
								type.header, type.rowsCount, type.flagVisible, type.password);
							kompas.ksMessage(buf);
							ksDynamicArray pCol = (ksDynamicArray)type.GetColumns();
							if (pCol != null)
							{
								ShowColumns(pCol, false);
								pCol.ksDeleteArray();
							}
							//					ShowColumns(attrType.columns, FALSE); //пользовательская функция
						}
					}
				}
				while(numb != 0);
				//удалим  массив колонок
				//		DeleteArray(attrType.columns);
			}
		}

		// Заменить тип атрибута
		void ChangeType()
		{
			double numb = 0;
			string password = string.Empty;
			//запросить имя библиотеки
			string nameFile = string.Empty;
			nameFile = kompas.ksChoiceFile("*.lat", null, false);
			int j;

			ksAttributeTypeParam type = (ksAttributeTypeParam)kompas.GetParamStruct((short)StructType2DEnum.ko_AttributeType);
			if (type != null)
			{
				type.Init();
				do
				{
					j = kompas.ksReadDouble("Ввести номер типа атрибута", 1000, 0, 1e12, ref numb);
					if (j  != 0)
					{
						password = kompas.ksReadString("Ввести пароль типа атрибута", "");
						//считаем тип атрибута
						if (attr.ksGetAttrType(numb, nameFile, type)  != 1)
							kompas.ksMessageBoxResult();  // проверяем результат работы нашей функции
						else 
						{
							type.password = password;
							ksDynamicArray arr = (ksDynamicArray)type.GetColumns();
							ksColumnInfoParam par1 = (ksColumnInfoParam)kompas.GetParamStruct((short)StructType2DEnum.ko_ColumnInfoParam);
							ksColumnInfoParam parN = (ksColumnInfoParam)kompas.GetParamStruct((short)StructType2DEnum.ko_ColumnInfoParam);
							if (arr != null && par1 != null && parN != null)
							{
								par1.Init();
								parN.Init();
								// число колонок
								int n = arr.ksGetArrayCount();
								//считаем первую колонку
								arr.ksGetArrayItem(0, par1);
								//считаем последнюю колонку
								arr.ksGetArrayItem(n-1, parN);
								//заменим первую колонку
								arr.ksSetArrayItem(0, parN);
								//заменим последнюю колонку
								arr.ksSetArrayItem(n-1, par1);

								//заменим  тип атрибута  на новый
								double numbType = attr.ksSetAttrType(numb, nameFile, type, password);
								if (numbType > 1)
								{
									string buf = string.Empty;
									buf = string.Format("numbType = {0}", numbType);
									kompas.ksMessage(buf);
								}
								else
									kompas.ksMessageBoxResult();  // неудачное завершение - выдадим результат работы нашей функции
								arr.ksDeleteArray();
							}
						}
					}
				} while(j  != 0);
			}
		}

		// Создадим атрибут типа,полученного из функции FuncTypeAttr
		void NewAttr() 
		{
			ksAttributeParam attrPar = (ksAttributeParam)kompas.GetParamStruct((short)StructType2DEnum.ko_Attribute);
			ksUserParam usPar = (ksUserParam)kompas.GetParamStruct((short)StructType2DEnum.ko_UserParam); 
			ksDynamicArray fVisibl = (ksDynamicArray)kompas.GetDynamicArray(23);
			ksDynamicArray colKeys = (ksDynamicArray)kompas.GetDynamicArray(23);
			if (attrPar  != null && usPar  != null && fVisibl  != null && colKeys  != null)
			{
				attrPar.Init();
				usPar.Init();
				attrPar.SetValues(usPar);
				attrPar.SetColumnKeys(colKeys);
				attrPar.SetFlagVisible(fVisibl);
				attrPar.key1 = 1;
				attrPar.key2 = 10;
				attrPar.key3 = 100;
				attrPar.password = "111";

				ksLtVariant item = (ksLtVariant)kompas.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
				ksDynamicArray arr = (ksDynamicArray)kompas.GetDynamicArray(23);
				if (item  != null && arr  != null)
				{
					usPar.SetUserArray(arr);
					item.Init();
					item.doubleVal = 987654321.0;
					arr.ksAddArrayItem(-1, item);
					item.Init();
					item.strVal = "qwerty";
					arr.ksAddArrayItem(-1, item);
					item.Init();
					item.longVal = 999991;
					arr.ksAddArrayItem(-1, item);

					item.Init();
					item.uCharVal = 1;
					fVisibl.ksAddArrayItem(-1, item);
					fVisibl.ksAddArrayItem(-1, item);
					fVisibl.ksAddArrayItem(-1, item);
				}

				ksRequestInfo info = (ksRequestInfo)kompas.GetParamStruct((short)StructType2DEnum.ko_RequestInfo);
				if (info != null)
				{
					info.Init();
					info.prompt = "Укажите объект";
					double x = 0, y = 0;
					reference pObj = 0;
					int j = doc.ksCursor(info, ref x, ref y, null);
					if (j != 0)
					{
						pObj = doc.ksFindObj(x, y, 1e6);
						if (doc.ksExistObj(pObj) != 0)
						{
							doc.ksLightObj(pObj, 1);
							//запросить имя библиотеки
							string nameFile = kompas.ksChoiceFile("*.lat", null, false);
							double numb = 0;
							j = kompas.ksReadDouble("Ввести номер типа атрибута", 1000, 0, 1e12, ref numb);
							if (j != 0)
							{
								reference pAttr = attr.ksCreateAttr(pObj, attrPar, numb, nameFile);
								if (pAttr == 0)
									kompas.ksMessageBoxResult();  // неудачное завершение - выдадим результат работы нашей функции
							}
							doc.ksLightObj(pObj, 0);
						}
					}
				}
			}
		}

		// Удалить первый атрибут  у данного объекта
		void DelObjAttr() 
		{
			ksRequestInfo info = (ksRequestInfo)kompas.GetParamStruct((short)StructType2DEnum.ko_RequestInfo);
			if (info != null) 
			{
				info.Init();
				info.prompt = "Укажите объект";
				double x = 0, y = 0;
				int j;
				do
				{
					j = doc.ksCursor(info, ref x, ref y, null);
					if (j != 0)
					{
						reference pObj = doc.ksFindObj(x, y, 1e6);
						if (doc.ksExistObj(pObj) == 1)
						{
							doc.ksLightObj(pObj, 1);
							//создадим итератор для хождения по атрибутам объекта
							ksIterator iter = (ksIterator)kompas.GetIterator();
							if (iter != null && iter.ksCreateAttrIterator(pObj, 0, 0, 0, 0, 0))
							{
								//встали на первый атрибут
								reference pAttr = iter.ksMoveAttrIterator("F", ref pObj);
								if (pAttr != 0)
								{
									string password = kompas.ksReadString("Ввести пароль типа атрибута", "");
									if (attr.ksDeleteAttr(pObj, pAttr, password) != 1) 
										kompas.ksMessageBoxResult();
								}
								else 
									kompas.ksMessage("атрибут не найден");
								doc.ksLightObj(pObj, 0);
							}
						}
					}
				} while (j != 0);
			}
		}

		// Считать  атрибут  типа double_str_long
		void ReadObjAttr() 
		{
			bool res = false;
			ksUserParam usPar = (ksUserParam)kompas.GetParamStruct((short)StructType2DEnum.ko_UserParam); 
			if (usPar != null)
			{
				usPar.Init();
				ksLtVariant item = (ksLtVariant)kompas.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
				ksDynamicArray arr = (ksDynamicArray)kompas.GetDynamicArray(23);
				if (item != null && arr != null)
				{
					usPar.SetUserArray(arr);
					item.Init();
					item.doubleVal = 987654321.0;
					arr.ksAddArrayItem(-1, item);
					item.Init();
					item.strVal = "qwerty";
					arr.ksAddArrayItem(-1, item);
					item.Init();
					item.longVal = 999991;
					arr.ksAddArrayItem(-1, item);
					res = true;
				}
			}
			if (res)
			{
				ksRequestInfo info = (ksRequestInfo)kompas.GetParamStruct((short)StructType2DEnum.ko_RequestInfo);
				if (info != null)
				{
					info.Init();
					info.prompt = "Укажите объект";
					double x = 0, y = 0;
					int j;
					do
					{
						j = doc.ksCursor(info, ref x, ref y, null);
						if (j != 0)
						{
							reference pObj = doc.ksFindObj(x, y, 1e6);
							if (doc.ksExistObj(pObj) != 0)
							{
								doc.ksLightObj(pObj, 1);
								//создадим итератор для хождения по атрибутам объекта
								ksIterator iter = (ksIterator)kompas.GetIterator();
								if (iter != null && iter.ksCreateAttrIterator(pObj, 0, 0, 0, 0, 0)) 
								{
									//встали на первый атрибут
									reference pAttr = iter.ksMoveAttrIterator("F", ref pObj);	
									if (pAttr != 0)
									{
										kompas.ksMessage("тип и ключи атрибута");
										int k1 = 0, k2 = 0, k3 = 0, k4 = 0;
										double numb = 0;
										attr.ksGetAttrKeysInfo(pAttr, out k1, out k2, out k3, out k4, out numb);
										string buf = string.Format("k1 = {0} k2 = {1} k3 = {2} k4 = {3} numb = {4}",
											k1, k2, k3, k4, numb);
										kompas.ksMessage(buf);

										kompas.ksMessage("строка атрибута");
										attr.ksGetAttrRow(pAttr, 0, 0, 0, usPar);

										kompas.ksMessage("заменим строку атрибута");
										ksLtVariant item = (ksLtVariant)kompas.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
										ksDynamicArray arr = (ksDynamicArray)usPar.GetUserArray();
										if (item != null && arr != null)
										{
											item.Init();
											item.doubleVal = numb;
											arr.ksSetArrayItem(0, item);
											item.Init();
											item.strVal = "1234567\nasdfgh\nzxcvb";
											arr.ksSetArrayItem(1, item);
											item.Init();
											item.longVal = 888881;
											arr.ksSetArrayItem(2, item);
											attr.ksSetAttrRow(pAttr, 0, 0, 0, usPar, "111");
										}
									}
									else 
										kompas.ksMessage("атрибут не найден");
								}
								doc.ksLightObj(pObj, 0);
							}
						}
					}
					while (j != 0);
				}
			}
		}

		// Просмотреть  атрибут
		void ShowObjAttr()
		{
			ksRequestInfo info = (ksRequestInfo)kompas.GetParamStruct((short)StructType2DEnum.ko_RequestInfo);
			if (info != null)
			{
				info.Init();
				info.prompt = "Укажите объект";
				double x = 0, y = 0;
				int j;
				do
				{
					j = doc.ksCursor(info, ref x, ref y, null);
					if (j != 0)
					{
						reference pObj = doc.ksFindObj(x, y, 1e6);
						if (doc.ksExistObj(pObj) != 0)
						{
							doc.ksLightObj(pObj, 1);
							attr.ksChoiceAttr(pObj);
							doc.ksLightObj(pObj, 0);
						}
					}
				}
				while (j != 0);
			}
		}

		void ShowLib()
		{
			//запросить имя библиотеки
			string nameFile = kompas.ksChoiceFile("*.lat", null, false);
			double numb = attr.ksChoiceAttrTypes(nameFile);
			if (numb > 1)  
			{
				string buf = string.Format("numbType = {0}", numb);
				kompas.ksMessage(buf);
			}
		}

		void ShowType()
		{
			//запросить имя библиотеки
			string nameFile = kompas.ksChoiceFile("*.lat", null, false);
			string password = string.Empty;
			double numb = 0;
			int j = kompas.ksReadDouble("Ввести номер типа атрибута", 1000, 0, 1e12, ref numb);
			if (j != 0)
			{
				password = kompas.ksReadString("Ввести пароль типа атрибута", "");
				attr.ksViewEditAttrType(nameFile, 2, numb, password);
			}
		}

		// Пройтись у объекта по атрибутам с ключом
		// key1 = 10 и выдать количество колонок и строк для данного атрибута
		void WalkFromObjWithAttr()
		{
			ksRequestInfo info = (ksRequestInfo)kompas.GetParamStruct((short)StructType2DEnum.ko_RequestInfo);
			if (info != null)
			{
				info.Init();
				info.prompt = "Укажите объект";
				double x = 0, y = 0;
				int j;
				do
				{
					j = doc.ksCursor(info, ref x, ref y, null);
					if (j != 0)
					{
						reference pObj = doc.ksFindObj(x, y, 1e6);
						if (doc.ksExistObj(pObj) == 1)
						{
							//создадим итератор для движения по атрибутам с ключом 10
							ksIterator iter = (ksIterator)kompas.GetIterator();
							if (iter != null && iter.ksCreateAttrIterator(pObj, 0, 0, 0, 0, 0)) 
							{
								doc.ksLightObj(pObj, 1);
								//встали на первый атрибут
								reference pAttr = iter.ksMoveAttrIterator("F", ref pObj);	
								if (pAttr != 0)
								{
									do
									{
										attr.ksViewEditAttr(pAttr, 1, string.Empty);
										pAttr = iter.ksMoveAttrIterator("N", ref pObj);
									}
									while(pAttr != 0);
								}
								doc.ksLightObj(pObj, 0);
							}
						}
					}
				}
				while (j != 0);
			}
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
