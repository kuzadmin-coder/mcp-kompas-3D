using Kompas6API5;

using System;
using Microsoft.Win32;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Drawing;
using Kompas6Constants;
using Kompas6Constants3D;
using KAPITypes;

using reference = System.Int32;

namespace Steps.NET
{
	// Класс Step3d2 - Работа с компонентой (деталь или сборка)
	public class Step3d2
	{
		private KompasObject kompas;
		private ksDocument3D doc;
		private string buf;

		// Имя библиотеки
		[return: MarshalAs(UnmanagedType.BStr)] public string GetLibraryName()
		{
			return "Step3d2 - Работа с компонентой (деталь или сборка)";
		}
		

		// Головная функция библиотеки
		public void ExternalRunCommand([In] short command, [In] short mode, [In, MarshalAs(UnmanagedType.IDispatch)] object kompas_)
		{
			kompas = (KompasObject) kompas_;
			if (kompas != null)
			{
				switch (command)
				{
					case 1 : CreateDocument3D();     break;  // Создание документа
					case 2 : DocIterator();          break;  // итератор по документам
					case 3 : UseEntityCollection();  break;  // использование массива элементов
					default:
						doc = (ksDocument3D)kompas.ActiveDocument3D();
						if (doc != null && doc.reference != 0)
						{
							switch (command)
							{
								case 4  : GetSetPartName();				break; // взять/изменить имя компоненты
								case 5  : FixAndStandartComponent();	break; // фиксирование и установка стандартного объекта
								case 6  : GetSetColorProperty();		break; // получить и заменить параметры цвета компоненты
								case 7  : GetSetArrayVariable();		break; // взять и поменять внешние переменные компоненты
								case 8  : GetSetPlacmentComponent();	break; // получить и изменить место расположения детали в сборке
								case 9  : GetSetEntity();				break; // Получить интерфейс ksEntity объекта создаваемого системой по умолчанию и поменять параметры
								case 10 : CreateSketch();				break; // создать эскиз
								case 11 : GetArraySketch();				break; // Формирует массив объектов(здесь эскизов) и возвращает его интерфейс ksEntityCollection (IEntityCollection)
								case 12 : GetSetUserParamComponent();	break; // Установить и получить параметры пользователя в компоненте
							}
						}
						break;
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
					result = "Создание 3D документа";
					command = 1;
					break;
				case 2:
					result = "Итератор по документам";
					command = 2;
					break;
				case 3:
					result = "Массивы элементов";
					command = 3;
					break;
				case 4:
					result = "Взять и поменять имя компоненты";
					command = 4;
					break;
				case 5:
					result = "Фиксирование и установка стандартного объекта";
					command = 5;
					break;
				case 6:
					result = "Получить и заменить параметры цвета компоненты";
					command = 6;
					break;
				case 7:
					result = "Взять и поменять внешние переменные компоненты";
					command = 7;
					break;
				case 8:
					result = "Получить и изменить расположение детали в сборке";
					command = 8;
					break;
				case 9:
					result = "Получить и изменить парметры базовой плоскости";
					command = 9;
					break;
				case 10:
					result = "Создать эскиз";
					command = 10;
					break;
				case 11:
					result = "Вернуть массив эскизов";
					command = 11;
					break;
				case 12:
					result = "Установить и получить параметры пользователя в компоненте";
					command = 12;
					break;
				case 13:
					itemType = 3;			//ENDMENU
					command = -1;
					break;
			}

			return result;
		}


		// Взять/изменить имя компоненты
		void GetSetPartName()
		{
			ksPart part = (ksPart)doc.GetPart((short)Part_Type.pTop_Part);	// верхний компонент
			if (part != null)
			{
				buf = string.Format("Имя компоненты {0}", part.name);
				kompas.ksMessage(buf);
				part.name = "Втулка";
				part.Update();
			}	 
		}


		// Фиксирование и установка стандартного объекта
		void FixAndStandartComponent()
		{
			if (doc.IsDetail())
			{
				kompas.ksError("Текущий документ должен быть сборкой");
				return;
			}
			ksPart part = (ksPart)doc.GetPart(0);	// первая деталь в сборке
			if (part != null)
			{
				// Получить состояние фиксации компонента - в системе поддерживающей работу со свойствами - fixedComponent
				bool fix = part.fixedComponent;
				// Получить состояние стандартного компонента - в системе поддерживающей работу со свойствами - standardComponent
				bool stand = part.standardComponent;
				kompas.ksMessage(fix ? "Компонент зафиксирован" : "Компонент не зафиксирован");
				// Изменить состояние фиксации компонента - в системе поддерживающей работу со свойствами - fixedComponent
				part.fixedComponent = !fix;
				kompas.ksMessage(stand ? "Компонент стандартный" : "Компонент нестандартный");
				// Изменить состояние стандартного компонента - в системе поддерживающей работу со свойствами - standardComponent
				part.standardComponent = !stand;
			}
		}


		// Получить и заменить параметры цвета компоненты
		void GetSetColorProperty()
		{
			ksPart part = (ksPart)doc.GetPart((short)Part_Type.pTop_Part); // верхний компонент
			if (part != null)
			{
				ksColorParam colorPr = (ksColorParam)part.ColorParam();
				if (colorPr != null)
				{
					buf = string.Format("Номер цвета = {0}\nОбщий цвет = {1}\nДиффузия = {2}\nЗеркальность = {3}\nБлеск = {4}\nПрозрачность = {5}\nИзлучение = {6}", 
						colorPr.color, colorPr.ambient, colorPr.diffuse,
						colorPr.specularity, colorPr.shininess,
						colorPr.transparency, colorPr.emission);
					kompas.ksMessage(buf);
					colorPr.color = 5421504;
					colorPr.transparency = 0.5;
					colorPr.ambient = 0.1;
					colorPr.diffuse = 0.1;
					part.Update();
					buf = string.Format("Номер цвета = {0}\nОбщий цвет = {1}\nДиффузия = {2}\nЗеркальность = {3}\nБлеск = {4}\nПрозрачность = {5}\nИзлучение = {6}", 
						colorPr.color, colorPr.ambient, colorPr.diffuse, 
						colorPr.specularity, colorPr.shininess, 
						colorPr.transparency, colorPr.emission);
					kompas.ksMessage(buf);
				}
			}
		}


		// Взять и поменять внешние переменные компоненты
		void GetSetArrayVariable()
		{
			if (doc.IsDetail())
			{
				kompas.ksError("Текущий документ должен быть сборкой");
				return;
			}
			ksPart part = (ksPart)doc.GetPart(0);	// первая деталь в сборке
			if (part != null)
			{
				// работа с массивом внешних переменных
				ksVariableCollection varCol = (ksVariableCollection)part.VariableCollection();
				if (varCol != null)
				{
					ksVariable var = (ksVariable)kompas.GetParamStruct((short)StructType2DEnum.ko_VariableParam);
					if (var == null)
						return;
					for (int i = 0; i < varCol.GetCount(); i ++)
					{
						var = (ksVariable)varCol.GetByIndex(i);
						buf = string.Format("Номер переменной {0}\nИмя переменной {1}\nЗначение переменной {2}\nКомментарий {3}", i, var.name, var.value, var.note);
						kompas.ksMessage(buf);
						if (i == 0)
						{
							var.note = "qwerty";
							double d = 0;
							kompas.ksReadDouble("Введи переменную", 10, 0, 100, ref d);
							var.value = d;
						}
					}
			
					for (int j = 0; j < varCol.GetCount(); j ++)
					{
						// просмотр изменненных переменных
						var = (ksVariable)varCol.GetByIndex(j);
						buf = string.Format("Номер переменной {0}\nИмя переменной {1}\nЗначение переменной {2}\nКомментарий {3}", j, var.name, var.value, var.note);
						kompas.ksMessage(buf);
					}
					part.RebuildModel();	// перестроение модели
				}
			}
		}


		// Получить и изменить место расположения детали в сборке
		void GetSetPlacmentComponent()
		{
			if (doc.IsDetail())
			{
				kompas.ksError("Текущий документ должен быть сборкой");
				return;
			}
			ksPart part = (ksPart)doc.GetPart(0);	// первая деталь в сборке
			if (part != null)
			{
				ksPlacement plac = (ksPlacement)part.GetPlacement();
				if (plac != null)
				{
					double x = 0, y = 0, z = 0;
					plac.GetOrigin(out x, out y, out z);
					buf = string.Format("x = {0}\ny = {1}\nz = {2}", x, y, z);
					kompas.ksMessage(buf);
					plac.SetOrigin(20, 20, 20);
					part.SetPlacement(plac);
					part.UpdatePlacement();
					part.Update();
				}
			}
		}


		// Получить интерфейс ksEntity объекта создаваемого системой по умолчанию и поменять параметры
		void GetSetEntity()
		{
			ksPart part = (ksPart)doc.GetPart((short)Part_Type.pTop_Part);	// верхний компонент
			if (part != null)
			{
				ksEntity planeXOY = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);	// 1-интерфейс на плоскость XOY
				if (planeXOY != null)
				{
					kompas.ksMessage(planeXOY.name);
					planeXOY.name = "Plane";
					planeXOY.Update();
				}
			}
		}


		// Создание эскиза
		void CreateSketch()
		{
			ksPart part = (ksPart)doc.GetPart((short)Part_Type.pTop_Part);	// верхний компонент
			if (part != null) 
			{
				ksEntity planeXOY = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);	// 1-интерфейс на плоскость XOY
				ksEntity entity = (ksEntity)part.NewEntity((short)Obj3dType.o3d_sketch);
				if (planeXOY != null && entity != null)
				{
					ksSketchDefinition sketch = (ksSketchDefinition)entity.GetDefinition();
					if (sketch != null)
					{
						sketch.SetPlane(planeXOY);
						entity.Create();
						ksDocument2D sketchDoc = (ksDocument2D)sketch.BeginEdit(); //GetSketchEditor());
						sketchDoc.ksLineSeg(0, 0, 100, 100, 1);
						sketch.EndEdit();
					}
				}
			}
		}


		// Формирует массив объектов(здесь эскизов) и возвращает его интерфейс ksEntityCollection (IEntityCollection)
		void GetArraySketch()
		{
			ksPart part = (ksPart)doc.GetPart((short)Part_Type.pTop_Part);	// верхний компонент
			if (part != null)
			{
				ksEntityCollection entityCollection = (ksEntityCollection)part.EntityCollection((short)Obj3dType.o3d_sketch);
				ksEntity currentEntity;
				if (entityCollection != null)
				{
					for (int k = 0; k < entityCollection.GetCount(); k++)
					{
						currentEntity = (ksEntity)entityCollection.GetByIndex(k);
						kompas.ksMessage(currentEntity.name);
					}
				}
			}
		}


		private struct dstruct
		{
			public double a, b;
			public int c, d;
		}


		// Установить и получить параметры пользователя в компоненте
		void GetSetUserParamComponent() 
		{
			if (doc.IsDetail()) 
			{
				kompas.ksError("Текущий документ должен быть сборкой");
				return;
			}

			ksPart part = (ksPart)doc.GetPart(0);	// первая деталь в сборке
			ksUserParam par = (ksUserParam)kompas.GetParamStruct((short)StructType2DEnum.ko_UserParam);
			ksLtVariant item = (ksLtVariant)kompas.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
			ksDynamicArray arr = (ksDynamicArray)kompas.GetDynamicArray(ldefin2d.LTVARIANT_ARR);
			if (par == null || item == null || arr == null || part == null)
				return;

			par.Init();
			par.SetUserArray(arr);
			item.Init();
			item.doubleVal = 12.12;
			arr.ksAddArrayItem(-1, item);
			item.Init();
			item.doubleVal = 21.21;
			arr.ksAddArrayItem(-1, item);
			item.Init();
			item.intVal = 666;
			arr.ksAddArrayItem(-1, item);
			item.Init();
			item.intVal = 999;
			arr.ksAddArrayItem(-1, item);

			part.SetUserParam(par);	// установка пользовательской структуры
			part.Update();

			buf = string.Format("Размер пользовательской структуры {0}", part.GetUserParamSize()); // размер пользовательской структуры
			kompas.ksMessage(buf);

			ksUserParam par2 = (ksUserParam)kompas.GetParamStruct((short)StructType2DEnum.ko_UserParam);
			ksLtVariant item2 = (ksLtVariant)kompas.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
			ksDynamicArray arr2 = (ksDynamicArray)kompas.GetDynamicArray(ldefin2d.LTVARIANT_ARR);
			if (par2 == null || item2 == null || arr2 == null)
				return;

			par2.Init();
			par2.SetUserArray(arr2);
			item2.Init();
			item2.doubleVal = 0;
			arr2.ksAddArrayItem(-1, item2);
			item2.Init();
			item2.doubleVal = 0;
			arr2.ksAddArrayItem(-1, item2);
			item2.Init();
			item2.intVal = 0;
			arr2.ksAddArrayItem(-1, item2);
			item2.Init();
			item2.intVal = 0;
			arr2.ksAddArrayItem(-1, item2);

			part.GetUserParam(par2);	// берем пользовательскeую структуру
  
			dstruct d;

			arr2.ksGetArrayItem(0, item2);
			d.a = item2.doubleVal;
			arr2.ksGetArrayItem(1, item2);
			d.b = item2.doubleVal;
			arr2.ksGetArrayItem(2, item2);
			d.c = item2.intVal;
			arr2.ksGetArrayItem(3, item2);
			d.d = item2.intVal;
			buf = string.Format("Переменные пользовательского масства\na = {0}\nb = {1}\nc = {2}\nd = {3}", d.a, d.b, d.c, d.d);
			kompas.ksMessage(buf);	// просмотрим переменные из пользовательского массива
		}


		// Создание документа 
		void CreateDocument3D()
		{
			ksDocument3D doc = (ksDocument3D)kompas.Document3D();
			if (doc.Create(false /*видимый*/, true /*деталь*/))
			{
				doc.author = "Ethereal";				// Автор документа
				doc.comment = "Пример документа 3D";	// Комментарии к документу
				doc.fileName = @"c:\example.m3d";		// Имя файла Документа
				doc.UpdateDocumentParam();				// Обновить параметры Документа
				doc.Save();								// Сохранить документ
				kompas.ksMessage("Сохраним документ под другим именем");
				doc.SaveAs(@"c:\example_1.m3d");		// сохранить документ под другим именем

				// Автор документа
				buf = string.Format("Автор документа: {0}", doc.author);
				kompas.ksMessage(buf);
				// Комментарий к документу
				buf = string.Format("Комментарий к документу: {0}", doc.comment);
				kompas.ksMessage(buf);
				// Имя файла
				buf = string.Format("Имя файла: {0}", doc.fileName);
				kompas.ksMessage(buf);

				doc.close();	// закроем документ
			}
		}


		// Открытие группы документов и создание итератора по ним
		void DocIterator()
		{
			ksDocument3D doc = null;
			ksDynamicArray arrDoc = (ksDynamicArray)kompas.GetDynamicArray(ldefin2d.CHAR_STR_ARR);	// динамический массив указателей на строки символов
			// если массив создан запускаем диалог выбора файлов 
			if (arrDoc != null && kompas.ksChoiceFiles("*.m3d","Документы (*.m3d)|*.m3d|Все файлы (*.*)|*.*", arrDoc, false) == 1)
			{
				ksChar255 item = (ksChar255)kompas.GetParamStruct((short)StructType2DEnum.ko_Char255); 
				if (item != null)
				{
					// откроем все файлы указанные пользователем
					for (int i = 0, count = arrDoc.ksGetArrayCount(); i < count; i++)
					{
						if (arrDoc.ksGetArrayItem(i, item) == 1)
						{ 
							// захватываем интерфейс документа
							doc = (ksDocument3D)kompas.Document3D();
							doc.Open(item.str, false);	// открываем файл с заданным именем 
						}
					}
				}

				// сообщаем количество открытых файлов
				buf = string.Format("Открыто {0} файлов", arrDoc.ksGetArrayCount());
				kompas.ksMessage(buf); 

				// создаем итератор по документам
				ksIterator iter = (ksIterator)kompas.GetIterator();
				if (iter != null)
				{
					//тип объекта             
					if (iter.ksCreateIterator(ldefin2d.D3_DOCUMENT_OBJ, 0))
					{
						reference rf; // референс документа
						if ((rf = iter.ksMoveIterator("F")) != 0)
						{
							// смещаем итератор на первый элемент
							do
							{
								// смещаем итератор на следующую позицию
								// захватываем интерфейс документа
								doc = (ksDocument3D)kompas.ksGet3dDocumentFromRef(rf);
								if (doc != null)
								{
									// автор документа
									buf = string.Format("Автор документа: {0}", doc.author);
									kompas.ksMessage(buf);
									// комментарии к документу
									buf = string.Format("Комментарий к документу: {0}", doc.comment);
									kompas.ksMessage(buf);
									// имя файла документа 
									buf = string.Format("Имя файла: {0}", doc.fileName);
									kompas.ksMessage(buf);
									// тип документа
									buf = string.Format("Тип документа: {0}", doc.IsDetail() ? "Деталь" : "Сборка");
									kompas.ksMessage(buf);
								}
							}
							while ((rf = iter.ksMoveIterator("N")) != 0);

							iter.ksDeleteIterator(); // удалить итератор
						}
					}
				}
			}
		}


		// Использование массива элементов
		void UseEntityCollection()
		{
			ksDocument3D doc = (ksDocument3D)kompas.ActiveDocument3D();	// привязываемся к активному документу
			if (doc != null)
			{
        ksPart part = (ksPart)doc.GetPart((short)Part_Type.pTop_Part);	// новый компонент
				if (part != null) 
				{
					int count1 = 0;
					int count2 = 0;	// количество плоских поверхностей
					int count = 0;	// количество конических поверхностей

					// массив поверхностей
					ksEntityCollection collect = (ksEntityCollection)part.EntityCollection((short)Obj3dType.o3d_face);      
					if (collect != null)
					{
						count = collect.GetCount();
						count1 = 0; 
						count2 = 0; 
						ksColorParam colorPr = null;	// интерфейс свойств цвета
						if (collect != null && count != 0)
						{
							for (int i = 0; i < count; i ++)
							{
								ksEntity ent = (ksEntity)collect.GetByIndex(i);
								colorPr = (ksColorParam)ent.ColorParam();
								// интерфейс свойств поверхности
								ksFaceDefinition faceDef = (ksFaceDefinition)ent.GetDefinition();
								if (faceDef != null)
								{
									// коническая по-ть		//цилиндрическая по-ть
									if (faceDef.IsCone() || faceDef.IsCylinder()) 
									{      
										colorPr.color = Color.FromArgb(0, 255, 255, 0).ToArgb();  
										count2 ++;	// считаем количество объектов
									}
									// плоская по-ть  
									if (faceDef.IsPlanar()) 
									{
										colorPr.color = Color.FromArgb(0, 0, 255, 255).ToArgb(); 
										count1 ++;	// считаем количество объектов
									}

									ent.Update();	// обновить параметры
								}
							}
						}
					}

					// сообщяем о результатах работы
					if (count == 0) 
						kompas.ksMessage("Не найдено ни одной поверхности");
					else
					{
						buf = string.Format("Найдено {0} конических и {1} плоских объектов", count2, count1);
						kompas.ksMessage(buf);
					}

					count1 = 0;
					count2 = 0;
					// массив ребер
					ksEntityCollection collect2 = (ksEntityCollection)part.EntityCollection((short)Obj3dType.o3d_edge);
					count = collect2.GetCount();
					if (collect2 != null && count != 0)
					{
						for (int i = 0; i < count; i++)
						{
							ksEntity ent = (ksEntity)collect2.GetByIndex(i);
							ksEdgeDefinition edgeDef = (ksEdgeDefinition)ent.GetDefinition();
							if (edgeDef != null)
							{
								if (edgeDef.IsStraight()) 
									count1++;	// количество прямых ребер
								else
									count2++;	// количество криволинейных ребер
							}
						}
					}

					// сообщяем о результатах работы
					if (count == 0) 
						kompas.ksMessage("Не найдено ни одного ребра");
					else
					{
						buf = string.Format("Найдено {0} прямых и {1} криволинейных ребер", count1, count2);
						kompas.ksMessage(buf);
					}
				}
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
