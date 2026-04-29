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
	// Класс Step10 - Hавигация по модели
  [ClassInterface(ClassInterfaceType.AutoDual)]
	public class Step10
	{
		private KompasObject kompas;
		private ksDocument2D docActive;

		// Имя библиотеки
		[return: MarshalAs(UnmanagedType.BStr)] public string GetLibraryName()
		{
			return "Step10 - Hавигация по модели";
		}
		

		// Головная функция библиотеки
		public void ExternalRunCommand([In] short command, [In] short mode, [In, MarshalAs(UnmanagedType.IDispatch)] object kompas_)
		{
			kompas = (KompasObject) kompas_;
			if (kompas != null)
			{
				if (command == 2 || command == 3)
				{
					docActive = (ksDocument2D) kompas.ActiveDocument2D();
					if (docActive != null && docActive.reference != 0)
					{
						switch (command)
						{
							case 2:	CreateDet(docActive);		break;	// создать объект спецификации деталь
							case 3:	CreateStandart(docActive);	break;	// создать стандартный объект
						}
					}
					else
						kompas.ksError("Документ не активизирован или \nне является листом/фрагментом");
				}
				else
				{
					switch (command)
					{
						case 1:	TypeAttrBolt();	break; // создадим  тип атрибута болта
						case 4: DecomposeSpc();	break; // kонвертировать спецификацию во фрагмент
						case 5: ShowSpc();		break; // просмотреть спецификацию
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
					result = "Создать шаблон обозначения Болт";
					command = 1;
					break;
				case 2:
					result = "Создать объект для раздела Детали";
					command = 2;
					break;
				case 3:
					result = "Создать объект для раздела Стандартные изделия";
					command = 3;
					break;
				case 4:
					result = "Конвертировать спецификацию во фрагмент";
					command = 4;
					break;
				case 5:
					result = "Просмотреть текущую спецификацию";
					command = 5;
					break;
				case 6:
					itemType = 3;			//ENDMENU
					command = -1;
					break;
			}

			return result;
		}


		// Создадим  тип атрибута болта
		void TypeAttrBolt()
		{
			ksAttributeObject attr = (ksAttributeObject)kompas.GetAttributeObject();
			ksAttributeTypeParam type = (ksAttributeTypeParam)kompas.GetParamStruct((short)StructType2DEnum.ko_AttributeType);
			ksColumnInfoParam col = (ksColumnInfoParam)kompas.GetParamStruct((short)StructType2DEnum.ko_ColumnInfoParam);
			if (type != null && col != null && attr != null) 
			{
				type.Init();
				col.Init(); 
				type.header = "Болт111";		// заголовoк-комментарий типа
				type.rowsCount = 1;				// кол-во строк в таблице
				type.flagVisible = true;		// видимый, невидимый   в таблице
				type.password = string.Empty;	// пароль, если не пустая строка  - защищает от несанкционированного изменения типа
				type.key1 = 10;
				type.key2 = 20;
				type.key3 = 30;
				type.key4 = 0;
				ksDynamicArray arr = (ksDynamicArray)type.GetColumns();
				if (arr != null) 
				{
					// колонка 1 "Имя элем."
					col.header = "Имя элем.";				// заголовoк-комментарий столбца
					col.type = ldefin2d.STRING_ATTR_TYPE;	// тип данных в столбце - см.ниже
					col.key = 1;							// дополнительный признак, который позволит отличить две переменные с одинаковым типом
					col.def = "Болт";						// значение по умолчанию
					col.flagEnum = false;					// флаг включающий режим, когда значение поля атрибута
					arr.ksAddArrayItem(-1, col);

					// колонка 2 "Исполнение"
					col.header = "Исполнение";				// заголовoк-комментарий столбца
					col.type = ldefin2d.UINT_ATTR_TYPE;		// тип данных в столбце - см.ниже
					col.key = 3;							// дополнительный признак, который позволит отличить две переменные с одинаковым типом
					col.def = "1";							// значение по умолчанию
					col.flagEnum = false;					// флаг включающий режим, когда значение поля атрибута
					arr.ksAddArrayItem(-1, col);

					// колонка 3 "Резьба"
					col.header = "Резьба";					// заголовoк-комментарий столбца
					col.type = ldefin2d.STRING_ATTR_TYPE;	// тип данных в столбце - см.ниже
					col.key = 0;							// дополнительный признак, который позволит отличить две переменные с одинаковым типом
					col.def = "M";							// значение по умолчанию
					col.flagEnum = false;					// флаг включающий режим, когда значение поля атрибута
					arr.ksAddArrayItem(-1, col);

					// колонка 4 "Диаметр"
					col.header = "Диаметр";					// заголовoк-комментарий столбца
					col.type = ldefin2d.UINT_ATTR_TYPE;		// тип данных в столбце - см.ниже
					col.key = 3;							// дополнительный признак, который позволит отличить две переменные с одинаковым типом
					col.def = "12";							// значение по умолчанию
					col.flagEnum = false;					// флаг включающий режим, когда значение поля атрибута
					arr.ksAddArrayItem(-1, col);

					// колонка 5 "разделитель"
					col.header = string.Empty;				// заголовoк-комментарий столбца
					col.type = ldefin2d.STRING_ATTR_TYPE;	// тип данных в столбце - см.ниже
					col.key = 0;							// дополнительный признак, который позволит отличить две переменные с одинаковым типом
					col.def = "x";							// значение по умолчанию
					col.flagEnum = false;					// флаг включающий режим, когда значение поля атрибута
					arr.ksAddArrayItem(-1, col);

					// колонка 6 "Шаг"
					col.header = "Шаг";						// заголовoк-комментарий столбца
					col.type = ldefin2d.FLOAT_ATTR_TYPE;	// тип данных в столбце - см.ниже
					col.key = 3;							// дополнительный признак, который позволит отличить две переменные с одинаковым типом
					col.def = "1.25";						// значение по умолчанию
					col.flagEnum = false;					// флаг включающий режим, когда значение поля атрибута
					arr.ksAddArrayItem(-1, col);

					// колонка 7 "Поле допуска"
					col.header = "Поле допуска";			// заголовoк-комментарий столбца
					col.type = ldefin2d.STRING_ATTR_TYPE;	// тип данных в столбце - см.ниже
					col.key = 3;							// дополнительный признак, который позволит отличить две переменные с одинаковым типом
					col.def = "-6g";						// значение по умолчанию
					col.flagEnum = false;					// флаг включающий режим, когда значение поля атрибута
					arr.ksAddArrayItem(-1, col);

					// колонка 8 "разделитель"
					col.header = string.Empty;				// заголовoк-комментарий столбца
					col.type = ldefin2d.STRING_ATTR_TYPE;	// тип данных в столбце - см.ниже
					col.key = 0;							// дополнительный признак, который позволит отличить две переменные с одинаковым типом
					col.def = "x";							// значение по умолчанию
					col.flagEnum = false;					// флаг включающий режим, когда значение поля атрибута
					arr.ksAddArrayItem(-1, col);

					// колонка 9 "Длина"
					col.header = "Длина";					// заголовoк-комментарий столбца
					col.type = ldefin2d.UINT_ATTR_TYPE;		// тип данных в столбце - см.ниже
					col.key = 3;							// дополнительный признак, который позволит отличить две переменные с одинаковым типом
					col.def = "60";							// значение по умолчанию
					col.flagEnum = false;					// флаг включающий режим, когда значение поля атрибута
					arr.ksAddArrayItem(-1, col);

					// колонка 10 "Кл. прочности"
					col.header = "Кл. прочности";			// заголовoк-комментарий столбца
					col.type = ldefin2d.STRING_ATTR_TYPE;	// тип данных в столбце - см.ниже
					col.key = 0;							// дополнительный признак, который позволит отличить две переменные с одинаковым типом
					col.def = "58";							// значение по умолчанию
					col.flagEnum = false;					// флаг включающий режим, когда значение поля атрибута
					arr.ksAddArrayItem(-1, col);

					// колонка 11 "Материал"
					col.header = "Материал";				// заголовoк-комментарий столбца
					col.type = ldefin2d.STRING_ATTR_TYPE;	// тип данных в столбце - см.ниже
					col.key = 0;							// дополнительный признак, который позволит отличить две переменные с одинаковым типом
					col.def = ".35Х";						// значение по умолчанию
					col.flagEnum = false;					// флаг включающий режим, когда значение поля атрибута
					arr.ksAddArrayItem(-1, col);

					// колонка 12 "Покрытие"
					col.header = "Покрытие";				// заголовoк-комментарий столбца
					col.type = ldefin2d.STRING_ATTR_TYPE;	// тип данных в столбце - см.ниже
					col.key = 0;							// дополнительный признак, который позволит отличить две переменные с одинаковым типом
					col.def = ".16";						// значение по умолчанию
					col.flagEnum = false;					// флаг включающий режим, когда значение поля атрибута
					arr.ksAddArrayItem(-1, col);

					// колонка 13 "ГОСТ"
					col.header = "ГОСТ";					// заголовoк-комментарий столбца
					col.type = ldefin2d.STRING_ATTR_TYPE;	// тип данных в столбце - см.ниже
					col.key = 0;							// дополнительный признак, который позволит отличить две переменные с одинаковым типом
					col.def = "ГОСТ";						// значение по умолчанию
					col.flagEnum = false;					// флаг включающий режим, когда значение поля атрибута
					arr.ksAddArrayItem(-1, col);

					// колонка 14 "Номер"
					col.header = "Номер";					// заголовoк-комментарий столбца
					col.type = ldefin2d.UINT_ATTR_TYPE;		// тип данных в столбце - см.ниже
					col.key = 2;							// дополнительный признак, который позволит отличить две переменные с одинаковым типом
					col.def = "7808";						// значение по умолчанию
					col.flagEnum = false;					// флаг включающий режим, когда значение поля атрибута
					arr.ksAddArrayItem(-1, col);

					// колонка 15 "разделитель"
					col.header = string.Empty;				// заголовoк-комментарий столбца
					col.type = ldefin2d.STRING_ATTR_TYPE;	// тип данных в столбце - см.ниже
					col.key = 0;							// дополнительный признак, который позволит отличить две переменные с одинаковым типом
					col.def = "-";							// значение по умолчанию
					col.flagEnum = false;					// флаг включающий режим, когда значение поля атрибута
					arr.ksAddArrayItem(-1, col);

					// колонка 16 "Год"
					col.header = "Год";						// заголовoк-комментарий столбца
					col.type = ldefin2d.STRING_ATTR_TYPE;	// тип данных в столбце - см.ниже
					col.key = 0;							// дополнительный признак, который позволит отличить две переменные с одинаковым типом
					col.def = "70";							// значение по умолчанию
					col.flagEnum = false;					// флаг включающий режим, когда значение поля атрибута
					arr.ksAddArrayItem(-1, col);
				}

				//запросить имя библиотеки
				string nameFile = kompas.ksChoiceFile("*.lat", null, false);
				if (nameFile.Length > 0) 
				{
					//создать тип атрибута
					double  numbType = attr.ksCreateAttrType(type,	// информация о типе атрибута
						nameFile);									// имя библиотеки типов атрибутов
					if (numbType > 1)  
					{
						string buf = string.Format("numbType = {0}", numbType);
						kompas.ksMessage(buf);
					}
					else
						kompas.ksMessageBoxResult();	// проверяем результат работы нашей функции
				}
				arr.ksDeleteArray();	//удалим массив колонок
			}
		}

		// Создать объект спецификации деталь
		void CreateDet(ksDocument2D doc)
		{
			ksSpecification spc = (ksSpecification)doc.GetSpecification();
			if (spc != null) 
			{
				//создать модельную группу   1
				reference gr = doc.ksNewGroup(0);
				doc.ksLineSeg (20, 30, 70, 30, 2);
				doc.ksLineSeg (70, 30, 70, 80, 2);
				doc.ksLineSeg (70, 80, 20, 80, 2);
				doc.ksLineSeg (20, 80, 20, 30, 2);
				doc.ksEndGroup();

				//создать  объект спецификации
				reference spcObj = EditSpcObjDet(gr, doc, spc);
				if (spcObj != 0)
					DrawPosLeader(spcObj, doc, spc);
			}
		}

		// Создать стандартный объект
		void CreateStandart(ksDocument2D doc)
		{
			ksSpecification spc = (ksSpecification)doc.GetSpecification();
			if (spc != null) 
			{
				BOLT tmp = new BOLT();
				tmp.gost = 7808;
				tmp.f = 0;
				tmp.p = 1; 
				tmp.L = 55;
				tmp.l1 = 49;
				tmp.b = 46;
				tmp.h2 = 10;
				tmp.klass = 1; /*klass = B*/
				tmp.d2 = 22.5F;
				tmp.k = 2;
				tmp.dr = 20;
				//создать  объект спецификации
				reference spcObj = EditStandartSpcObj(tmp, 0, doc, spc, 0);
				if (spcObj != 0)
					DrawPosLeader(spcObj, doc, spc);
			}
		}

		// Просмотреть спецификацию
		void ShowSpc() 
		{
			ksDocument2D doc = (ksDocument2D)kompas.Document2D();
			ksSpcDocument spc = (ksSpcDocument)kompas.SpcActiveDocument();
			if (doc != null && spc != null && spc.reference != 0) 
			{
				ksSpecification specification = (ksSpecification)spc.GetSpecification();
				ksIterator iter = (ksIterator)kompas.GetIterator();
				iter.ksCreateSpcIterator(null, 0, 0);
				if (iter.reference != 0 && specification != null) 
				{
					int obj = iter.ksMoveIterator("F");
					if (obj != 0) 
					{
						do 
						{
							//узнаем количество колонок у базового объекта спецификации
							int count = specification.ksGetSpcTableColumn(null, 0, 0);
  			  
							string buf = string.Format("Кол-во колонок = {0}", count);
							kompas.ksMessage(buf);
      
							// пройдем по всем колонкам
							for (int i = 1; i <= count; i++) 
							{
								// для текущего номера определим тип колонки, номер исполнения и блок
								ksSpcColumnParam spcColPar = (ksSpcColumnParam)kompas.GetParamStruct((short)StructType2DEnum.ko_SpcColumnParam);
								if (specification.ksGetSpcColumnType(obj,	//объект спецификации
									i,										// номер колонки, начиная с 1
									spcColPar) == 1)
								{
									// возьмем текст
									int columnType = spcColPar.columnType;
									int ispoln = spcColPar.ispoln;
									int blok = spcColPar.block;
									buf = specification.ksGetSpcObjectColumnText(obj, columnType, ispoln, blok);
									kompas.ksMessage(buf);
									// по типу колонки, номеру исполнения и блоку определим номер колонки
									int colNumb = specification.ksGetSpcColumnNumb(obj,	//объект спецификации
										spcColPar.columnType, spcColPar.ispoln, spcColPar.block);
									buf = string.Format("i = {0} colNumb = {1}", i, colNumb);
									kompas.ksMessage(buf);
								}
							}
						}
						while ((obj = iter.ksMoveIterator("N")) != 0);
					}
				}	
			}
			else 
				kompas.ksError("Спецификация должна быть текущей");
		}

		// Просмотреть спецификацию
		void DecomposeSpc() 
		{
			ksSpcDocument spc = (ksSpcDocument)kompas.SpcActiveDocument();
			ksDocument2D doc = (ksDocument2D)kompas.Document2D();
			if (spc != null && doc != null) 
			{
				// берем текущий документ спецификацию
				reference pDoc = spc.reference;
				if (pDoc == 0)
					return;
				// найдем количество листов спецификации
				int pageCount = spc.ksGetSpcDocumentPagesCount();
				// найдем габариры одного листа спецификации
				ksRectParam spcGabarit = (ksRectParam)kompas.GetParamStruct((short)StructType2DEnum.ko_RectParam);
				if (spcGabarit == null)
					return;
				doc.ksGetObjGabaritRect(pDoc, spcGabarit);

				// создадим фрагмент  
				ksDocumentParam docPar = (ksDocumentParam)kompas.GetParamStruct((short)StructType2DEnum.ko_DocumentParam);
				if (docPar == null)
					return;
				docPar.Init();
				docPar.regime = 0;
				docPar.type = 3;
				doc.ksCreateDocument (docPar);
				for (int i = 0; i < pageCount; i ++) 
				{
					// получили временную группу i-го листа спецификации
					reference group = doc.ksDecomposeObj(pDoc,	//указатель на объект
						0,										//уровень разбиения 0- отрезки,дуги,тексты,точки:1-отрезки,тексты,точки; 2-отрезки, дуги, тексты
						0.4,									//стрелка прогиба
						Convert.ToInt16(i + 1));				// 0 - разбиение объекта в СК вида 1- в СК листа
					if (group != 0) 
					{
						int column = i % 3;
						ksMathPointParam mathBop = (ksMathPointParam)spcGabarit.GetpBot();
						ksMathPointParam mathTop = (ksMathPointParam)spcGabarit.GetpTop();
						if (mathBop == null && mathTop == null)
							return;
						double x = (mathTop.x - mathBop.x + 5) * column;
						int row = i / 3;
						double y = (mathTop.y - mathBop.y + 5) * row;
						// сдвинули группу
						doc.ksMoveObj(group, x, -y);
						// поставили в модель
						doc.ksStoreTmpGroup(group);	// указатель группы
						doc.ksClearGroup(group, true);
						doc.ksDeleteObj(group);
					}
				}
				string buf = string.Format("Декомпозировано {0} листов СП", pageCount);
				kompas.ksMessage(buf);
			}
			else
				kompas.ksError("Спецификация должна быть текущей");
		}

		// Создадим объект спецификации для раздела "Детали"
		reference  EditSpcObjDet(reference geom, ksDocument2D doc, ksSpecification spc)
		{
			reference spcObj = 0;
			// если редактируем макро объект, то ввойдем в режим редактирования объекта
			if (doc.ksEditMacroMode() == 1) 
			{
				// найдем объкт спецификации по геометрии
				spcObj = spc.ksGetSpcObjForGeom("graphic.lyt",	// имя библиотеки типов 
					1,											// номер типа спецификации
					0,											// для макро редактирования
					1, 1);
				// ввойдем в режим редактирования                               
				if (spc.ksSpcObjectEdit(spcObj) != 1)
					spcObj  = 0;
			}

			// если объекта нет, ввойдем в режим создания объекта спецификации
			if (spcObj != 0 || spc.ksSpcObjectCreate("graphic.lyt",	// имя библиотеки типов 
				1,													// номер типа спецификации
				20, 0,												// номер раздела и подраздела
				0, 0) == 1) 
			{
				// тип атрибута
				ksUserParam par = (ksUserParam)kompas.GetParamStruct((short)StructType2DEnum.ko_UserParam);
				ksLtVariant item = (ksLtVariant)kompas.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
				ksDynamicArray arr = (ksDynamicArray)kompas.GetDynamicArray(ldefin2d.LTVARIANT_ARR);
				if (par == null || item == null || arr == null)
					return 0;
				par.Init();
				par.SetUserArray(arr);
				item.Init();
				item.strVal = "Втулка";
				arr.ksAddArrayItem(-1, item);
				// наименование
				spc.ksSpcChangeValue(5 /*номер колонки*/, 1, par, ldefin2d.STRING_ATTR_TYPE );

				// подключим геометрию
				if (geom != 0)
					spc.ksSpcIncludeReference(geom, 1);

				spcObj = spc.ksSpcObjectEnd();
				// если объект спецификации создан дадим возможность пользователю на него посмотреть и 
				// отредактировать  функцию нужно запускать вне Cursor и Placement
				if (spcObj != 0 && spc.ksEditWindowSpcObject(spcObj) == 1)
					return spcObj;
			}
			return 0;
		}

		// Отрисовать позиционную линию выноски
		// уже проверено, что объект спецификации есть
		// Функцию нужно запускать вне Cursor и Placement
		void DrawPosLeader(reference spcObj, ksDocument2D doc, ksSpecification spc)
		{
			ksRequestInfo info = (ksRequestInfo)kompas.GetParamStruct((short)StructType2DEnum.ko_RequestInfo);
			info.Init();
			bool flag = false;
			reference posLeater = 0;
			double x1 = 0, y1 = 0;
			do 
			{
				info.commandsString = "!Создать !Подключить ";
				info.prompt = "Укажите позиционную линию выноски";
				int j1 = doc.ksCursor(info, ref x1, ref y1, null);
				switch (j1) 
				{
					case 1: // Создать новую линию выноски
						posLeater = doc.ksCreateViewObject(ldefin2d.POSLEADER_OBJ);
						flag = false;
						break;
					case 2: // Подключить существующую
						info.prompt = "Укажите линию выноски";
						if (doc.ksCursor(info, ref x1, ref y1, 0) == -1) 
						{
							posLeater = doc.ksFindObj(x1, y1, 100);      // величина стороны окошка-ловушки с центром x,y
							if (!(posLeater != 0 && doc.ksGetObjParam(posLeater, 0, 0) == ldefin2d.POSLEADER_OBJ)) 
							{
								posLeater = 0;
								flag = true;
							}
							else
								flag = false;
							break;
						}
						else
							flag = false;
						break;
					case -1:
						posLeater = doc.ksFindObj(x1, y1, 100);      // величина стороны окошка-ловушки с центром x,y
						if (!(posLeater != 0 && doc.ksGetObjParam(posLeater, 0, 0) == ldefin2d.POSLEADER_OBJ)) 
						{
							kompas.ksError("Ошибка! Объект - не позиционная линия выноски!");
							posLeater = 0;
							flag = true;
						}
						else
							flag = false;
						break;
				}
			}
			while(flag);
  
			// линия выноски есть, подключим ее к объекту спецификации 
			if (posLeater != 0) 
			{
				// ввойдем в режим редактирования объекта спецификации
				if (spc.ksSpcObjectEdit(spcObj) == 1) 
				{
					// подключим линию выноски
					spc.ksSpcIncludeReference(posLeater, 1);
					// закроем объект спецификации
					spc.ksSpcObjectEnd();
				}
			}
		}

		// Создадим объект спецификации для раздела "Стандартные изделия"
		reference EditStandartSpcObj (BOLT tmp,  reference geom, ksDocument2D doc, ksSpecification spc, reference spcObj)
		{
			// reference spcObj = 0;
			// если редактируем макро объект , то ввойдем в режим редактирования объекта
			if (doc.ksEditMacroMode() == 1) 
			{
				// найдем объкт спецификации по геометрии
				spcObj = spc.ksGetSpcObjForGeom("graphic.lyt",	// имя библиотеки типов
					1,											// номер типа спецификации
					0,											// для макро редактирования
					1, 1);
				// ввойдем в режим редактирования
				if (spc.ksSpcObjectEdit(spcObj) != 1)
					spcObj  = 0;
			}

			if (spcObj!= 0 || spc.ksSpcObjectCreate("graphic.lyt",	// имя библиотеки типов
				1,													// номер типа спецификации
				25, 0,												// номер раздела и подраздела определены в стиле спецификации
				313277777065.0, 0) == 1) 
			{
				// тип атрибута определен в библиотеке типов атрибутов spc.lat
				int uBuf;
				ksUserParam par = (ksUserParam)kompas.GetParamStruct((short)StructType2DEnum.ko_UserParam);
				ksLtVariant item = (ksLtVariant)kompas.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
				ksDynamicArray arr = (ksDynamicArray)kompas.GetDynamicArray(ldefin2d.LTVARIANT_ARR);
				if (par == null || item == null || arr == null)
					return 0;
				par.Init();
				par.SetUserArray(arr);
				item.Init();
				item.strVal = "Болт111";
				arr.ksAddArrayItem(-1, item);

				spc.ksSpcChangeValue(5 /*номер колонки*/, 1, par, ldefin2d.STRING_ATTR_TYPE);
				// исполнение
				if ((tmp.f & 0x80) == 0) // если исполнения нет
					spc.ksSpcVisible(ldefin2d.SPC_CLM_NAME, 2, 0); // выключим исполнение
				else 
				{
					uBuf = 2;
					spc.ksSpcVisible(ldefin2d.SPC_CLM_NAME, 2, 1);
					arr.ksClearArray();
					item.Init();
					item.intVal = uBuf;
					arr.ksAddArrayItem(-1, item);
					spc.ksSpcChangeValue(ldefin2d.SPC_CLM_NAME, 2, par, ldefin2d.UINT_ATTR_TYPE );
				}

				// изменим диаметр
				arr.ksClearArray();
				item.Init();
				item.doubleVal = 40;
				arr.ksAddArrayItem(-1, item);
				spc.ksSpcChangeValue(ldefin2d.SPC_CLM_NAME, 4, par, ldefin2d.DOUBLE_ATTR_TYPE );

				// отследим мелкий шаг
				if ((tmp.f & 0x2) == 0)
				{
					// выключить шаг и его разделитель
					spc.ksSpcVisible(ldefin2d.SPC_CLM_NAME, 5, 0);
					spc.ksSpcVisible(ldefin2d.SPC_CLM_NAME, 6, 0);	// шаг
				}
				else 
				{
					spc.ksSpcVisible(ldefin2d.SPC_CLM_NAME, 5, 1);
					spc.ksSpcVisible(ldefin2d.SPC_CLM_NAME, 6, 1);   //шаг
					arr.ksClearArray();
					item.Init();
					item.floatVal = tmp.p;
					arr.ksAddArrayItem(-1, item);
					spc.ksSpcChangeValue(ldefin2d.SPC_CLM_NAME, 6, par, ldefin2d.FLOAT_ATTR_TYPE);
				}

				// выключим поле допуска
				spc.ksSpcVisible(ldefin2d.SPC_CLM_NAME, 7, 0);

				// изменим длину
				arr.ksClearArray();
				item.Init();
				item.intVal = (int)tmp.L;
				arr.ksAddArrayItem(-1, item);
				spc.ksSpcChangeValue(ldefin2d.SPC_CLM_NAME, 9, par, ldefin2d.UINT_ATTR_TYPE);

				// выключим класс прочности
				spc.ksSpcVisible(ldefin2d.SPC_CLM_NAME, 10, 0);

				// выключим материал
				spc.ksSpcVisible(ldefin2d.SPC_CLM_NAME, 11, 0);

				// выключим покрытие
				spc.ksSpcVisible(ldefin2d.SPC_CLM_NAME, 12, 0);

				// изменим ГОСТ
				arr.ksClearArray();
				item.Init();
				item.floatVal = tmp.gost;
				arr.ksAddArrayItem(-1, item);
				spc.ksSpcChangeValue(ldefin2d.SPC_CLM_NAME, 14, par, ldefin2d.UINT_ATTR_TYPE);

				// подключим геометрию
				if (geom != 0)
					spc.ksSpcIncludeReference(geom, 1);

				spcObj = spc.ksSpcObjectEnd();
				// если объект спецификации создан дадим возможность пользователю на него посмотреть и 
				// отредактировать. Функцию нужно запускать вне Cursor и Placement
				if (spcObj != 0)
					if (spc.ksEditWindowSpcObject(spcObj) != 0)
						return spcObj;
			}
			return 0;
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
