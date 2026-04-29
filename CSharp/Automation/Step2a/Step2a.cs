using Kompas6API5;

using System;
using Microsoft.Win32;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using KAPITypes;
using Kompas6Constants;

namespace Steps.NET
{
	// Класс Step2a - Массив неопределенной длины
	// 1. Массив строк                           - StrIndefiniteArray
	// 2. Массив математических точек            - PointIndefiniteArray
	// 3. Массив строк объекта "текст"           - TextIndefiniteArray
	// 4. Массив колонок типа атрибута           - AttrIndefiniteArray
	// 5. Массив полилиний                       - PolyLineArray
	// 6. Массив габаритных прямоугольников      - RectArray
	// 7. Массив структур пользователя           - UserDataArray
	// 8. Массив экземпляров класса пользователя - UserClassArray

	public class Step2a
	{
		private KompasObject kompas = null;

		[return: MarshalAs(UnmanagedType.BStr)]
		public string GetLibraryName()
		{
			return "Step2a - Массив неопределенной длины";
		}

		public void ExternalRunCommand([In] short command, [In] short mode, [In, MarshalAs(UnmanagedType.IDispatch)] object kompas_)
		{
			kompas = (KompasObject)kompas_;
			if (kompas != null)
			{
				switch (command)
				{
					case 1: StrIndefiniteArray(); break;	// массив строк
					case 2: PointIndefiniteArray(); break;	// массив математических точек
					case 3: TextIndefiniteArray(); break;	// массив строк объекта "текст"
					case 4: AttrIndefiniteArray(); break;	// массив колонок типа атрибута
					case 5: PolyLineArray(); break;			// массив полилиний
					case 6: RectArray(); break;				// массив габаритных прямоугольников
				}
			}
		}

		[return: MarshalAs(UnmanagedType.BStr)]
		public string ExternalMenuItem(short number, ref short itemType, ref short command)
		{
			string result = string.Empty;
			itemType = 1; // "MENUITEM"
			switch (number)
			{
				case 1:
					result = "Массив простых строк";
					command = 1;
					break;
				case 2:
					result = "Массив точек";
					command = 2;
					break;
				case 3:
					result = "Массив строк текста";
					command = 3;
					break;
				case 4:
					result = "Массив колонок типа атрибута";
					command = 4;
					break;
				case 5:
					result = "Массив полилиний";
					command = 5;
					break;
				case 6:
					result = "Массив габ. прямоугольников";
					command = 6;
					break;
				case 7:
					command = -1;
					itemType = 3; // "ENDMENU"
					break;
			}
			return result;
		}


		// Массив предназначен для создания и получения параетров параграфа текста,
		// состоящего из   нескольких строк, которые в свою очередь
		// состоят из компонент, с разными параметрами
		// (с наклоном, с утолщением, спецзнаки и т. д.)
		private void TextIndefiniteArray()
		{
			string buf = string.Empty;
			ksTextLineParam par = (ksTextLineParam)kompas.GetParamStruct((short)StructType2DEnum.ko_TextLineParam);
			ksTextItemParam par1 = (ksTextItemParam)kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemParam);
			// создали массив строк текста
			ksDynamicArray p = (ksDynamicArray)kompas.GetDynamicArray(ldefin2d.TEXT_LINE_ARR);
			if (par != null && par1 != null && p != null)
			{
				par.Init();
				par1.Init();
				// массив компонент строки текста
				ksDynamicArray p1 = (ksDynamicArray)par.GetTextItemArr();
				if (p1 != null)
				{
					ksTextItemFont font = (ksTextItemFont)par1.GetItemFont();
					if (font != null)
					{
						// создаем первую строку текста
						font.height = 10;   // высота текста
						font.ksu = 1;       // сужение текста
						font.color = 1000;  // цвет
						font.bitVector = 1; // битовый вектор (наклон, толщина, подчеркивание, тип составной части(дробь, отклонение, выражение типа суммы))
						par1.s = "1 компонента 1 строка";
						// добавили 1-ю компоненту  в массив компонент
						p1.ksAddArrayItem(-1, par1);

						font.height = 20;   // высота текста
						font.ksu = 2;       // сужение текста
						font.color = 2000;  // цвет
						font.bitVector = 2; // битовый вектор (наклон, толщина, подчеркивание, тип составной части(дробь, отклонение, выражение типа суммы))
						par1.s = "2 компонента 1 строка";
						// добавили 2-ю компоненту  в массив компонент
						p1.ksAddArrayItem(-1, par1);
						par.style = 1;

						// 1-я строка текста состоит из двух компонент добавим строку текста в
						// массив строк текста
						p.ksAddArrayItem(-1, par);

						// очистили массив компонент, чтобы использовать для создания второй
						// строки текста
						p1.ksClearArray();

						// создаем вторую строку текста
						font.height = 30;   // высота текста
						font.ksu = 3;       // сужение текста
						font.color = 3000;  // цвет
						font.bitVector = 3; // битовый вектор (наклон, толщина, подчеркивание, тип составной части(дробь, отклонение, выражение типа суммы))
						par1.s = "1 компонента 2 строка";
						// добавили 1-ю компоненту  в массив компонент
						p1.ksAddArrayItem(-1, par1);

						font.height = 40;   // высота текста
						font.ksu = 4;       // сужение текста
						font.color = 4000;  // цвет
						font.bitVector = 4; // битовый вектор (наклон, толщина, подчеркивание, тип составной части(дробь, отклонение, выражение типа суммы))
						par1.s = "2 компонента 2 строка";
						// добавили 2-ю компоненту  в массив компонент
						p1.ksAddArrayItem(-1, par1);

						par.style = 2;

						// 2-я строка текста состоит из двух компонент добавим строку текста в
						// массив строк текста		 }
						p.ksAddArrayItem(-1, par);

						kompas.ksMessageBoxResult();

						int n = p.ksGetArrayCount();
						buf = string.Format(" n = {0} ", n);
						kompas.ksMessage(buf);
						// просмотрим массив строк текста
						for (int i = 0; i < n; i++)
						{  // цикл по строкам текста
							p.ksGetArrayItem(i, par);
							buf = string.Format("i = {0}: style = {1},", i, par.style);
							kompas.ksMessage(buf);

							int n1 = p1.ksGetArrayCount();
							for (int j = 0; j < n1; j++)
							{  // цикл по компонентам строки текста
								p1.ksGetArrayItem(j, par1);
								buf = string.Format("j = {0}:  h = {1:0.#}, s = {2}", j, font.height, par1.s);
								kompas.ksMessage(buf);
							}
						}

						kompas.ksMessageBoxResult(); // проверяем результат работы нашей функции

						// заменим вторую компоненту у первой строки
						p.ksGetArrayItem(0, par);
						font.height = 50;   // высота текста
						font.ksu = 1;       // сужение текста
						font.color = 1000;  // цвет
						font.bitVector = 1; // битовый вектор (наклон, толщина, подчеркивание, тип составной части(дробь, отклонение, выражение типа суммы))
						par1.s = "2 комп. 1 стр.";
						p1.ksSetArrayItem(1, par1);
						par.style = 3;
						p.ksSetArrayItem(0, par);

						// заменим первую компоненту у второй строки
						p.ksGetArrayItem(1, par);
						font.height = 60;   // высота текста
						font.ksu = 1;       // сужение текста
						font.color = 1000;  // цвет
						font.bitVector = 1; // битовый вектор (наклон, толщина, подчеркивание, тип составной части(дробь, отклонение, выражение типа суммы))
						par1.s = "1 комп. 2 стр.";
						p1.ksSetArrayItem(0, par1);
						par.style = 4;
						p.ksSetArrayItem(1, par);

						n = p.ksGetArrayCount();
						// просмотрим массив строк текста
						for (int i = 0; i < n; i++)
						{  // цикл по строкам текста
							p.ksGetArrayItem(i, par);
							buf = string.Format("i = {0}: style = {1}, ", i, par.style);
							kompas.ksMessage(buf);

							int n1 = p1.ksGetArrayCount();
							for (int j = 0; j < n1; j++)
							{  // цикл по компонентам строки текста
								p1.ksGetArrayItem(j, par1);
								buf = string.Format("j = {0}:  h = {1:0.#}, s = {2}", j, font.height, par1.s);
								kompas.ksMessage(buf);
							}
						}

						kompas.ksMessageBoxResult(); // проверяем результат работы нашей функции

						p1.ksDeleteArray();
						p.ksDeleteArray();
					}
				}
			}
		}


		// Массив предназначен для хранения математических точек типа  MathPointParam
		void PointIndefiniteArray()
		{
			string buf = string.Empty;
			ksMathPointParam par = (ksMathPointParam)kompas.GetParamStruct((short)StructType2DEnum.ko_MathPointParam);
			// создать массив
			ksDynamicArray p = (ksDynamicArray)kompas.GetDynamicArray(ldefin2d.POINT_ARR);
			if (p != null && par != null)
			{
				// наполнить массив
				par.Init();
				par.x = 10;
				par.y = 10;
				p.ksAddArrayItem(-1, par);
				par.x = 20;
				par.y = 20;
				p.ksAddArrayItem(-1, par);
				par.x = 30;
				par.y = 30;
				p.ksAddArrayItem(-1, par);

				kompas.ksMessageBoxResult();

				// просмотрим массив
				int n = p.ksGetArrayCount();
				buf = string.Format("n = {0}", n);
				kompas.ksMessage(buf);

				for (int i = 0; i < n; i++)
				{
					p.ksGetArrayItem(i, par);
					buf = string.Format("i = {0},  x = {1:0.#}, y = {2:0.#}",
						i, par.x, par.y);
					kompas.ksMessage(buf);
				}

				// заменим параметры 1-го элемента
				par.x = 50;
				par.y = 50;
				p.ksSetArrayItem(1, par);

				// заменим параметры 0-го элемента
				par.x = 60;
				par.y = 60;
				p.ksSetArrayItem(0, par);

				// просмотрим массив
				n = p.ksGetArrayCount();
				for (int i = 0; i < n; i++)
				{
					p.ksGetArrayItem(i, par);
					buf = string.Format("i = {0}:  x = {1:0.#}, y = {2:0.#}",
						i, par.x, par.y);
					kompas.ksMessage(buf);
				}

				kompas.ksMessageBoxResult();

				p.ksDeleteArray();
			}
		}


		// Массив предназначен для хранения строк
		void StrIndefiniteArray()
		{
			string buf = string.Empty;
			// создать массив
			ksChar255 charS = (ksChar255)kompas.GetParamStruct((short)StructType2DEnum.ko_Char255);
			ksDynamicArray p = (ksDynamicArray)kompas.GetDynamicArray(ldefin2d.CHAR_STR_ARR);
			if (p != null && charS != null)
			{
				charS.Init();
				// наполним массив
				charS.str = "12345";
				p.ksAddArrayItem(-1, charS);
				charS.str = "67890";
				p.ksAddArrayItem(-1, charS);
				charS.str = "qwerty";
				p.ksAddArrayItem(-1, charS);

				kompas.ksMessageBoxResult();

				// просмотрим массив
				int n = p.ksGetArrayCount();
				buf = string.Format("n = {0}", n);
				kompas.ksMessage(buf);

				for (int i = 0; i < n; i++)
				{
					p.ksGetArrayItem(i, charS);
					kompas.ksMessage(charS.str);
				}

				// исключить из массива эл 1
				p.ksExcludeArrayItem(1);
				n = p.ksGetArrayCount();

				// просмотрим массив
				for (int i = 0; i < n; i++)
				{
					p.ksGetArrayItem(i, charS);
					kompas.ksMessage(charS.str);
				}

				kompas.ksMessageBoxResult();

				p.ksDeleteArray();
			}
		}


		private void ShowColumns(ksDynamicArray pCol, bool fl)
		{
			ksColumnInfoParam par = (ksColumnInfoParam)kompas.GetParamStruct((short)StructType2DEnum.ko_ColumnInfoParam);
			if (par != null)
			{
				par.Init();

				string buf = string.Empty;
				string s = string.Empty;
				if (fl)
					s = "структура";

				int n = pCol.ksGetArrayCount();

				for (int i = 0; i < n; i ++)
				{
					if (pCol.ksGetArrayItem(i, par) != 1)
						kompas.ksMessageBoxResult();  // проверяем результат работы нашей функции
					else
					{
						//выдадим поля колонки не указатели
						buf = string.Format("{0} i = {1} header = {2} type = {3} def = {4} flagEnum = {5}",
							s, i, par.header,
							par.type, par.def, par.flagEnum);
						kompas.ksMessage(buf);
						if (par.type == ldefin2d.RECORD_ATTR_TYPE)
						{
							// структура
							pCol = (ksDynamicArray)par.GetColumns();
							if (pCol != null)
							{
								ShowColumns(pCol, true);
								//pCol.ksDeleteArray();
							}
						}
						else
						{
							if (par.flagEnum)
							{
								// выдадим массив перечислений
								ksDynamicArray fEnum = (ksDynamicArray)par.GetFieldEnum();
								if (fEnum != null)
								{
									int n1 = fEnum.ksGetArrayCount();
									kompas.ksMessage("массив перечислений");
									ksChar255 charS = (ksChar255)kompas.GetParamStruct((short)StructType2DEnum.ko_Char255);
									for (int i1 = 0; i1 < n1; i1 ++)
										if (fEnum.ksGetArrayItem(i1, charS) != 1)
											kompas.ksMessageBoxResult();  // проверяем результат работы нашей функции
										else
											kompas.ksMessage(charS.str);
								}
							}
						}
					}
				}
			}
		}


		// Массив предназначен для создания и получения типа атрибута,
		// который может состоять из нескольких колонок
		void AttrIndefiniteArray()
		{
			// создадим массив из 3 колонок
			// первая колонка описывает  int с перечисленными значениями ( 100, 200, 300 )
			// вторая колонка - запись соответствует структуре
			// struct {
			//   double ;// умолчательное значение 123456789
			//   long   ;// умолчательное значение 1000000
			//   char   ;// умолчательное значение 10
			// }
			// третья колонка строка символов умолчательное значение "text"

			ksColumnInfoParam parCol1 = (ksColumnInfoParam)kompas.GetParamStruct((short)StructType2DEnum.ko_ColumnInfoParam);
			ksColumnInfoParam parCol2 = (ksColumnInfoParam)kompas.GetParamStruct((short)StructType2DEnum.ko_ColumnInfoParam);
			ksColumnInfoParam parCol3 = (ksColumnInfoParam)kompas.GetParamStruct((short)StructType2DEnum.ko_ColumnInfoParam);
			ksColumnInfoParam parStruct = (ksColumnInfoParam)kompas.GetParamStruct((short)StructType2DEnum.ko_ColumnInfoParam);
			ksDynamicArray pCol = (ksDynamicArray)kompas.GetDynamicArray(ldefin2d.ATTR_COLUMN_ARR);
			ksChar255 charS = (ksChar255)kompas.GetParamStruct((short)StructType2DEnum.ko_Char255);
			if (parCol1 != null && parCol2 != null && parCol3 != null
				&& parStruct != null && charS != null && pCol != null)
			{
				// первая колонка
				parCol1.Init();
				parCol1.header = "int";					// заголовoк-комментарий столбца
				parCol1.type = ldefin2d.INT_ATTR_TYPE;	// тип данных в столбце - см.ниже
				parCol1.key = 0;						// дополнительный признак, который позволит отличить две переменные с одинаковым типом
				parCol1.def = "100";					// значение по умолчанию
				parCol1.flagEnum = true;				// флаг включающий режим, когда значение поля атрибута будет заполнятся из массива перечисленных значений

				ksDynamicArray pEnum = (ksDynamicArray)kompas.GetDynamicArray(ldefin2d.CHAR_STR_ARR);
				if (pEnum != null)
				{
					//заполним массив перечисленных значений для первой колонки
					charS.Init();
					charS.str = "100";
					pEnum.ksAddArrayItem(-1, charS);
					charS.str = "200";
					pEnum.ksAddArrayItem(-1, charS);
					charS.str = "300";
					pEnum.ksAddArrayItem(-1, charS);

					parCol1.SetFieldEnum(pEnum);	// массив неопределенной длины перечислений (строки)
				}

				pCol.ksAddArrayItem(-1, parCol1);

				// вторая колонка
				parCol2.Init();
				parCol2.header = "struct";					// заголовoк-комментарий столбца
				parCol2.type = ldefin2d.RECORD_ATTR_TYPE;	// тип данных в столбце - см.ниже
				parCol2.key = 0;							// дополнительный признак, который позволит отличить две переменные с одинаковым типом
				parCol2.def = "\0";							// значение по умолчанию
				parCol2.flagEnum = false;					// флаг включающий режим, когда значение поля атрибута будет заполнятся из массива перечисленных значений

				ksDynamicArray pStruct = (ksDynamicArray)kompas.GetDynamicArray(ldefin2d.ATTR_COLUMN_ARR);
				if (pStruct != null)
				{
					// заполним массив солонок для структуры
					// первая колонка  структуры
					parStruct.Init();
					parStruct.header = "double";				// заголовoк-комментарий столбца
					parStruct.type = ldefin2d.DOUBLE_ATTR_TYPE;	// тип данных в столбце - см.ниже
					parStruct.def = "123456789";				// значение по умолчанию
					pStruct.ksAddArrayItem(-1, parStruct);

					// вторая колонка структуры
					parStruct.Init();
					parStruct.header = "long ";					// заголовoк-комментарий столбца
					parStruct.type = ldefin2d.LINT_ATTR_TYPE;	// тип данных в столбце - см.ниже
					parStruct.def = "1000000";					// значение по умолчанию
					pStruct.ksAddArrayItem(-1, parStruct);

					// третья колонка структуры
					parStruct.Init();
					parStruct.header = "char";					// заголовoк-комментарий столбца
					parStruct.type = ldefin2d.CHAR_ATTR_TYPE;	// тип данных в столбце - см.ниже
					parStruct.def = "10";						// значение по умолчанию
					pStruct.ksAddArrayItem(-1, parStruct);

					parCol2.SetColumns(pStruct);	// массив неопределенной длины информации о колонках для записи
				}

				pCol.ksAddArrayItem(-1, parCol2);

				// третья  колонка
				parCol3.Init();
				parCol3.header = "string";					// заголовoк-комментарий столбца
				parCol3.type = ldefin2d.STRING_ATTR_TYPE;	// тип данных в столбце - см.ниже
				parCol3.key = 0;							// дополнительный признак, который позволит отличить две переменные с одинаковым типом
				parCol3.def = "text";						// значение по умолчанию

				pCol.ksAddArrayItem(-1, parCol3);

				kompas.ksMessageBoxResult();	// проверяем результат работы нашей функции

				// просмотрим массив колонок
				ShowColumns(pCol, false);	// функция пользователя

				kompas.ksMessageBoxResult();	// проверяем результат работы нашей функции

				// поменяем  колонки местами 2->1 1->3 3->2
				pCol.ksSetArrayItem(0, parCol2);
				pCol.ksSetArrayItem(2, parCol1);
				pCol.ksSetArrayItem(1, parCol3);

				// просмотрим массив колонок
				ShowColumns(pCol, false);	// функция пользователя

				kompas.ksMessageBoxResult();	// проверяем результат работы нашей функции

				pStruct.ksDeleteArray();
				pEnum.ksDeleteArray();
				pCol.ksDeleteArray();
			}
		}


		// Массив полилиний это массив массивов математических точек
		void PolyLineArray()
		{
			string buf = string.Empty;
			ksMathPointParam par = (ksMathPointParam)kompas.GetParamStruct((short)StructType2DEnum.ko_MathPointParam);
			//создадим массив точек
			ksDynamicArray arrPoint = (ksDynamicArray)kompas.GetDynamicArray(ldefin2d.POINT_ARR);
			//создадим массив полилиний
			ksDynamicArray arrPoly = (ksDynamicArray)kompas.GetDynamicArray(ldefin2d.POLYLINE_ARR);
			if (par != null && arrPoint != null && arrPoly != null)
			{
				//наполним массив точек
				par.x = 10;
				par.y = 10;
				arrPoint.ksAddArrayItem(-1, par);
				par.x = 100;
				par.y = 100;
				arrPoint.ksAddArrayItem(-1, par);
				par.x = 1000;
				par.y = 1000;
				arrPoint.ksAddArrayItem(-1, par);
				//добавим массив точек в массив полилиний
				arrPoly.ksAddArrayItem(-1, arrPoint);
				kompas.ksMessageBoxResult();

				//наполним массив точек
				arrPoint.ksClearArray();
				par.x = 20;
				par.y = 20;
				arrPoint.ksAddArrayItem(-1, par);
				par.x = 200;
				par.y = 200;
				arrPoint.ksAddArrayItem(-1, par);
				par.x = 2000;
				par.y = 2000;
				arrPoint.ksAddArrayItem(-1, par);
				//добавим массив точек в массив полилиний
				arrPoly.ksAddArrayItem(-1, arrPoint);
				kompas.ksMessageBoxResult();

				//наполним массив точек
				arrPoint.ksClearArray();
				par.x = 30;
				par.y = 30;
				arrPoint.ksAddArrayItem(-1, par);
				par.x = 300;
				par.y = 300;
				arrPoint.ksAddArrayItem(-1, par);
				par.x = 3000;
				par.y = 3000;
				arrPoint.ksAddArrayItem(-1, par);
				//добавим массив точек в массив полилиний
				arrPoly.ksAddArrayItem(-1, arrPoint);
				kompas.ksMessageBoxResult();

				int count = arrPoly.ksGetArrayCount();
				buf = string.Format("n = {0}", count);
				kompas.ksMessage(buf);

				arrPoly.ksGetArrayItem(1, arrPoint);
				//просмоьрим 2-ой элемент массива полилиний
				count = arrPoint.ksGetArrayCount();
				for (int i = 0; i < count; i ++)
				{
					arrPoint.ksGetArrayItem(i, par);
					buf = string.Format("i = {0}, x = {1}, y = {2}", i, par.x, par.y);
					kompas.ksMessage(buf);
				}

				//заменим у 2 -го элемента массива полилиний 2-ой элемент
				par.x = 50;
				par.y = 50;
				arrPoint.ksSetArrayItem(1, par);
				par.x = 500;
				par.y = 500;
				arrPoint.ksSetArrayItem(0, par);
				arrPoly.ksSetArrayItem(1, arrPoint);

				count = arrPoly.ksGetArrayCount();
				for (int i = 0; i < count; i++)
				{
					arrPoly.ksGetArrayItem(i, arrPoint);
					int n = arrPoint.ksGetArrayCount();
					for (int j = 0; j < n; j++)
					{
						arrPoint.ksGetArrayItem(j, par);
						buf = string.Format("i = {0}, j = {1}, x = {2}, y = {3}",
							i, j, par.x, par.y);
						kompas.ksMessage(buf);
					}
				}
				kompas.ksMessageBoxResult();
				arrPoint.ksDeleteArray();
				arrPoly.ksDeleteArray();
			}
		}


		// Массив неопределенной длины габаритных прямоугольников-(структура RectParam)
		void RectArray()
		{
			ksRectParam par = (ksRectParam)kompas.GetParamStruct((short)StructType2DEnum.ko_RectParam); // параметры прямоугольника
			ksDynamicArray arr = (ksDynamicArray)kompas.GetDynamicArray(ldefin2d.RECT_ARR);    // создать массив
			ksMathPointParam pBot = (ksMathPointParam)kompas.GetParamStruct((short)StructType2DEnum.ko_MathPointParam);
			ksMathPointParam pTop = (ksMathPointParam)kompas.GetParamStruct((short)StructType2DEnum.ko_MathPointParam);
			if (arr != null && par != null && pBot != null && pTop != null)
			{
				//наполнить массив
				pTop.x = 10;
				pTop.y = 10;
				par.SetpTop(pTop);
				pBot.x = 20;
				pBot.y = -10;
				par.SetpBot(pBot);
				arr.ksAddArrayItem(-1, par);

				pTop.x = 20;
				pTop.y = 50;
				par.SetpTop(pTop);
				pBot.x = 50;
				pBot.y = 10;
				par.SetpBot(pBot);
				arr.ksAddArrayItem(-1, par);

				pTop.x = 20;
				pTop.y = 150;
				par.SetpTop(pTop);
				pBot.x = 50;
				pBot.y = 110;
				par.SetpBot(pBot);
				arr.ksAddArrayItem(-1, par);
				kompas.ksMessageBoxResult();

				//просмотреть массив
				int n = arr.ksGetArrayCount();

				string buf = string.Empty;
				buf = string.Format("n = {0}", n);
				kompas.ksMessage(buf);

				for (int i = 0; i < n; i++)
				{
					arr.ksGetArrayItem(i, par);
					pBot = (ksMathPointParam)par.GetpBot();
					pTop = (ksMathPointParam)par.GetpTop();
					buf = string.Format("i = {0}, x1 = {1:0.#}, y1 = {2:0.#},\n x2 = {3:0.#}, y2 = {4:0.#}, ",
						i, pTop.x, pTop.y,
						pBot.x, pBot.y);
					kompas.ksMessage(buf);
				}

				//редактируем массив
				pTop.x = -20;
				pTop.y = -50;
				par.SetpTop(pTop);
				pBot.x = 20;
				pBot.y = -10;
				par.SetpBot(pBot);
				arr.ksSetArrayItem(1, par);

				pTop.x = 0;
				pTop.y = 0;
				par.SetpTop(pTop);
				pBot.x = 10;
				pBot.y = -20;
				par.SetpBot(pBot);
				arr.ksSetArrayItem(0, par);

				pTop.x = 5;
				pTop.y = 5;
				par.SetpTop(pTop);
				pBot.x = 25;
				pBot.y = 0;
				par.SetpBot(pBot);
				arr.ksAddArrayItem(-1, par);

				//просмотрим массив
				n = arr.ksGetArrayCount();
				for (int i = 0; i < n; i ++)
				{
					arr.ksGetArrayItem(i, par);
					pBot = (ksMathPointParam)par.GetpBot();
					pTop = (ksMathPointParam)par.GetpTop();
					buf = string.Format("i = {0}, x1 = {1:0.#}, y1 = {2:0.#},\n x2 = {3:0.#}, y2 = {4:0.#},",
						i, pTop.x, pTop.y, pBot.x, pBot.y);
					kompas.ksMessage(buf);
				}
				kompas.ksMessageBoxResult();

				//удалить массив
				arr.ksDeleteArray();
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
