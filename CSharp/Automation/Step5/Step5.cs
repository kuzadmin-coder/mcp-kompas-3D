using Kompas6API5;
using System;
using Microsoft.Win32;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Drawing;
using Kompas6Constants;
using KAPITypes;

using reference = System.Int32;

namespace Steps.NET
{
	// Класс Step5 - Редактирование
  [ClassInterface(ClassInterfaceType.AutoDual)]
	public class Step5
	{
		KompasObject kompas = null;

		// Имя библиотеки
		[return: MarshalAs(UnmanagedType.BStr)] public string GetLibraryName()
		{
			return "Step5 - Редактирование";
		}
		

		// Головная функция библиотеки
		public void ExternalRunCommand([In] short command, [In] short mode, [In, MarshalAs(UnmanagedType.IDispatch)] object kompas_)
		{
			kompas = (KompasObject) kompas_;
			if (kompas != null)
			{
				ksDocument2D doc = (ksDocument2D) kompas.ActiveDocument2D();
				if (doc != null && doc.reference != 0)
				{
					switch (command)
					{
						case 1  : DrawTransform(doc);         break; //трансформация объекта по матрице
						case 2  : DrawCopy(doc);              break; //копирование объекта
						case 3  : DrawSymmetry(doc);          break; //симметрия объекта
						case 4  : EditTolerance(doc);         break; //просмотр допуска формы
						case 5  : EditTable(doc);             break; //просмотр таблицы
						case 6  : EditStamp(doc);             break; //взять тексты граф и редактировать штамп
						case 7  : GetTextTT(doc);             break; //получить текст ТТ
						case 8	: ChangeTechnicalDemand(doc); break; //редактирование TT
						case 9	: ShowInsertFragment(doc);    break; //вставка фрагмента
						case 10 : EditFragmentLibrary(doc);   break; //работа с библиотекой фрагментов
						case 11 : ShowInsertFragment1(doc);   break; //вставка фрагмента россыпью
					}
				}
			}
		}


		// Формирование меню библиотеки
		public string ExternalMenuItem(short number, ref short itemType, ref short command)
		{
			string result = string.Empty;
			itemType = 1; // "MENUITEM"
			switch (number)
			{
				case 1:
					result = "Tрансформация объекта";
					command = 1;
					break;
				case 2:
					result = "Копия  объекта";
					command = 2;
					break;
				case 3:
					result = "Симметрия объекта";
					command = 3;
					break;
				case 4:
					result = "Просмотр и редактирование допуска формы";
					command = 4;
					break;
				case 5:
					result = "Просмотр и редактирование таблицы";
					command = 5;
					break;
				case 6:
					result = "Взять тексты граф и редактировать штамп";
					command = 6;
					break;
				case 7:
					result = "Получить текст ТТ";
					command = 7;
					break;
				case 8:
					result = "Редактировать ТТ";
					command = 8;
					break;
				case 9:
					result = "Вставка фрагмента из библиотеки фрагментов";
					command = 9;
					break;
				case 10:
					result = "Работа с библиотекой фрагментов";
					command = 10;
					break;
				case 11:
					result = "Вставка фрагмента россыпью";
					command = 11;
					break;
				case 12:
					command = -1;
					itemType = 3; // "ENDMENU"
					break;
			}
			return result;
		}


		void DrawTransform(ksDocument2D doc)
		{
			//трансформация  рабочей группы
			doc.ksMtr(-30, -30, 0, 1, 1);
			reference rf = doc.ksNewGroup(0);
			doc.ksLineSeg(30, 30, 60, 30, 1);
			doc.ksLineSeg(60, 30, 60, 60, 1);
			doc.ksLineSeg(60, 60, 30, 60, 1);
			doc.ksLineSeg(30, 60, 30, 30, 1);
			doc.ksHatch(0, 45, 2, 0, 0, 0);
			doc.ksLineSeg(30, 30, 60, 30, 1);
			doc.ksLineSeg(60, 30, 60, 60, 1);
			doc.ksLineSeg(60, 60, 30, 60, 1);
			doc.ksLineSeg(30, 60, 30, 30, 1);
			doc.ksEndObj();
			doc.ksEndGroup();
			doc.ksDeleteMtr();
  
			kompas.ksMessage("Создали матрицу 20, 20, 45, 2");

			doc.ksMtr(20, 20, 45, 2, 2);
			doc.ksTransformObj(rf);
			doc.ksDeleteMtr();

			kompas.ksMessageBoxResult();
			kompas.ksMessage("вернем обратно");

			doc.ksMtr(-20, -20, 0, 1, 1);
			doc.ksTransformObj(rf);
			doc.ksDeleteMtr();

			doc.ksMtr(0, 0, 0, 0.5, 0.5);
			doc.ksTransformObj(rf);
			doc.ksDeleteMtr();

			doc.ksMtr(0, 0, -45, 1, 1);
			doc.ksTransformObj(rf);
			doc.ksDeleteMtr();

			kompas.ksMessageBoxResult();
		}

		void DrawCopy(ksDocument2D doc)
		{
			ksViewParam par = (ksViewParam) kompas.GetParamStruct((short)StructType2DEnum.ko_ViewParam);

			if (par != null) 
			{
				par.Init();
				par.x = 20;
				par.y = 60;
				par.scale_ = 1 ;
				par.color = Color.FromArgb(0, 10,20,10).ToArgb();
				par.state = ldefin2d.stACTIVE;
				par.name = "user view";

				int number = 5;
				//создали вид
				reference v = doc.ksCreateSheetView(par, ref number);
				//создали слой
				doc.ksLayer(5);

				doc.ksLineSeg(20, 10, 20, 30, 1);
				doc.ksLineSeg(20, 30, 40, 30, 1);
				doc.ksLineSeg(40, 30, 40, 10, 1);
				doc.ksLineSeg(40, 10, 20, 10, 1);

				//копируем вид (для вида точки задаются в листовых координатах)
				doc.ksCopyObj(v, 20, 60, 40, 80, 1, 0);

				kompas.ksMessageBoxResult();
			}
		}

		void DrawSymmetry(ksDocument2D doc)
		{
			reference grp = doc.ksNewGroup(0);
			doc.ksLineSeg(20, 10, 20, 30, 1);
			doc.ksLineSeg(20, 30, 40, 30, 1);
			doc.ksLineSeg(40, 30, 40, 10, 1);
			doc.ksLineSeg(40, 10, 20, 10, 1);
			doc.ksEndGroup();

			doc.ksSymmetryObj(grp, 40, 10, 40, 20, "1");

			kompas.ksMessageBoxResult();
		}

		void EditTolerance(ksDocument2D doc)
		{
			// редактирование допуска формы
			reference pObj;

			ksRequestInfo info = (ksRequestInfo) kompas.GetParamStruct((short)StructType2DEnum.ko_RequestInfo);
			if (info == null) 
				return;
			double x = 0, y = 0;
			info.prompt = "Укажите допуск формы" ;
			int cursor = doc.ksCursor(info, ref x ,ref y, 0);
			if (cursor != 0) 
			{
				if (doc.ksExistObj(pObj = doc.ksFindObj(x, y, 1000000)) == 1)
				{
					//узнаем тип объекта
					int type = doc.ksGetObjParam(pObj, 0, 0);   //указатель на графический объект
					if (type == ldefin2d.TOLERANCE_OBJ) 
					{
						int numb = 0;
						string buf = string.Empty;
						//открыть допуск формы для редактирования
						doc.ksOpenTolerance(pObj);

						ksToleranceParam par = (ksToleranceParam) kompas.GetParamStruct((short)StructType2DEnum.ko_ToleranceParam);
						//параметры допуска формы
						doc.ksGetObjParam( pObj,	//указатель на графический объект
							par,					//указатель на структуру параметров
							ldefin2d.ALLPARAM);		//тип считывания параметров

						buf = string.Format("базовая точка = {0} стиль = {1} расположение - {2}\nx = {3:0.##} y = {4:0.##}",
							par.tBase, par.style,
							par.type != 0 ? "вертикальное" : "горизонтальное",
							par.x, par.y);
						kompas.ksMessage(buf);

						ksTextLineParam par1 = (ksTextLineParam) kompas.GetParamStruct((short)StructType2DEnum.ko_TextLineParam);

						if (par1 == null)
							return;
				
						par1.Init();
        
						//в цикле будем брать все существующие ячейки
						while (doc.ksGetToleranceColumnText(ref numb, par1) != 0) 
						{
							buf = string.Format("numb = {0}\nstyle = {1}", numb, par1.style);
							kompas.ksMessage(buf);

							ksDynamicArray arrpTextItem = (ksDynamicArray) par1.GetTextItemArr();
							ksTextItemParam item = (ksTextItemParam) kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemParam);

							if (item == null || arrpTextItem == null)
								return;
          
							item.Init();

							for (int i = 0; i < arrpTextItem.ksGetArrayCount(); i ++) 
							{
								arrpTextItem.ksGetArrayItem(i, item);
								ksTextItemFont textItemFont = (ksTextItemFont) item.GetItemFont();
								if (item.type == 0)
									buf = string.Format("компонента = {0} h = {1:0.##}\ns = {2}\nfontName = {3}\nбитвектор = {4}",
										i, textItemFont.height, item.s, textItemFont.fontName, textItemFont.bitVector);
								else
									buf = string.Format("компонента = {0} тип = {1} номер спецзнака = {3}", i, item.type, item.iSNumb);
								item.s = "Допуск формы";
								kompas.ksMessage(buf);
							}
							arrpTextItem.ksClearArray();
							arrpTextItem.ksAddArrayItem(-1, item);
							doc.ksSetToleranceColumnText(numb, par1);
							arrpTextItem.ksDeleteArray();  //очистим массив компонент
						}
						//заменим параметры
						par.x =  par.x + 10 ;
						par.y = par.y + 10 ;
						doc.ksSetObjParam(pObj,	//указатель на графический объект
							par,				//указатель на структуру параметров
							ldefin2d.ALLPARAM);	//тип считывания параметров
						doc.ksEndObj();			//закрыли объект "допуск формы"
					}
					else
						kompas.ksError("Это не допуск формы");
				}
				else
					kompas.ksError("Нет объекта");
			}

			kompas.ksMessageBoxResult();
		}

		void EditTable(ksDocument2D doc)
		{
			// редактирование таблицы
			reference pObj;

			ksRequestInfo info = (ksRequestInfo) kompas.GetParamStruct((short)StructType2DEnum.ko_RequestInfo);
			if (info == null) 
				return;
			double x = 0, y = 0;
			info.prompt = "Укажите таблицу";
			// взять таблицу на чертеже
			int cursor = doc.ksCursor(info, ref x, ref y, 0);
			if (cursor != 0) 
			{
				if(doc.ksExistObj(pObj = doc.ksFindObj(x, y, 100000)) == 1)
				{
					//узнаем тип объекта
					int type = doc.ksGetObjParam(pObj, 0, 0);	//указатель на графический объект
					//проверить полученный объкет  - таблица
					if (type == ldefin2d.TABLE_OBJ) 
					{
						int numb = 0;
						//reference p;
						string buf = string.Empty;
						//открыть таблицу для редактирования
						doc.ksOpenTable(pObj);

						ksTextParam par = (ksTextParam) kompas.GetParamStruct((short)StructType2DEnum.ko_TextParam);
						if (par == null)
							return;
						par.Init();
						//в цикле будем брать все существующие ячейки
						while (doc.ksGetTableColumnText(ref numb, par)!=0) 
						{
							buf = string.Format("numb = {0}", numb);
							kompas.ksMessage(buf);


							ksDynamicArray arrpLineText = (ksDynamicArray) par.GetTextLineArr() ;
							ksTextLineParam itemLineText = (ksTextLineParam) kompas.GetParamStruct((short)StructType2DEnum.ko_TextLineParam);
							if (itemLineText == null)
								return;
							itemLineText.Init();

							for (int i = 0; i < arrpLineText.ksGetArrayCount(); i++) 
							{
								arrpLineText.ksGetArrayItem(i, itemLineText);
								buf = string.Format("i = {0} style = {1}", i, itemLineText.style);
								kompas.ksMessage(buf);

								ksDynamicArray arrpTextItem = (ksDynamicArray) itemLineText.GetTextItemArr() ;
								ksTextItemParam item = (ksTextItemParam) kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemParam) ;
								if (item == null || arrpTextItem == null)
									return;
								item.Init();
          
								for (int j = 0; j < arrpTextItem.ksGetArrayCount(); j ++) 
								{
									arrpTextItem.ksGetArrayItem(j, item);
									ksTextItemFont textItemFont = (ksTextItemFont) item.GetItemFont() ;

									if (item.type == 0)
										buf = string.Format("Компонента = {0} h = {1:0.##}\ns = {2}\nfontName = {3}\nбитвектор = {4}",
											j, textItemFont.height, item.s, textItemFont.fontName, textItemFont.bitVector);
									else
										buf = string.Format("компонента = {0} тип = {1} номер спецзнака = {2}",
											j, item.type, item.iSNumb);
									kompas.ksMessage(buf);
								}
								arrpTextItem.ksDeleteArray();  //очистим массив компонент
							}
							//очистим массив текстовых строк
							arrpLineText.ksDeleteArray();
						}

						//берем ячейку 2
						doc.ksColumnNumber(2);
						doc.ksText(0, 0, 0, 5, 1, 0, "вторая ячейка");

						doc.ksDivideTableItem(3, true, 2);
						doc.ksColumnNumber(4);
						doc.ksText(0, 0, 0, 5, 1, 0, "4");

						doc.ksEndObj();	//закрыли объект "таблица"
					}
					else
						kompas.ksError("Это не таблица");
				}
				else
					kompas.ksError("Нет объекта");
			}

			kompas.ksMessageBoxResult();
		}

		void EditStamp(ksDocument2D doc)
		{
			ksStamp stamp = (ksStamp) doc.GetStamp() ;
			if (stamp != null && stamp.ksOpenStamp() == 1) 
			{
				int numb = 0;
				//в цикле будем брать все существующие графы
				ksDynamicArray arr = (ksDynamicArray) stamp.ksGetStampColumnText(ref numb);
				ksTextItemParam item = null;
				while (numb != 0 && arr != null) 
				{
					string buf = string.Empty;
					buf = string.Format("numb = {0}", numb);
					kompas.ksMessage(buf);

					ksDynamicArray arrpLineText = (ksDynamicArray) kompas.GetDynamicArray(ldefin2d.TEXT_LINE_ARR);
					ksTextLineParam itemLineText = (ksTextLineParam) kompas.GetParamStruct((short)StructType2DEnum.ko_TextLineParam);
					if (itemLineText == null)
						return;
					itemLineText.Init();

					for(int i = 0, count = arr.ksGetArrayCount(); i < count; i++) 
					{
						arr.ksGetArrayItem(i, itemLineText);
						buf = string.Format("i = {0} style = {1}", i, itemLineText.style);
						kompas.ksMessage(buf);

						ksDynamicArray arrpTextItem = (ksDynamicArray) itemLineText.GetTextItemArr() ;
						item = (ksTextItemParam) kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemParam);

						if (item == null || arrpTextItem == null)
							return;
						item.Init();
          
						for (int j=0, count2 = arrpTextItem.ksGetArrayCount(); j < count2; j++) 
						{
							arrpTextItem.ksGetArrayItem(j, item);
							ksTextItemFont textItemFont = (ksTextItemFont) item.GetItemFont() ;
							buf = string.Format("Компонента = {0} h = {1:0.##}\ns = {2}\nfontName = {3}",
								j, textItemFont.height, item.s, textItemFont.fontName);
							kompas.ksMessage(buf);
						}
						arrpTextItem.ksDeleteArray();  //очистим массив компонент
					}
					//очистим массив текстовых строк
					arrpLineText.ksDeleteArray();

					arr.ksDeleteArray();
					arr = (ksDynamicArray) stamp.ksGetStampColumnText(ref numb);
				}
				//заменим  графу 2
				doc.ksColumnNumber(2);
				item = (ksTextItemParam) kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemParam) ;
				if (item == null) 
				{
					stamp.ksCloseStamp();
					return;
				}
				item.Init();
				ksTextItemFont itemFont = (ksTextItemFont) item.GetItemFont() ;
				if (item != null && itemFont != null) 
				{
					itemFont.SetBitVectorValue(ldefin2d.NEW_LINE, true);
					item.s = "Графа 2";
					doc.ksTextLine(item);
				}
				stamp.ksCloseStamp();
			}
			else
				kompas.ksError ("Штамп не найден");

			kompas.ksMessageBoxResult();
		}

		void GetTextTT(ksDocument2D doc)
		{
			//получим указатель на технические трбования
			reference pTT = doc.ksGetReferenceDocumentPart(1);
			ksTechnicalDemandParam technicalDemandParam = (ksTechnicalDemandParam) kompas.GetParamStruct((short)StructType2DEnum.ko_TechnicalDemandParam) ;
			if (pTT != 0 && technicalDemandParam != null) 
			{
				technicalDemandParam.Init();
				//получим параметры описания ТТ
				doc.ksGetObjParam(pTT, technicalDemandParam, ldefin2d.TECHNICAL_DEMAND_PAR);
				string buf = string.Empty;
				ksDynamicArray pGab = (ksDynamicArray) technicalDemandParam.GetPGab();
				int count = pGab.ksGetArrayCount();
				buf = string.Format("стиль = {0} число страниц  TT = {1}",
					technicalDemandParam.style, count);
				kompas.ksMessage(buf);

				// создадим массив текстовых строк
				ksDynamicArray pTextLine = (ksDynamicArray) kompas.GetDynamicArray(ldefin2d.TEXT_LINE_ARR);
				//пройдемся по листам ТТ и получим текст
				for(int i = 0; i < count; i++) 
				{
					doc.ksGetObjParam(pTT, pTextLine, i);//ALLPARAM);
					ksTextLineParam itemLineText = (ksTextLineParam) kompas.GetParamStruct((short)StructType2DEnum.ko_TextLineParam) ;
					if (itemLineText == null)
						return;
					itemLineText.Init();

					//      TextLineParam par2;
					for(int i1 = 0, count1 = pTextLine.ksGetArrayCount(); i1 < count1; i1 ++) 
					{
						pTextLine.ksGetArrayItem(i1, itemLineText);
						buf = string.Format("компонента = {0} style = {1}",
							i1, itemLineText.style);
						kompas.ksMessage(buf);

						ksDynamicArray arrpTextItem = (ksDynamicArray) itemLineText.GetTextItemArr() ;
						ksTextItemParam item = (ksTextItemParam) kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemParam) ;

						if (item == null || arrpTextItem == null)
							return;
						item.Init();

						for (int j = 0, count2 = arrpTextItem.ksGetArrayCount(); j < count2; j ++) 
						{
							arrpTextItem.ksGetArrayItem(j, item);
							ksTextItemFont textItemFont = (ksTextItemFont) item.GetItemFont() ;
							buf = string.Format("Компонента = {0} h = {1:0.##}\ns = {2}\nfontName = {3}",
								j, textItemFont.height, item.s, textItemFont.fontName);
							kompas.ksMessage(buf);
						}
						arrpTextItem.ksDeleteArray(); //очистим массив компонент
					}
				}
				pTextLine.ksDeleteArray(); //очистим массив текстовых строк
			}

			kompas.ksMessageBoxResult();
		}

		void ShowInsertFragment(ksDocument2D doc)
		{
			string libName = string.Empty;
			int res = 0;
			//выберем  библиотеку фрагментов
			libName = kompas.ksChoiceFile("*.lfr", "Библиотки фрагментов(*.lfr)|*.lfr|Все файлы (*.*)|*.*|", true) ;
			if (libName.Length > 0)
			{ 
				string buf = string.Empty; 
				do 
				{
					//выбрать фрагмент в библиотеке фрагментов
					ksFragment fr = (ksFragment) doc.GetFragment() ;
					// ABB K6    ksFragmentLibrary
					ksFragmentLibrary frLib = (ksFragmentLibrary) kompas.GetFragmentLibrary() ; 
					if  (fr == null && frLib == null)
						return;
					buf = frLib.ksChoiceFragmentFromLib(libName, out res) ;
					if (buf != null && res == 3) // res = 3 - выбран фрагмент
					{
						// выделим имя вставки 
						string insertName = buf.Substring(buf.IndexOf('|'));
						if (insertName != null && insertName != string.Empty) 
						{
							double x = 0, y = 0;
							//подготовим структуры фанома и запросов для Placement
							ksPhantom rub = (ksPhantom) kompas.GetParamStruct((short)StructType2DEnum.ko_Phantom) ;
							rub.phantom = 1;
							ksType1 type1 = (ksType1) rub.GetPhantomParam() ;
							if (rub == null || type1 == null)
								return;

							type1.Init(); rub.Init(); 

							type1.scale_ = 1;
							rub.phantom = 1;

							reference pDefFrg;
							// создадим описание  всавки фрагментов
							pDefFrg = fr.ksFragmentDefinition(buf,	//имя файла фрагмента
								insertName + 1,						//имя вставки
								1);									//тип вставки -действителен для внешнего фрагмента
																	// 0- взять в документ, 1-внешней ссылкой

							if(pDefFrg > 0) 
							{
								//во временную группу положим вставку фрагмента, взятую из библиотеки фрагментов 
								type1.gr = doc.ksNewGroup(1);
   
								ksPlacementParam par = (ksPlacementParam) kompas.GetParamStruct((short)StructType2DEnum.ko_PlacementParam) ;
								if (par == null)
									return; 
								par.Init();
								par.scale_ = 1;
								reference  p = fr.ksInsertFragment(pDefFrg, false, par);

								doc.ksEndGroup();
								int j;
								do 
								{
									type1.angle = 0;
									double ang = type1.angle;
									if ((j = doc.ksPlacement(null, ref x, ref y, ref ang, rub))!=0) 
									{
										type1.angle = ang;
										doc.ksCopyObj(p,	// указатель на графический объект
											0, 0,			// базовая точка объекта
											x, y,			// точка куда копировать
											1, type1.angle);// масштаб и угол поворота а градусах
									}
								}
								while (j > 0); 
								doc.ksDeleteObj(type1.gr);
							}
							else
								kompas.ksError ("Ошибка создания описания вставки фрагмента");
						}
						else
							kompas.ksError("Имя вставки не определено");
					}
				} while(res > 0);
			}

			kompas.ksMessageBoxResult();
		}

		void ShowInsertFragment1(ksDocument2D doc)
		{
			string frwName = string.Empty;
			ksFragment fr = (ksFragment) doc.GetFragment() ;
			if (fr == null)
				return;
			//выберем  фрагмент
			frwName = kompas.ksChoiceFile("*.frw", "фрагменты(*.frw)|*.frw|Все файлы (*.*)|*.*|", true);
			if(frwName != null && frwName != string.Empty) 
			{ 
				double x = 0, y = 0;
				//подготовим структуры фанома и запросов для Placement
				ksPhantom rub = (ksPhantom) kompas.GetParamStruct((short)StructType2DEnum.ko_Phantom);
				rub.phantom = 1;
				ksType1 type1 = (ksType1) rub.GetPhantomParam();
				if (rub == null || type1 == null)
					return;
				type1.Init();
				rub.Init(); 

				type1.scale_ = 1;
				rub.phantom = 1;

				//во временную группу положим вставку фрагмента, взятую из библиотеки фрагментов
				ksPlacementParam par = (ksPlacementParam) kompas.GetParamStruct((short)StructType2DEnum.ko_PlacementParam);
				if (par == null)
					return; 
				par.Init();
				par.scale_ = 1;

				int j;
				do 
				{
					//если нужно вставить несколько фрагментов, группу лучше порождать новую,
					//так как вместе с геометрией могут придти атрибуты, объекты спецификации, стили,
					//которые связаны с геометрией. При простом копировании группы эта связь будет
					//потеряна.
					type1.gr = fr.ksReadFragmentToGroup(frwName, false, par);
					if (type1.gr > 0) 
					{
						double ang = type1.angle;
						if ((j = doc.ksPlacement(null, ref x, ref y, ref ang, rub)) != 0) 
						{
							//сдвигаем группу
							doc.ksMoveObj(type1.gr, x, y);
							//поворачиваем группу
							if(Math.Abs(ang) > 0.001)
								doc.ksRotateObj(type1.gr, x, y, ang);
							//ставим группу в модель
							doc.ksStoreTmpGroup(type1.gr);
							doc.ksClearGroup(type1.gr, true);
							doc.ksDeleteObj(type1.gr);
						}
					}
					else 
					{
						if (type1.gr > 0)
							doc.ksDeleteObj(type1.gr);
						j = 0;
					}
				} while (j > 0);
			}

			kompas.ksMessageBoxResult();
		}

		void EditFragmentLibrary(ksDocument2D doc)
		{
			string libName = string.Empty;
			string buf = string.Empty;
			//выберем  библиотеку фрагментов
			libName =  kompas.ksChoiceFile("*.lfr", "Библиотки фрагментов(*.lfr)|*.lfr|Все файлы (*.*)|*.*|", true) ;
			if (libName != null && libName != string.Empty) 
			{ 
				ksRequestInfo info = (ksRequestInfo) kompas.GetParamStruct((short)StructType2DEnum.ko_RequestInfo) ;
				ksFragment fr = (ksFragment) doc.GetFragment() ;
				// АВВ К6 ksFragmentLibrary
				ksFragmentLibrary frLib = (ksFragmentLibrary) kompas.GetFragmentLibrary(); 
				if  (info == null || fr == null || frLib == null)
					return;
				info.Init();

				info.commandsString = "!Новый фрагмент !Редактировать фрагмент !Удалить фрагмент ";
				int j;
				int typeEdit = 0;
				string nameFrg = string.Empty;
				do 
				{
					j = doc.ksCommandWindow(info);
					switch (j) 
					{
						case 1:  //!Новый_фрагмент
							buf =  kompas.ksReadString("Введите имя нового фрагмента", "") ;
							if (buf != null && buf != string.Empty) 
							{
								nameFrg = libName;
								if (buf[0] != '|')
									nameFrg += "|";
								nameFrg += buf;
								typeEdit = 2; //запустить на редактирование
							}
							else
								typeEdit = 0;
							break;
						case 2 : //Редактировать_фрагмент
						case 3 : //Удалить_фрагмент
							//выберем имя файла фрагмента
							int res = 0;
							buf = frLib.ksChoiceFragmentFromLib(libName, out res);
							if (res > 0 && buf != null && buf != string.Empty && (j == 2 || j == 3))
							{
								nameFrg = buf;
								typeEdit = j; // 2- запустить на редактирование, 3-удалить ;
							}
							else
								typeEdit = 0;
							break;
					}

					if (j > 0 && typeEdit > 0) 
					{
						if (frLib.ksFragmentLibraryOperation(nameFrg, typeEdit) == 1) 
						{
							if (typeEdit == 2) 
							{
								frLib.ksFragmentLibraryOperation(nameFrg, 4 /*минимизировать окно библиотекаря*/);
								//редактируем фрагмент из библиотеки
								doc.ksText(0, 100,	//точка привязки текста
									0,				//угол наклона текста
									5,				//высота текста
									1,				//сужение текста
									0,				//свойства строки, которые задаются вкл.-выкл.
									"Редактируем фрагмент из библиотеки");	//строка символов

								doc.ksLineSeg (0, 100, 110, 100, 1);
								//редактируем фрагмент в интерактивном режиме 
								//после выбора в меню "Сервис" команды "Закончить редактирование фрагмента", 
								//возвращаемся в библиотеку
								if (kompas.ksSystemControlStart("Закончить редактирование фрагмента") != 1)
									kompas.ksStrResult();
								if (!kompas.ksSystemControlStop())
									kompas.ksStrResult();
								frLib.ksFragmentLibraryOperation(nameFrg, 0 /*закрыть c сохранением*/);
							}
						}  
						else
							kompas.ksMessageBoxResult();
					}

				} while (j != -1);
			}
		}

		void ChangeTechnicalDemand(ksDocument2D doc)
		{
			int rf = doc.ksGetReferenceDocumentPart(1);
			ksTechnicalDemandParam par = (ksTechnicalDemandParam) kompas.GetParamStruct((short)StructType2DEnum.ko_TechnicalDemandParam);
			if (doc.ksGetObjParam(rf, par, ldefin2d.TECHNICAL_DEMAND_PAR) != 0) 
			{
				string buf = string.Empty;
				buf = string.Format("число строк TT = {0}", par.strCount);
				kompas.ksMessage(buf);
	
				doc.ksOpenTechnicalDemand(par.GetPGab(), par.style);

				ksTextLineParam parLine = (ksTextLineParam) kompas.GetParamStruct((short)StructType2DEnum.ko_TextLineParam);
				if (parLine != null)
				{
					parLine.Init();
					for(int i = 0; i < par.strCount; i++) 
					{
						doc.ksGetObjParam(rf, parLine, ldefin2d.TT_FIRST_STR + i);
						ksTextItemParam parItem = (ksTextItemParam) kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemParam);
						ksDynamicArray arr = (ksDynamicArray) parLine.GetTextItemArr();
						if (parItem != null && arr != null) 
						{
							parItem.Init();
							for (int j = 0, count1 = arr.ksGetArrayCount(); j < count1; j++) 
							{
								arr.ksGetArrayItem(j, parItem);
								kompas.ksMessage(parItem.s);
								parItem.s = string.Format("{0}-я строка", i + 1);  
								arr.ksSetArrayItem(j, parItem);
							}
						}
						doc.ksSetObjParam(rf, parLine, ldefin2d.TT_FIRST_STR + i);
					}
				}
				doc.ksCloseTechnicalDemand();
			}

			kompas.ksMessageBoxResult();
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
