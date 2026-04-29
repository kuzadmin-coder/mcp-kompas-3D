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
	// Класс Step9 - Размеры и другие технологические объекты
  [ClassInterface(ClassInterfaceType.AutoDual)]
	public class Step9
	{
		private KompasObject kompas;
		private ksDocument2D doc;

		// Имя библиотеки
		[return: MarshalAs(UnmanagedType.BStr)] public string GetLibraryName()
		{
			return "Step9 - Размеры и другие технологические объекты";
		}
		

		// Головная функция библиотеки
		public void ExternalRunCommand([In] short command, [In] short mode, [In, MarshalAs(UnmanagedType.IDispatch)] object kompas_)
		{
			kompas = (KompasObject) kompas_;
			if (kompas != null)
			{
				doc = (ksDocument2D)kompas.ActiveDocument2D();
				if (doc != null && doc.reference != 0)
				{
					switch (command)
					{
						case 1  : DrawLinDim();			break; //линейный размер
						case 2  : DrawAngDim();			break; //угловой размер
						case 3  : DrawRough();			break; //шероховатость
						case 4  : DrawLeader();			break; //линия выноски
						case 5  : DrawPosLeader();		break; //позиционная линия выноски
						case 6  : DrawBrandLeader();	break; //клеймение
						case 7  : DrawMarkerLeader();	break; //маркирование
						case 8  : DrawBase();			break; //обозначение базы
						case 9  : DrawCutLine();		break; //маркирование
						case 10 : DrawDiamDim();		break; //диаметральный размер
						case 11 : DrawRadDimt();		break; //радиальный размер
						case 12 : DrawRadBreakDimt();	break; //радиальный размер c изломом
						case 13 : DrawViewPointer();	break; //cтрелка вида
					}
				}
				else
					kompas.ksError("Документ не активизирован или \nне является листом/фрагментом");
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
					result = "Линейный размер";
					command = 1;
					break;
				case 2:
					result = "Угловой  размер";
					command = 2;
					break;
				case 3:
					result = "Шероховатость";
					command = 3;
					break;
				case 4:
					result = "Линия выноски";
					command = 4;
					break;
				case 5:
					result = "Позиционная линия выноски";
					command = 5;
					break;
				case 6:
					result = "Клеймение";
					command = 6;
					break;
				case 7:
					result = "Маркирование";
					command = 7;
					break;
				case 8:
					result = "Обозначение базы";
					command = 8;
					break;
				case 9:
					result = "Линия разреза/cечения";
					command = 9;
					break;
				case 10:
					result = "Диаметральный размер";
					command = 10;
					break;
				case 11:
					result = "Радиальный размер";
					command = 11;
					break;
				case 12:
					result = "Радиальный размер с изломом";
					command = 12;
					break;
				case 13:
					result = "Стрелка вида";
					command = 13;
					break;
				case 14:
					itemType = 3;			//ENDMENU
					command = -1;
					break;
			}

			return result;
		}


		void DrawLinDim() 
		{
			ksLDimParam param = (ksLDimParam)kompas.GetParamStruct((short)StructType2DEnum.ko_LDimParam);
			if (param == null)
				return;

			ksDimDrawingParam dPar = (ksDimDrawingParam)param.GetDPar();
			ksLDimSourceParam sPar = (ksLDimSourceParam)param.GetSPar();
			ksDimTextParam tPar = (ksDimTextParam)param.GetTPar();

			if (dPar == null || sPar == null || tPar == null)
				return;

			dPar.Init();
			dPar.textPos = 10;
			dPar.textBase = 2;
			dPar.pt1 = 2;
			dPar.pt2 = 2;
			dPar.ang = -30;
			dPar.lenght = 20;

			sPar.Init();
			sPar.x1 = 50;
			sPar.y1 = 50;
			sPar.x2 = 70;
			sPar.y2 = 60;
			sPar.dx = 0;
			sPar.dy = -20;
			sPar.basePoint = 1;

			tPar.Init(false);
			tPar.SetBitFlagValue(ldefin2d._AUTONOMINAL, true);
			tPar.SetBitFlagValue(ldefin2d._PREFIX, true);
			tPar.SetBitFlagValue(ldefin2d._DEVIATION, true);
			tPar.SetBitFlagValue(ldefin2d._UNIT, true);
			tPar.SetBitFlagValue(ldefin2d._SUFFIX, true);
			tPar.sign = 0;

			ksChar255 str = (ksChar255)kompas.GetParamStruct((short)StructType2DEnum.ko_Char255);
			ksDynamicArray arrText = (ksDynamicArray)tPar.GetTextArr();

			if (str == null || arrText == null)
				return;

			str.str = "prefix";
			arrText.ksAddArrayItem(-1, str);

			str.str = "+0,5";
			arrText.ksAddArrayItem(-1, str);

			str.str = "-0,5";
			arrText.ksAddArrayItem(-1, str);

			str.str = "mm";
			arrText.ksAddArrayItem(-1, str);

			str.str = "pp&04ww&01oo";
			arrText.ksAddArrayItem(-1, str);

			int obj = doc.ksLinDimension(param);

			if (obj != 0) 
			{
				doc.ksGetObjParam(obj, param, ldefin2d.ALLPARAM);
				sPar.x2 = 50;
				sPar.y2 = 60;
				kompas.ksMessage(dPar.pl1 ? "Да" : "Нет"); 
				kompas.ksMessage("Поменяем параметры");
				doc.ksSetObjParam(obj, param, ldefin2d.ALLPARAM);
			}
		}

		void DrawAngDim() 
		{
			ksADimParam aDim = (ksADimParam)kompas.GetParamStruct((short)StructType2DEnum.ko_ADimParam);  
			ksTextLineParam textLine = (ksTextLineParam)kompas.GetParamStruct((short)StructType2DEnum.ko_TextLineParam);
			ksTextItemParam textItem = (ksTextItemParam)kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemParam);

			if (aDim == null || textLine == null || textItem == null)
				return;
    
			textLine.Init();
		
			textItem.Init();
			textItem.s = "Угловой размер";

			ksTextItemFont font = (ksTextItemFont)textItem.GetItemFont();
			ksDimTextParam tPar = (ksDimTextParam)aDim.GetTPar(); 
			ksADimSourceParam sPar = (ksADimSourceParam)aDim.GetSPar(); 
			ksDimDrawingParam dPar = (ksDimDrawingParam)aDim.GetDPar();
	
			if (font == null || tPar == null || sPar == null || dPar == null)
				return;

			dPar.Init();

			sPar.Init();
			sPar.rad = 50;

			tPar.Init(true);
	
			font.Init();
			font.height = 5;
			font.ksu = 1;
			font.fontName = "GOST type A";
			font.SetBitVectorValue(ldefin2d.NEW_LINE, true);

			ksDynamicArray arr = (ksDynamicArray)textLine.GetTextItemArr();
			if (arr == null)
				return;

			arr.ksAddArrayItem(-1, textItem);

			ksDynamicArray arr1 = (ksDynamicArray)tPar.GetTextArr();
			if (arr1 == null)
				return;

			arr1.ksAddArrayItem(-1, textLine);
		
			int obj = doc.ksAngDimension(aDim);
			if (obj != 0) 
			{
				doc.ksGetObjParam(obj, aDim, ldefin2d.ALLPARAM);
				sPar.Init();
				sPar.rad = 100;
				kompas.ksMessage("Поменяем параметры");
				doc.ksSetObjParam(obj, aDim, ldefin2d.ALLPARAM);
			}
		}

		void DrawRough() 
		{
			ksRoughParam roughPar = (ksRoughParam)kompas.GetParamStruct((short)StructType2DEnum.ko_RoughParam);
			ksRoughPar rPar = (ksRoughPar)roughPar.GetrPar();
			ksShelfPar shPar = (ksShelfPar)roughPar.GetshPar();
			ksChar255 str = (ksChar255)kompas.GetParamStruct((short)StructType2DEnum.ko_Char255);
			if (roughPar != null && rPar != null && shPar != null && str != null) 
			{
				rPar.Init();
				shPar.Init();
				str.Init();

				//заполним параметры текста шероховатости
				rPar.style = 0;
				rPar.type = 0;
				rPar.around = 0;
				rPar.x = 50;
				rPar.y = 50;
				rPar.ang = 90;
				rPar.cText0 = 2;
				rPar.cText1 = 2;
				rPar.cText2 = 2;
				rPar.cText3 = 1;

				//режим, когда тексты задаются массивом строк символов
				ksDynamicArray ptext = (ksDynamicArray)rPar.GetpText();
				if (ptext == null)
					return;
				str.str = "1";
				ptext.ksAddArrayItem(-1, str);
				str.str = "2";
				ptext.ksAddArrayItem(-1, str);
				str.str = "3";
				ptext.ksAddArrayItem(-1, str);
				str.str = "4";
				ptext.ksAddArrayItem(-1, str);
				str.str = "5";
				ptext.ksAddArrayItem(-1, str);
				str.str = "6";
				ptext.ksAddArrayItem(-1, str);
				str.str = "7";
				ptext.ksAddArrayItem(-1, str);

				//параметры выносной полки
				shPar.psh = 0;     //полки нет
				shPar.ang = 130;   //угол наклона ножки
				shPar.length = 20; //длина ножки

				int obj = doc.ksRough(roughPar);    

				if (obj != 0) 
				{
					doc.ksGetObjParam(obj, roughPar, ldefin2d.ALLPARAM);
					rPar.ang = 100;
					kompas.ksMessage("Поменяем параметры");
					doc.ksSetObjParam(obj, roughPar, ldefin2d.ALLPARAM);
				} 
			}
		}

		void DrawLeader() 
		{
			ksLeaderParam lead = (ksLeaderParam)kompas.GetParamStruct((short)StructType2DEnum.ko_LeaderParam);
			ksTextLineParam tLinePar = (ksTextLineParam)kompas.GetParamStruct((short)StructType2DEnum.ko_TextLineParam);
			ksTextItemParam ItemPar = (ksTextItemParam)kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemParam);
			ksTextItemFont tFont = (ksTextItemFont)ItemPar.GetItemFont();
			ksMathPointParam tMathPoint = (ksMathPointParam)kompas.GetParamStruct((short)StructType2DEnum.ko_MathPointParam);
			if (lead != null && tLinePar != null && ItemPar != null && tFont != null && tMathPoint != null) 
			{
				lead.Init();
				tLinePar.Init();
				ItemPar.Init();
				tFont.Init();
				tMathPoint.Init();

				tFont.SetBitVectorValue(ldefin2d.NEW_LINE, true);
				tLinePar.style = 0;

				ksDynamicArray pText = (ksDynamicArray)lead.GetpTextline();
				ksDynamicArray TextItemArr = (ksDynamicArray)tLinePar.GetTextItemArr();
				if (TextItemArr == null || pText == null)
					return;

				ItemPar.s = "1";
				TextItemArr.ksAddArrayItem(-1, ItemPar);
				pText.ksAddArrayItem(-1, tLinePar);

				TextItemArr.ksClearArray();
				ItemPar.s = "2";
				TextItemArr.ksAddArrayItem(-1, ItemPar);
				pText.ksAddArrayItem(-1, tLinePar);

				TextItemArr.ksClearArray();
				ItemPar.s = "3";
				TextItemArr.ksAddArrayItem(-1, ItemPar);
				pText.ksAddArrayItem(-1, tLinePar);
        
				ksDynamicArray pPolyLin = (ksDynamicArray)lead.GetpPolyline();
				ksDynamicArray pMathPoint = (ksDynamicArray)kompas.GetDynamicArray(ldefin2d.POINT_ARR);
				if (pPolyLin == null || pMathPoint == null)
					return;

				tMathPoint.x = 10;
				tMathPoint.y = 10;

				pMathPoint.ksAddArrayItem(-1, tMathPoint);
				pPolyLin.ksAddArrayItem(-1, pMathPoint);

				tMathPoint.x = 30;
				tMathPoint.y = 10;
				pMathPoint.ksClearArray();
				pMathPoint.ksAddArrayItem(-1, tMathPoint);
				pPolyLin.ksAddArrayItem(-1, pMathPoint);

				lead.SetpPolyline(pPolyLin);
        
				//заполним параметры 
				lead.x = 50;		// координаты базовой точки (начало полки)
				lead.y = 50;
				lead.arrowType = 1;	// тип стрелки
				lead.dirX = 1;		// направление полки по X (1 - вправо -1 - влево)
				lead.signType = 0;	// тип знака
				lead.around = 0;	// знак обработки по контуру 0 - выключен 1 - включен
				lead.cText0 = 1;	// количество строк текста над полкой 0 - текст отсутствует
				lead.cText1 = 1;	// количество строк текста под полкой 0 - текст отсутствует
				lead.cText2 = 0;	// количество строк текста над ножкой (не более 1 строки) 0 - текст отсутствует
				lead.cText3 = 1;	// количество строк текста под ножкой (не более 1 строки) 0 - текст отсутствует

				int obj = doc.ksLeader(lead);
				if (obj != 0) 
				{
					doc.ksGetObjParam(obj, lead, ldefin2d.ALLPARAM);
					lead.x = 100;
					kompas.ksMessage("Поменяем параметры");
					doc.ksSetObjParam(obj, lead, ldefin2d.ALLPARAM);
				} 
			}
		}

		void DrawPosLeader() 
		{
			ksPosLeaderParam lead = (ksPosLeaderParam)kompas.GetParamStruct((short)StructType2DEnum.ko_PosLeaderParam);
			ksTextLineParam tLinePar = (ksTextLineParam)kompas.GetParamStruct((short)StructType2DEnum.ko_TextLineParam);
			ksTextItemParam ItemPar = (ksTextItemParam)kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemParam);
			ksTextItemFont tFont = (ksTextItemFont)ItemPar.GetItemFont();
			ksMathPointParam tMathPoint = (ksMathPointParam)kompas.GetParamStruct((short)StructType2DEnum.ko_MathPointParam);
			if (lead != null && tLinePar != null && ItemPar != null && tFont != null && tMathPoint != null) 
			{
				lead.Init();
				tLinePar.Init();
				ItemPar.Init();
				tFont.Init();
				tMathPoint.Init();

				lead.style = ldefin2d.INDICATIN_TEXT_LINE_ARR;
				tFont.SetBitVectorValue(ldefin2d.NEW_LINE, true);
				tLinePar.style = 0;

				ksDynamicArray pText = (ksDynamicArray)lead.GetpTextline();
				ksDynamicArray TextItemArr = (ksDynamicArray)tLinePar.GetTextItemArr();
				if (pText == null || TextItemArr == null)
					return;
	  
				ItemPar.s = "1";
				TextItemArr.ksAddArrayItem(-1, ItemPar);
				pText.ksAddArrayItem(-1, tLinePar);


				TextItemArr.ksClearArray();
				ItemPar.s = "4";
				TextItemArr.ksAddArrayItem(-1, ItemPar);
				pText.ksAddArrayItem(-1, tLinePar);

        
				ksDynamicArray pPolyLin = (ksDynamicArray)lead.GetpPolyline();
				ksDynamicArray pMathPoint = (ksDynamicArray)kompas.GetDynamicArray(ldefin2d.POINT_ARR);
				if (pPolyLin == null || pMathPoint == null)
					return;

				tMathPoint.x = 10;
				tMathPoint.y = 10;

				pMathPoint.ksAddArrayItem(-1, tMathPoint);
				pPolyLin.ksAddArrayItem(-1, pMathPoint);

				tMathPoint.x = 30;
				tMathPoint.y = 10;
				pMathPoint.ksClearArray();
				pMathPoint.ksAddArrayItem(-1, tMathPoint);
				pPolyLin.ksAddArrayItem(-1, pMathPoint);

				lead.SetpPolyline(pPolyLin);
        
				//заполним параметры 
				lead.x = 50;// координаты базовой точки (начало полки)
				lead.y = 50;
				lead.arrowType = 1;
				lead.dirX = -1;

				int obj = doc.ksPositionLeader(lead);
				if (obj != 0) 
				{
					doc.ksGetObjParam(obj, lead, ldefin2d.ALLPARAM);
					lead.x = 100;
					kompas.ksMessage("Поменяем параметры");
					doc.ksSetObjParam(obj, lead, ldefin2d.ALLPARAM);
				} 
			}
		}

		void DrawBrandLeader() 
		{
			ksBrandLeaderParam lead = (ksBrandLeaderParam)kompas.GetParamStruct((short)StructType2DEnum.ko_BrandLeaderParam);
			ksTextLineParam tLinePar = (ksTextLineParam)kompas.GetParamStruct((short)StructType2DEnum.ko_TextLineParam);
			ksTextItemParam ItemPar = (ksTextItemParam)kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemParam);
			ksChar255 str = (ksChar255)kompas.GetParamStruct((short)StructType2DEnum.ko_Char255);
			ksMathPointParam tMathPoint = (ksMathPointParam)kompas.GetParamStruct((short)StructType2DEnum.ko_MathPointParam);
			ksDynamicArray pMathPoint = (ksDynamicArray)kompas.GetDynamicArray(ldefin2d.POINT_ARR);
			if (lead != null && tLinePar != null && ItemPar != null && tMathPoint != null && str != null && pMathPoint != null) 
			{
				lead.Init();
				tLinePar.Init();
				ItemPar.Init();
				str.Init();
				tMathPoint.Init();

				ksTextItemFont tFont = (ksTextItemFont)ItemPar.GetItemFont();
				ksDynamicArray ptext = (ksDynamicArray)lead.GetpTextline();
				ksDynamicArray pPolyLin = (ksDynamicArray)lead.GetpPolyline();
				if (ptext == null || tFont == null || pPolyLin == null)
					return;
				tFont.Init(); 

				lead.cText0 = 1;
  
				str.str = "1";
				ptext.ksAddArrayItem(-1, str);
				str.str = "2";
				ptext.ksAddArrayItem(-1, str);
				str.str = "3";
				ptext.ksAddArrayItem(-1, str);

				tMathPoint.x = 10;
				tMathPoint.y = 10;

				pMathPoint.ksAddArrayItem(-1, tMathPoint);
				pPolyLin.ksAddArrayItem(-1, pMathPoint);

				tMathPoint.x = 30;
				tMathPoint.y = 10;
				pMathPoint.ksClearArray();
				pMathPoint.ksAddArrayItem(-1, tMathPoint);
				pPolyLin.ksAddArrayItem(-1, pMathPoint);

				//заполним параметры 
				lead.x = 50;	// координаты базовой точки (начало полки)
				lead.y = 50;
				lead.arrowType = 1;
				lead.dirX = -1;
				lead.cText0 = 1;
				lead.cText1 = 1;
				lead.cText2 = 1;
        
				int obj = doc.ksBrandLeader(lead);
				if (obj != 0) 
				{
					doc.ksGetObjParam(obj, lead, ldefin2d.ALLPARAM);
					lead.x = 100;
					kompas.ksMessage("Поменяем параметры");
					doc.ksSetObjParam(obj, lead, ldefin2d.ALLPARAM);
				} 
			}
		}

		void DrawMarkerLeader() 
		{
			ksMarkerLeaderParam lead = (ksMarkerLeaderParam)kompas.GetParamStruct((short)StructType2DEnum.ko_MarkerLeaderParam);
			ksTextLineParam tLinePar = (ksTextLineParam)kompas.GetParamStruct((short)StructType2DEnum.ko_TextLineParam);
			ksTextItemParam ItemPar = (ksTextItemParam)kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemParam);
			ksMathPointParam tMathPoint = (ksMathPointParam)kompas.GetParamStruct((short)StructType2DEnum.ko_MathPointParam);
			ksChar255 str = (ksChar255)kompas.GetParamStruct((short)StructType2DEnum.ko_Char255);
			if (lead != null && tLinePar != null && ItemPar != null && tMathPoint != null && str != null) 
			{
				lead.Init();
				tLinePar.Init();
				ItemPar.Init();
				tMathPoint.Init();
				str.Init();

				//заполним параметры 
				lead.x = 50;// координаты базовой точки (начало полки)
				lead.y = 50;
				lead.arrowType = 1;
				lead.cText0 = 1;
				lead.style1 = 0;

				ksDynamicArray pCharArr = (ksDynamicArray)lead.GetpTextline();
				ksDynamicArray pPolyLin = (ksDynamicArray)lead.GetpPolyline();
				ksDynamicArray pMathPoint = (ksDynamicArray)kompas.GetDynamicArray(ldefin2d.POINT_ARR);
				ksTextItemFont tFont = (ksTextItemFont)ItemPar.GetItemFont();
				if (tFont == null || pCharArr == null || pPolyLin == null)
					return;
				tFont.Init(); 

				str.str = "1";
				pCharArr.ksAddArrayItem(-1, str);
        
				tMathPoint.x = 10;
				tMathPoint.y = 10;

				pMathPoint.ksAddArrayItem(-1, tMathPoint);
				pPolyLin.ksAddArrayItem(-1, pMathPoint);

				tMathPoint.x = 30;
				tMathPoint.y = 10;
				pMathPoint.ksClearArray();
				pMathPoint.ksAddArrayItem(-1, tMathPoint);
				pPolyLin.ksAddArrayItem(-1, pMathPoint);

				int obj = doc.ksMarkerLeader(lead);
				if (obj != 0) 
				{
					doc.ksGetObjParam(obj, lead, ldefin2d.ALLPARAM);
					lead.x = 100;
					kompas.ksMessage("Поменяем параметры");
					doc.ksSetObjParam(obj, lead, ldefin2d.ALLPARAM);
				} 
			}
		}

		void DrawBase() 
		{
			ksBaseParam par = (ksBaseParam)kompas.GetParamStruct((short)StructType2DEnum.ko_BaseParam);
			if (par != null) 
			{
				par.style = 0;
				par.type = false; // строка
				par.x1 = 10;
				par.y1 = 10;
				par.x2 = 30;
				par.y2 = 40;
				par.str = "Это база";
				reference bas = doc.ksBase(par);
				par.Init();
				if (bas != 0) 
				{
					doc.ksGetObjParam(bas, par, ldefin2d.ALLPARAM);
					par.x2 = -30;
					kompas.ksMessage("Поменяем параметры");
					doc.ksSetObjParam(bas, par, ldefin2d.ALLPARAM);
				}
			}
		}

		void DrawCutLine() 
		{
			ksCutLineParam cut = (ksCutLineParam)kompas.GetParamStruct((short)StructType2DEnum.ko_CutLineParam);
			ksTextLineParam tLinePar = (ksTextLineParam)kompas.GetParamStruct((short)StructType2DEnum.ko_TextLineParam);
			ksTextItemParam ItemPar = (ksTextItemParam)kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemParam);
			ksTextItemFont tFont = (ksTextItemFont)ItemPar.GetItemFont();
			ksMathPointParam tMathPoint = (ksMathPointParam)kompas.GetParamStruct((short)StructType2DEnum.ko_MathPointParam);
			if (cut != null && tLinePar != null && ItemPar != null && tFont != null && tMathPoint != null) 
			{
				cut.Init();
				tLinePar.Init();
				ItemPar.Init();
				tFont.Init();
				tMathPoint.Init();

				cut.type = 0;
				cut.x1 = 30;
				cut.y1 = 65;
				cut.x2 = 95;
				cut.y2 = 15;
				cut.right = 1;
				cut.str = "A$;1$";
				ksDynamicArray pMathPoint = (ksDynamicArray)cut.GetpMathPoint();

				tMathPoint.x = 50;
				tMathPoint.y = 50;

				pMathPoint.ksAddArrayItem(-1, tMathPoint);

				tMathPoint.x = 50;
				tMathPoint.y = 30;

				pMathPoint.ksAddArrayItem(-1, tMathPoint);

				tMathPoint.x = 80;
				tMathPoint.y = 30;

				pMathPoint.ksAddArrayItem(-1, tMathPoint);
        
				int obj = doc.ksCutLine(cut);
				if (obj != 0) 
				{
					doc.ksGetObjParam(obj, cut, ldefin2d.ALLPARAM);
 		  
					pMathPoint.ksClearArray();

					tMathPoint.x = 30;
					tMathPoint.y = 50;

					pMathPoint.ksAddArrayItem(-1, tMathPoint);

					tMathPoint.x = 30;
					tMathPoint.y = 30;

					pMathPoint.ksAddArrayItem(-1, tMathPoint);

					tMathPoint.x = 80;
					tMathPoint.y = 30;

					pMathPoint.ksAddArrayItem(-1, tMathPoint);

					kompas.ksMessage("Поменяем параметры");
					doc.ksSetObjParam(obj, cut, ldefin2d.ALLPARAM);
				}
			}		
		}

		void DrawDiamDim() 
		{
			int cir = doc.ksCircle(100, 100, 50, 1); 
			ksRDimParam aDim = (ksRDimParam)kompas.GetParamStruct((short)StructType2DEnum.ko_RDimParam);  
			ksTextLineParam textLine = (ksTextLineParam)kompas.GetParamStruct((short)StructType2DEnum.ko_TextLineParam);
			ksTextItemParam textItem = (ksTextItemParam)kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemParam);
			if (aDim == null || textLine == null || textItem == null ) 
				return;

			textLine.Init();
			textItem.Init();

			ksDimTextParam tPar = (ksDimTextParam)aDim.GetTPar(); 
			ksTextItemFont font = (ksTextItemFont)textItem.GetItemFont();
			ksDynamicArray arr = (ksDynamicArray)textLine.GetTextItemArr();
			ksRDimSourceParam sPar = (ksRDimSourceParam)aDim.GetSPar(); 
			ksRDimDrawingParam dPar = (ksRDimDrawingParam)aDim.GetDPar();
			if (tPar == null || font == null || sPar == null || dPar == null)
				return;

			tPar.Init(true);
			tPar.SetBitFlagValue(ldefin2d._AUTONOMINAL, true);
			tPar.sign = 1;

			font.Init();
			font.height = 5;
			font.ksu = 1;
			font.fontName = "GOST type A";
			font.SetBitVectorValue(ldefin2d.NEW_LINE, true);

			arr.ksAddArrayItem(-1, textItem);

			ksDynamicArray arr1 = (ksDynamicArray)tPar.GetTextArr();
			if (arr1 == null) 
				return;
			arr1.ksAddArrayItem(-1, textLine);

			sPar.Init();
			sPar.xc = 100;
			sPar.yc = 100;
			sPar.rad =50;
		
			dPar.Init();
			dPar.textPos = 75;
			dPar.pt1 = 2;
			dPar.pt2 = 2;
			dPar.shelfDir = 1;
			dPar.ang = -30;

			int obj = doc.ksDiamDimension(aDim);
			if (obj != 0) 
			{
				doc.ksGetObjParam(obj, aDim, ldefin2d.ALLPARAM);
				sPar.rad = 100;
				kompas.ksMessage("Поменяем параметры");
				doc.ksDeleteObj(cir);
				doc.ksCircle(100, 100, 100, 1); 
				doc.ksSetObjParam(obj, aDim, ldefin2d.ALLPARAM);
			}
		}

		void DrawRadDimt() 
		{
			int cir = doc.ksCircle(100, 100, 50, 1); 
			ksRDimParam aDim = (ksRDimParam)kompas.GetParamStruct((short)StructType2DEnum.ko_RDimParam);  
			ksTextLineParam textLine = (ksTextLineParam)kompas.GetParamStruct((short)StructType2DEnum.ko_TextLineParam);
			ksTextItemParam textItem = (ksTextItemParam)kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemParam);
			if (aDim == null || textLine == null || textItem == null ) 
				return;
    
			textLine.Init();
			textItem.Init();

			ksDimTextParam tPar = (ksDimTextParam)aDim.GetTPar(); 
			ksTextItemFont font = (ksTextItemFont)textItem.GetItemFont();
			ksDynamicArray arr = (ksDynamicArray)textLine.GetTextItemArr();
			ksRDimSourceParam sPar = (ksRDimSourceParam)aDim.GetSPar(); 
			ksRDimDrawingParam dPar = (ksRDimDrawingParam)aDim.GetDPar();
			if (tPar == null || font == null || sPar == null || dPar == null)
				return;

			tPar.Init(true);
			tPar.SetBitFlagValue(ldefin2d._AUTONOMINAL, true);
			tPar.sign = 1;
		

			font.Init();
			font.height = 5;
			font.ksu = 1;
			font.fontName = "GOST type A";
			font.SetBitVectorValue(ldefin2d.NEW_LINE, true);

			arr.ksAddArrayItem(-1, textItem);

			ksDynamicArray arr1 = (ksDynamicArray)tPar.GetTextArr();
			if (arr1 == null)
				return;

			arr1.ksAddArrayItem(-1, textLine);

			sPar.Init();
			sPar.xc = 100;
			sPar.yc = 100;
			sPar.rad = 50;
		
			dPar.Init();
			dPar.textPos = 75;
			dPar.pt1 = 2;
			dPar.pt2 = 1;
			dPar.shelfDir = 1;
			dPar.ang = 30;

			int obj = doc.ksRadDimension(aDim);  
			if (obj != 0) 
			{
				doc.ksGetObjParam(obj, aDim, ldefin2d.ALLPARAM);
				sPar.rad = 100;
				kompas.ksMessage("Поменяем параметры");
				doc.ksDeleteObj(cir);
				doc.ksCircle(100, 100, 100, 1); 
				doc.ksSetObjParam(obj, aDim, ldefin2d.ALLPARAM);
			}
		}

		void DrawRadBreakDimt() 
		{
			int cir = doc.ksCircle(100, 100, 50, 1); 
			ksRBreakDimParam aDim = (ksRBreakDimParam)kompas.GetParamStruct((short)StructType2DEnum.ko_RBreakDimParam);  
			ksDimTextParam tPar = (ksDimTextParam)aDim.GetTPar(); 
			ksTextLineParam textLine = (ksTextLineParam)kompas.GetParamStruct((short)StructType2DEnum.ko_TextLineParam);
			ksTextItemParam textItem = (ksTextItemParam)kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemParam);
			ksTextItemFont font = (ksTextItemFont)textItem.GetItemFont();
			if (aDim != null && tPar != null && textLine != null && textItem != null && font != null) 
			{
				tPar.Init(true);
				textLine.Init();
				textItem.Init();
				font.Init();

				tPar.SetBitFlagValue(ldefin2d._AUTONOMINAL, true);
				tPar.sign = 1;

				font.height = 5;
				font.ksu = 1;
				font.fontName = "GOST type A";
				font.SetBitVectorValue(ldefin2d.NEW_LINE, true);

				ksDynamicArray arr = (ksDynamicArray)textLine.GetTextItemArr();
				ksDynamicArray arr1 = (ksDynamicArray)tPar.GetTextArr();
				ksRDimSourceParam sPar = (ksRDimSourceParam)aDim.GetSPar(); 
				ksRBreakDrawingParam dPar = (ksRBreakDrawingParam)aDim.GetDPar();
				if (sPar != null && dPar != null && arr != null && arr1 != null) 
				{
					sPar.Init();
					dPar.Init();

					arr.ksAddArrayItem(-1, textItem);
					arr1.ksAddArrayItem(-1, textLine);

					sPar.xc = 100;
					sPar.yc = 100;
					sPar.rad = 50;
			
					dPar.ang = 0;
					dPar.pb = 30;
					dPar.pt = 1;

					int obj = doc.ksRadBreakDimension(aDim);  
					if (obj != 0) 
					{
						doc.ksGetObjParam(obj, aDim, ldefin2d.ALLPARAM);
						sPar.rad = 100;
						kompas.ksMessage("Поменяем параметры");
						doc.ksDeleteObj(cir);
						doc.ksCircle(100, 100, 100, 1); 
						doc.ksSetObjParam(obj, aDim, ldefin2d.ALLPARAM);
					}
				}
			}
		}

		void DrawViewPointer() 
		{
			ksViewPointerParam viewPoint = (ksViewPointerParam)kompas.GetParamStruct((short)StructType2DEnum.ko_ViewPointerParam);
			ksTextItemParam ItemPar = (ksTextItemParam)kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemParam);
			if (viewPoint != null && ItemPar != null) 
			{
				ItemPar.Init();
				viewPoint.Init(); 
				ksTextItemFont tFont = (ksTextItemFont)ItemPar.GetItemFont();
				if (tFont != null) 
				{
					tFont.Init();

					viewPoint.x1 = 55;
					viewPoint.y1 = 50;
					viewPoint.x2 = 40;
					viewPoint.y2 = 50;
					viewPoint.xt = 40;
					viewPoint.yt = 52;
					viewPoint.type = 0;
					viewPoint.str = "стрелка";

					int p = doc.ksViewPointer(viewPoint);
					if (p != 0) 
					{
						doc.ksGetObjParam(p, viewPoint, ldefin2d.ALLPARAM);
						viewPoint.type = 0;
						viewPoint.str = "стрелка вида";
						kompas.ksMessage("Поменяем параметры");
						doc.ksSetObjParam(p, viewPoint, ldefin2d.ALLPARAM);
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
