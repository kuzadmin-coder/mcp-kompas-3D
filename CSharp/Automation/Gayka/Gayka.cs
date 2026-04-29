using Kompas6API5;
using KompasAPI7;

using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Kompas6Constants;
using KAPITypes;

namespace Steps.NET
{
  [ClassInterface(ClassInterfaceType.AutoDual)]
	public class GaykaObj
	{
		const short ID_VID = 1;
		const short ID_SIDEVID = 2;
		const short ID_TOPVID = 3;
		const short ID_VIDSEC = 4;
		const double MODSTEP_REAL = 0.5412765877;	// коэффициент для расчета диаметра резьбы по шагу резьбы


		#region Название колонок для спецификации
		short SPC_NAME = 5;			// наименование
		short SPC_CLEAR_GEOM = 0;	// очищать геометрию при редактировании объекта спецификации
		short STANDART_SECTION = 25;
		#endregion
	
		#region Контролы панели свойств
		const short DIAM_ID = 10001;			// Комбобокс диаметров резьбы
		const short VIEWSIDE_ID = 10002;		// Кнопки отображения
		const short PERFORMANCE_ID = 10003;		// Кнопки исполнений
		const short SIMPLES_ID = 10004;			// Кнопки упрощений
		const short ADD_PARAM_ID = 10005;		// Кнопки доп. параметров
		const short SPC_CHECK_ID = 10006;		// Чекбокс объекта сп
		const short ANGLE_HATCH_ID = 10007;		// угол штриховки
		const short STEP_HATCH_ID = 10008;		// шаг штриховки
		const short PARAMS_ID = 10009;			// список параметров
		const short VIEW_BOX_ID = 10010;		// окно просмотра
		const short STR_EDIT_ID = 10011;
		const short STR_LIST_ID = 10012;
		#endregion
	
		#region Кнопки панели свойств
		short BASE_VIEW = 2001;		// Главный вид
		short LEFT_VIEW = 2002;		// Вид слева
		short TOP_VIEW = 2003;		// Вид сверху
		short SEC_VIEW = 2004;		// Вид\разрез
		short PERF1_VIEW = 2005;	// Исполнение 1
		short PERF2_VIEW = 2006;	// Исполнение 2
		short SIMPLE_VIEW = 2007;	// Упрошено
		short DRAW_AXIS = 2008;		// Рисовать ось
		short ADDSTEP_BUTT = 2009;	// Мелкий шаг
		short KEY_BUTT = 2010;		// Доп. размер под ключ
		#endregion
		
		public struct GaykaParam
		{
			public float dr;
			public float s;
			public float d;
			public float da;
			public float H;
			public float d2;
			public float p;
			public short cls;			// класс точности
			public short gost;			// номер госта
			public float hatchAng;		// угол штриховки
			public float hatchStep;		// шаг штриховки
			public float massa;			// масса
			public short indexMassa;	// 0-металл 1- алюмин сплав 2-латунь
			public byte perform;		// исполнение
			public bool simple;			// упрощенно
			public byte axis_off;		// выкл/вкл ось
			public bool pitch;			// мелкий шаг
			public byte pitch_off;		// мелкий шаг не возможен
			public bool key_s;			// дополнительный размер под ключ
			public byte key_s_on;		// дополнительный размер под ключ вкл
			public byte key_s_gray;		// дополнительный размер под ключ не возможен
			public byte koef_mat_on;	// коэф материала
			public short ver;			// версия макро
		}

		public struct BaseMakroParam
		{
			public float ang;			// угол поворота
			public short flagAttr;		// флаг создания объекта спецификации
			public short drawType;		// тип изображения
			public byte typeSwitch;		// тип запроса положения базовой точки элемента 0 Placement 1 Curso
										// 0 - точка + направление оси 0X ( Placement );
										// 1 - точка, направление совпадает с осью 0X текущей СК ( Cursor ).
		}

		public struct SimpleBase
		{
			public int bg;	// указатель базы
			public int rg;	// указатель отношения
		}


		#region Базовые параметры
		private BaseMakroParam par;			// базовые параметры
		private GaykaParam tmp;				// структура параметров гайки ГОСТ 5915-70
		private SimpleBase base_;			// референсы БД
	
		private int flagMode;				// true - режим редактирования
		private int refMacr;				// референс МАКРО
		private ksUserParam paramTmp;		// ksUserParam параметры для чтения БД
		private ksUserParam Param;			// ksUserParam параметры для Get/SetMacroParam
		private ksDataBaseObject data;		// объект для работы
		private int[] spcObj = new int[4];	// массив референсов объектов спецификации
		public short countObj;
		private bool flagSwitch;			// false переключения нет true переключение с
		// Placement на Cursor
		public HatchControl hatchPar;
		private PropertyUserControl userCtrl;
		public ProcessParam processParam;
		private PropertyList diamEdit;
		private PropertyMultiButton viewButt;
		private PropertyMultiButton varButt;
		private PropertyMultiButton simpButt;
		private PropertyMultiButton addButt;
		private PropertyCheckBox spcCheck;
		private PropertyEdit angleEdit;
		private PropertyEdit stepEdit;
		private ksPhantom phantom;
		#endregion

		/// <summary>
		/// Инициализация
		/// </summary>
		public void InitUserParam()
		{
			if (Param != null)
			{
				ksLtVariant item = (ksLtVariant)Kompas.Instance.KompasObject.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
				ksDynamicArray arr = (ksDynamicArray)Kompas.Instance.KompasObject.GetDynamicArray(ldefin2d.LTVARIANT_ARR);
				if (item != null && arr != null)
				{
					Param.Init();
					Param.SetUserArray(arr);
					item.Init();
					item.floatVal = par.ang;		// 0 - ang
					arr.ksAddArrayItem(-1, item);
					item.shortVal = par.flagAttr;	// 1 - flagAttr
					arr.ksAddArrayItem(-1, item);
					item.shortVal = par.drawType;	// 2 - drawType
					arr.ksAddArrayItem(-1, item);
					item.shortVal = par.typeSwitch;	// 3 - typeSwitch
					arr.ksAddArrayItem(-1, item);
				
					item.floatVal = tmp.dr;			// 4 - dr
					arr.ksAddArrayItem(-1, item);
					item.floatVal = tmp.s;			// 5 - s
					arr.ksAddArrayItem(-1, item);
					item.floatVal = tmp.d;			// 6 - D
					arr.ksAddArrayItem(-1, item);
					item.floatVal = tmp.da;			// 7 - da
					arr.ksAddArrayItem(-1, item);
					item.floatVal = tmp.H;			// 8 - h
					arr.ksAddArrayItem(-1, item);
					item.floatVal = tmp.d2;			// 9 - d2
					arr.ksAddArrayItem(-1, item);
					item.floatVal = tmp.p;			// 10 - p
					arr.ksAddArrayItem(-1, item);
					item.shortVal = tmp.cls;		// 11 - class
					arr.ksAddArrayItem(-1, item);
					item.shortVal = tmp.gost;		// 12 - gost
					arr.ksAddArrayItem(-1, item);
					item.floatVal = tmp.hatchAng;	// 13 - hatchAng
					arr.ksAddArrayItem(-1, item);
					item.floatVal = tmp.hatchStep;	// 14 - hatchShag
					arr.ksAddArrayItem(-1, item);
					item.floatVal = tmp.massa;		// 15 - m
					arr.ksAddArrayItem(-1, item);
					item.shortVal = tmp.indexMassa;	// 16 - indexMassa
					arr.ksAddArrayItem(-1, item);
					item.uCharVal = tmp.perform;	// 17 - perform
					arr.ksAddArrayItem(-1, item);
					item.uCharVal = Convert.ToInt16(tmp.simple);	// 18 - simple
					arr.ksAddArrayItem(-1, item);
					item.uCharVal = tmp.axis_off;	// 19 - axis_off
					arr.ksAddArrayItem(-1, item);
				
					item.uCharVal = Convert.ToInt16(tmp.pitch);		// 20 - pitch
					arr.ksAddArrayItem(-1, item);
					item.uCharVal = tmp.pitch_off;	// 21 - pitch_off
					arr.ksAddArrayItem(-1, item);
					item.uCharVal = Convert.ToInt16(tmp.key_s);		// 22 - key_s
					arr.ksAddArrayItem(-1, item);
					item.uCharVal = tmp.key_s_on;	// 23 - key_s_on
					arr.ksAddArrayItem(-1, item);
					item.uCharVal = tmp.key_s_gray;	// 24 - key_s_gray
					arr.ksAddArrayItem(-1, item);
					item.uCharVal = tmp.koef_mat_on;// 25 - koef_mat_on
					arr.ksAddArrayItem(-1, item);
					item.shortVal = tmp.ver;		// 26 - ver
					arr.ksAddArrayItem(-1, item);
				}
			}
		}
	

		/// <summary>
		/// Сохранить параметры
		/// </summary>
		public void SetUserParam()
		{
			if (Param != null)
			{
				ksLtVariant item = (ksLtVariant)Kompas.Instance.KompasObject.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
				ksDynamicArray arr = (ksDynamicArray)Param.GetUserArray();
				if (item != null && arr != null)
				{
					item.Init();
					item.floatVal = par.ang;			// 0 - ang
					arr.ksSetArrayItem(0, item);
					item.shortVal = par.flagAttr;		// 1 - flagAttr
					arr.ksSetArrayItem(1, item);
					item.shortVal = par.drawType;		// 2 - drawType
					arr.ksSetArrayItem(2, item);
					item.shortVal = par.typeSwitch;		// 3 - typeSwitch
					arr.ksSetArrayItem(3, item);
				
					item.floatVal = tmp.dr;				// 4 - dr
					arr.ksSetArrayItem(4, item);
					item.floatVal = tmp.s;				// 5 - s
					arr.ksSetArrayItem(5, item);
					item.floatVal = tmp.d;				// 6 - D
					arr.ksSetArrayItem(6, item);
					item.floatVal = tmp.da;				// 7 - da
					arr.ksSetArrayItem(7, item);
					item.floatVal = tmp.H;				// 8 - h
					arr.ksSetArrayItem(8, item);
					item.floatVal = tmp.d2;				// 9 - d2
					arr.ksSetArrayItem(9, item);
					item.floatVal = tmp.p;				// 10 - p
					arr.ksSetArrayItem(10, item);
					item.shortVal = tmp.cls;			// 11 - class
					arr.ksSetArrayItem(11, item);
					item.shortVal = tmp.gost;			// 12 - gost
					arr.ksSetArrayItem(12, item);
					item.floatVal = tmp.hatchAng;		// 13 - hatchAng
					arr.ksSetArrayItem(13, item);
					item.floatVal = tmp.hatchStep;		// 14 - hatchShag
					arr.ksSetArrayItem(14, item);
					item.floatVal = tmp.massa;			// 15 - m
					arr.ksSetArrayItem(15, item);
					item.shortVal = tmp.indexMassa;		// 16 - indexMassa
					arr.ksSetArrayItem(16, item);
					item.uCharVal = tmp.perform;		// 17 - perform
					arr.ksSetArrayItem(17, item);
					item.uCharVal = Convert.ToInt16(tmp.simple);	// 18 - simple
					arr.ksSetArrayItem(18, item);
					item.uCharVal = tmp.axis_off;		// 19 - axis_off
					arr.ksSetArrayItem(19, item);
				
					item.uCharVal = Convert.ToInt16(tmp.pitch);		// 20 - pitch
					arr.ksSetArrayItem(20, item);
					item.uCharVal = tmp.pitch_off;		// 21 - pitch_off
					arr.ksSetArrayItem(21, item);
					item.uCharVal = Convert.ToInt16(tmp.key_s);		// 22 - key_s
					arr.ksSetArrayItem(22, item);
					item.uCharVal = tmp.key_s_on;		// 23 - key_s_on
					arr.ksSetArrayItem(23, item);
					item.uCharVal = tmp.key_s_gray;		// 24 - key_s_gray
					arr.ksSetArrayItem(24, item);
					item.uCharVal = tmp.koef_mat_on;	// 25 - koef_mat_on
					arr.ksSetArrayItem(25, item);
					item.shortVal = tmp.ver;			// 26 - ver
					arr.ksSetArrayItem(26, item);
				}
			}
		}


		/// <summary>
		/// Получить параметры
		/// </summary>
		public void GetUserParam()
		{
			if (Param != null)
			{
				ksLtVariant item = (ksLtVariant)Kompas.Instance.KompasObject.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
				ksDynamicArray arr = (ksDynamicArray)Param.GetUserArray();
				if (item != null && arr != null)
				{
					int count = arr.ksGetArrayCount();
					if (count >= 27)
					{
						item.Init();
						arr.ksGetArrayItem(0, item);
						par.ang = item.floatVal;
						arr.ksGetArrayItem(1, item);
						par.flagAttr = item.shortVal;
						arr.ksGetArrayItem(2, item);
						par.drawType = item.shortVal;
						arr.ksGetArrayItem(3, item);
						par.typeSwitch = Convert.ToByte(item.shortVal);
					
						arr.ksGetArrayItem(4, item);
						tmp.dr = item.floatVal;
						arr.ksGetArrayItem(5, item);
						tmp.s = item.floatVal;
						arr.ksGetArrayItem(6, item);
						tmp.d = item.floatVal;
						arr.ksGetArrayItem(7, item);
						tmp.da = item.floatVal;
						arr.ksGetArrayItem(8, item);
						tmp.H = item.floatVal;
						arr.ksGetArrayItem(9, item);
						tmp.d2 = item.floatVal;
						arr.ksGetArrayItem(10, item);
						tmp.p = item.floatVal;
						arr.ksGetArrayItem(11, item);
						tmp.cls = item.shortVal;
						arr.ksGetArrayItem(12, item);
						tmp.gost = item.shortVal;
						arr.ksGetArrayItem(13, item);
						tmp.hatchAng = item.floatVal;
						arr.ksGetArrayItem(14, item);
						tmp.hatchStep = item.floatVal;
						arr.ksGetArrayItem(15, item);
						tmp.massa = item.floatVal;
						arr.ksGetArrayItem(16, item);
						tmp.indexMassa = item.shortVal;
						arr.ksGetArrayItem(17, item);
						tmp.perform = Convert.ToByte(item.uCharVal);	// 17 - perform
						arr.ksGetArrayItem(18, item);
						tmp.simple = Convert.ToBoolean(item.uCharVal);		// 18 - simple
						arr.ksGetArrayItem(19, item);
						tmp.axis_off = Convert.ToByte(item.uCharVal);	// 19 - axis_off
						arr.ksGetArrayItem(20, item);
						tmp.pitch = Convert.ToBoolean(item.uCharVal);		// 20 - pitch
						arr.ksGetArrayItem(21, item);
						tmp.pitch_off = Convert.ToByte(item.uCharVal);	// 21 - pitch_off
						arr.ksGetArrayItem(22, item);
						tmp.key_s = Convert.ToBoolean(item.uCharVal);		// 22 - key_s
						arr.ksGetArrayItem(23, item);
						tmp.key_s_on = Convert.ToByte(item.uCharVal);	// 23 - key_s_on
						arr.ksGetArrayItem(24, item);
						tmp.key_s_gray = Convert.ToByte(item.uCharVal);	// 24 - key_s_gray
						arr.ksGetArrayItem(25, item);
						tmp.koef_mat_on = Convert.ToByte(item.uCharVal);// 25 - koef_mat_on
						arr.ksGetArrayItem(26, item);
						tmp.ver = item.shortVal;						// 26 - ver
					}
				}
			}
		}
						
			
		/// <summary>
		/// Сохранить параметры в модели
		/// </summary>
		public void SetParam()
		{
			SetUserParam();
			Kompas.Instance.Document2D.ksSetMacroParam(refMacr, Param, false, false, false);
		}
	

		/// <summary>
		/// Функция обратной связи, вызываемая из Placement
		/// </summary>
		/// <param name="comm"></param>
		/// <param name="X"></param>
		/// <param name="Y"></param>
		/// <param name="ang"></param>
		/// <returns></returns>
		public int CALLBACKPROCPLACEMENT(short comm, ref double X, ref double Y,
			ref double ang, object info_, object phan_, short dynamic)
		{
			ksPhantom phan = (ksPhantom)phan_;
			ksRequestInfo info = (ksRequestInfo)info_;
			string ChoiceMenu = string.Empty;
			ksType1 t1 = (ksType1)phan.GetPhantomParam();
			int gr = 0;
			if (dynamic == 0)	// фиксация
			{
				switch (comm)
				{
					case -1:	// поставить в модель
						par.ang = Convert.ToSingle(ang);
						SetParam();
						Kompas.Instance.Document2D.ksSetMacroPlacement(refMacr, X, Y, ang, 0);
						Kompas.Instance.Document2D.ksStoreTmpGroup(t1.gr);

						if (DrawSpcObj(t1.gr))
						{
							Kompas.Instance.Document2D.ksClearGroup(t1.gr, true);
							return 0;
						}

						Kompas.Instance.Document2D.ksClearGroup(t1.gr, true);
						if (flagMode > 0)
						{
							return 0;
						}
						break;
				}

				info.commandsString = ChoiceMenu;
				GetGroup(ref gr);
				t1.gr = gr;
			}
			else
			{
				par.ang = Convert.ToSingle(ang);
			}
			return 1;
		}


		/// <summary>
		/// Процесс вставки гайки в модель
		/// </summary>
		public void Draw()
		{
			if (phantom == null)
				phantom = (ksPhantom)Kompas.Instance.KompasObject.GetParamStruct((short)StructType2DEnum.ko_Phantom);
			par.typeSwitch = 0;
			object flagCalcPaket;
			int j = 0;
			double X = 0;
			double Y = 0;
			double ang = 0;
			int gr = 0;
			short count = 0;

			if (phantom != null)
			{
				phantom.Init();
				phantom.phantom = 1;
				ksType1 t1 = (ksType1)phantom.GetPhantomParam();
				if (t1 != null)
				{
					t1.Init();
					t1.scale_ = 1;
					t1.gr = 0;		// временная группа
					j = 1;
					if (Kompas.Instance.KompasObject.ksReturnResult() == 0)
					{
						flagMode = Kompas.Instance.Document2D.ksEditMacroMode();
					
						if (MacroElementParam())
						{
							do
							{
								flagSwitch = false;
								flagCalcPaket = false;
								GetGroup(ref gr);
								t1.gr = gr;
								t1.angle = par.ang;
							
								ksRequestInfo info = (ksRequestInfo)Kompas.Instance.KompasObject.GetParamStruct((short)StructType2DEnum.ko_RequestInfo);
								if (info != null)
								{
									info.Init();
									info.dynamic = 1;
								
									// указываем адрес обратной функции для Placement
									info.SetCallBackP("CALLBACKPROCPLACEMENT", 0, this);
									j = Kompas.Instance.Document2D.ksPlacementEx(info, ref X, ref Y, ref ang, phantom, processParam);
								
									if (spcObj[0] > 0)
									{
										ksSpecification spc = (ksSpecification)Kompas.Instance.Document2D.GetSpecification();
										count = countObj;
										for (int i = 0; i < count; i ++)
										{
											// Олегово окно - редактируем параметры
											if (spc != null && spc.ksEditWindowSpcObject(spcObj[i]) != 0)
												DrawPosLeader(spcObj[i], spc);
										}
									
										spcObj[0] = 0;
									
										if (flagMode == 0)
										{
											flagSwitch = true;
											if (par.typeSwitch != 0)
												par.typeSwitch = 0;
											else
												par.typeSwitch = 1;
										}
									}
								}
								
								if (flagSwitch == true)
								{
									j = 1;
									if (par.typeSwitch > 0)
										par.typeSwitch = 0;
									else
										par.typeSwitch = 1;
								}
							}
							while (j != 0);
							CloseGaykaBase();	// закрытие базы
						}
					}
				}
			}
		}

	
		/// <summary>
		/// Создание МАКРО гайки
		/// </summary>
		public void GetGroup(ref int gr)
		{
			short k2;
			if (tmp.perform == 1)
				k2 = 2;
			else
				k2 = 1;
		
			if (Kompas.Instance.Document2D.ksExistObj(gr) != 0)
			{
				Kompas.Instance.Document2D.ksDeleteObj(gr);
			}
		
			gr = Kompas.Instance.Document2D.ksNewGroup(1);
			Kompas.Instance.Document2D.ksMacro(0);
			switch (par.drawType)
			{
				case ID_VID:	// вид
					if (!tmp.simple)
					{
						gayka_k(0, 0, 0, tmp.s, tmp.d, 0, tmp.H, 1, 0, tmp.d2, k2);
						gayka_k(0, 0, 0, tmp.s, tmp.d, 0, tmp.H, -1, 0, tmp.d2, k2);
					}
					else
					{
						gayka_k_y(0, 1);
						gayka_k_y(0, -1);
					}
				
					if (tmp.axis_off == 0)
						Kompas.Instance.Document2D.ksLineSeg(-3, 0, tmp.H + 3, 0, 3);
					break;
				case ID_SIDEVID: // вид сбоку
					if (tmp.axis_off == 0)
						Kompas.Instance.Document2D.ksLineSeg(-3, 0, tmp.H + 3, 0, 3);
				
					if (tmp.simple)
						k2 = 3;
				
					gayka_k_side(0, tmp.s, tmp.d, tmp.d2, tmp.H, 1, k2);
					gayka_k_side(0, tmp.s, tmp.d, tmp.d2, tmp.H, -1, k2);
					break;
				case ID_TOPVID:
					gayka_sverhu();	// вид ссверху
					break;
				case ID_VIDSEC:		// пол-вида/пол-разреза
					if (!tmp.simple)
						gayka_k(0, 0, 0, tmp.s, tmp.d, 0, tmp.H, 1, 0, tmp.d2, k2);
					else
						gayka_k_y(0, 1);
					gayka_p_k(-1);
				
					if (tmp.axis_off == 0)
						Kompas.Instance.Document2D.ksLineSeg(-3, 0, tmp.H + 3, 0, 3);
					break;
			}
			refMacr = Kompas.Instance.Document2D.ksEndObj();
			Kompas.Instance.Document2D.ksEndGroup();
		}
	

		/// <summary>
		/// Создание и инициализация параметров гайки на панели свойств
		/// </summary>
		/// <returns></returns>
		public bool MacroElementParam()
		{
			bool result = true;
		
			PropertyTabs tabs;
			PropertyTab tab;
			PropertyControls ctrls;
			ksLtVariant item;
			ksDynamicArray arr;
			int i;
			PropertySeparator sep;
			string fullPath = string.Empty;
			if (Kompas.Instance.KompasApp != null)
				OpenGaykaBase();
			
			processParam = Kompas.Instance.KompasApp.CreateProcessParam();
			processParam.SpecToolbar = SpecPropertyToolBarEnum.pnEnterEscCreateHelp;
			processParam.ChangeControlValue += new ksPropertyManagerNotify_ChangeControlValueEventHandler(processParam_ChangeControlValue);
			processParam.ControlCommand += new ksPropertyManagerNotify_ControlCommandEventHandler(processParam_ControlCommand);
			tabs = processParam.PropertyTabs;
			tab = tabs.Add("Параметры гайки");
			ctrls = tab.PropertyControls;
			
			// Диаметр резьбы
			diamEdit = (PropertyList)ctrls.Add(ControlTypeEnum.ksControlListReal);
			diamEdit.Name = "&Диаметр";
			diamEdit.Id = DIAM_ID;
			diamEdit.NameVisibility = PropertyControlNameVisibility.ksNameAlwaysVisible;
			diamEdit.Sort = false;
			diamEdit.Width = 7;
			diamEdit.ReadOnly = true;
			diamEdit.Hint = "Диаметр резьбы";
			diamEdit.Tips = "Диаметр резьбы";
			
			// заполним список диаметров
			item = (ksLtVariant)Kompas.Instance.KompasObject.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
			arr = (ksDynamicArray)paramTmp.GetUserArray();

			if (item != null && arr != null)
			{
				i = 1;
				do
				{
					i = data.ksReadRecord(base_.bg, base_.rg, paramTmp);
					
					if (i > 0)
					{
						arr.ksGetArrayItem(0, item);
						diamEdit.Add((item.floatVal));
					}
				}
				while (i != 0);
			}
			
			diamEdit.Value = tmp.dr;	// выделим текущий диаметр гайки
			
			// Разделитель
			sep = (PropertySeparator)ctrls.Add(ControlTypeEnum.ksControlSeparator);
			sep.Name = "Отображение";
			sep.SeparatorType = SeparatorTypeEnum.ksSeparatorDownName;

			viewButt = (PropertyMultiButton)ctrls.Add(ControlTypeEnum.ksControlMultiButton);
			viewButt.Id = VIEWSIDE_ID;	// Кнопки отображения
			viewButt.Name = "&Вид отображения";
			viewButt.Tips = "Вид отображения";
			viewButt.Hint = "Вид отображения гайки";
			viewButt.ButtonsType = ButtonTypeEnum.ksRadioButton;
			viewButt.NameVisibility = PropertyControlNameVisibility.ksNameVerticalVisible;

			viewButt.ResModule = Assembly.GetExecutingAssembly().Location;
			
			if (GetFullName("G_view.bmp", ref fullPath))
			{
				viewButt.AddButton(BASE_VIEW, fullPath, -1);	// Главный вид
			}
			else
			{
				viewButt.AddButton(BASE_VIEW, BASE_VIEW, -1);	// Главный вид
			}
			
			if (GetFullName("G_left.bmp", ref fullPath))
			{
				viewButt.AddButton(LEFT_VIEW, fullPath, -1);	// Вид слева
			}
			else
			{
				viewButt.AddButton(LEFT_VIEW, LEFT_VIEW, -1);	// Вид слева
			}
			
			if (GetFullName("G_top.bmp", ref fullPath))
			{
				viewButt.AddButton(TOP_VIEW, fullPath, -1);	// Вид сверху
			}
			else
			{
				viewButt.AddButton(TOP_VIEW, TOP_VIEW, -1);	// Вид сверху
			}
			
			if (GetFullName("G_sec.bmp", ref fullPath))
			{
				viewButt.AddButton(SEC_VIEW, fullPath, -1);	// Вид\разрез
			}
			else
			{
				viewButt.AddButton(SEC_VIEW, SEC_VIEW, -1);	// Вид\разрез
			}
			
			switch (par.drawType)	// выделим текущий вид
			{
				case ID_VID: viewButt.set_ButtonChecked(BASE_VIEW, true);		break;	// главный вид
				case ID_SIDEVID: viewButt.set_ButtonChecked(LEFT_VIEW, true);	break;	// вид сбоку
				case ID_TOPVID: viewButt.set_ButtonChecked(TOP_VIEW, true);		break;	// вид сверху
				case ID_VIDSEC: viewButt.set_ButtonChecked(SEC_VIEW, true);		break;	// главный вид \ разрез
			}

			viewButt.set_ButtonTips(BASE_VIEW, "Главный вид");
			viewButt.set_ButtonHint(BASE_VIEW, "Главный вид отображения Гайки");
			viewButt.set_ButtonTips(LEFT_VIEW, "Вид сбоку");
			viewButt.set_ButtonTips(TOP_VIEW, "Вид сверху");
			viewButt.set_ButtonTips(SEC_VIEW, @"Вид\разрез");
			
			// Кнопки исполнений
			varButt = (PropertyMultiButton)ctrls.Add(ControlTypeEnum.ksControlMultiButton);
			varButt.Id = PERFORMANCE_ID;
			varButt.Name = "&Исполнение";
			varButt.ButtonsType = ButtonTypeEnum.ksRadioButton;
			varButt.NameVisibility = PropertyControlNameVisibility.ksNameVerticalVisible;
			varButt.ResModule = Assembly.GetExecutingAssembly().Location;
			varButt.Hint = "Выбор исполнения";
			varButt.Tips = "Исполнение";
			
			if (GetFullName("G_i1.bmp", ref fullPath))
				varButt.AddButton(PERF1_VIEW, fullPath, -1);	// Исполнение 1
			else
				varButt.AddButton(PERF1_VIEW, PERF1_VIEW, -1);	// Исполнение 1
			
			if (GetFullName("G_i2.bmp", ref fullPath))
				varButt.AddButton(PERF2_VIEW, fullPath, -1);	// Исполнение 2
			else
				varButt.AddButton(PERF2_VIEW, PERF2_VIEW, -1);	// Исполнение 2
			
			if (tmp.perform != 0)	// выделим исполнение
				varButt.set_ButtonChecked(PERF2_VIEW, true);	// исполнение 2
			else
				varButt.set_ButtonChecked(PERF1_VIEW, true);	// исполнение 1

			varButt.set_ButtonTips(PERF1_VIEW, "Исполнение 1");
			varButt.set_ButtonTips(PERF2_VIEW, "Исполнение 2");
			
			// Кнопки упрощений
			simpButt = (PropertyMultiButton)ctrls.Add(ControlTypeEnum.ksControlMultiButton);
			simpButt.Id = SIMPLES_ID;
			simpButt.Name = "&Упрощения";
			simpButt.ButtonsType = ButtonTypeEnum.ksCheckButton;
			simpButt.NameVisibility = PropertyControlNameVisibility.ksNameVerticalVisible;
			simpButt.ResModule = Assembly.GetExecutingAssembly().Location;
			simpButt.Hint = "Упрощения отображения гайки";
			simpButt.Tips = "Упрощения отображения";
			
			if (GetFullName("G_simple.bmp", ref fullPath))
				simpButt.AddButton(SIMPLE_VIEW, fullPath, -1);		// Упрошено
			else
				simpButt.AddButton(SIMPLE_VIEW, SIMPLE_VIEW, -1);	// Упрошено
			
			if (GetFullName("G_osx_on.bmp", ref fullPath))
				simpButt.AddButton(DRAW_AXIS, fullPath, -1);		// Рисовать ось
			else
				simpButt.AddButton(DRAW_AXIS, DRAW_AXIS, -1);		// Рисовать ось

			if (tmp.simple)
				simpButt.set_ButtonChecked(SIMPLE_VIEW, true);		// Упрошено

			if (tmp.axis_off == 0)
				simpButt.set_ButtonChecked(DRAW_AXIS, true);		// Рисовать ось

			simpButt.set_ButtonTips(SIMPLE_VIEW, "Упрошено");
			simpButt.set_ButtonTips(DRAW_AXIS, "Рисовать ось");
			
			// Разделитель
			sep = (PropertySeparator)ctrls.Add(ControlTypeEnum.ksControlSeparator);
			sep.SeparatorType = SeparatorTypeEnum.ksSeparatorDownName;
			
			// Кнопки доп. параметров
			addButt = (PropertyMultiButton)ctrls.Add(ControlTypeEnum.ksControlMultiButton);
			addButt.Id = ADD_PARAM_ID;
			addButt.Name = "Дополнительные па&раметры";
			addButt.ButtonsType = ButtonTypeEnum.ksCheckButton;
			addButt.NameVisibility = PropertyControlNameVisibility.ksNameVerticalVisible;
			addButt.ResModule = Assembly.GetExecutingAssembly().Location;
			addButt.Hint = "Дополнительные параметры";
			addButt.Tips = "Дополнительные параметры";
			
			if (GetFullName("G_step.bmp", ref fullPath))
				addButt.AddButton(ADDSTEP_BUTT, fullPath, -1);	// Мелкий шаг
			else
				addButt.AddButton(ADDSTEP_BUTT, ADDSTEP_BUTT, -1);// Мелкий шаг
			
			if (GetFullName("G_key.bmp", ref fullPath))
				addButt.AddButton(KEY_BUTT, fullPath, -1);	// Доп. размер под ключ
			else
				addButt.AddButton(KEY_BUTT, KEY_BUTT, -1);	// Доп. размер под ключ

			if (tmp.pitch)
				addButt.set_ButtonChecked(ADDSTEP_BUTT, true);

			if (tmp.key_s)
				addButt.set_ButtonChecked(KEY_BUTT, true);

			addButt.set_ButtonTips(ADDSTEP_BUTT, "Мелкий шаг");
			addButt.set_ButtonTips(KEY_BUTT, "Дополнительный размер под ключ");

			// Чекбокс объекта сп
			spcCheck = (PropertyCheckBox)ctrls.Add(ControlTypeEnum.ksControlCheckBox);
			spcCheck.Id = SPC_CHECK_ID;
			spcCheck.Name = "&Создать объект спецификации";
			spcCheck.Value = par.flagAttr;
			spcCheck.Hint = "Создать объект спецификации";
			spcCheck.Tips = "Создать объект спецификации";

			userCtrl = (PropertyUserControl)ctrls.Add(ControlTypeEnum.ksControlUser);
			userCtrl.SetOCXControl("Steps.NET.HatchControl");
			userCtrl.Id = 10011;
			userCtrl.Name = "&Тест OCX";
			userCtrl.Width = 200;
			userCtrl.Height = 200;
			userCtrl.CreateOCX += new ksPropertyUserControlNotify_CreateOCXEventHandler(userCtrl_CreateOCX);
			userCtrl.DestroyOCX += new ksPropertyUserControlNotify_DestroyOCXEventHandler(userCtrl_DestroyOCX);
			
//			// Разделитель
//			sep = (PropertySeparator)ctrls.Add(ControlTypeEnum.ksControlSeparator);
//			sep.Name = "Параметры штриховки";
//			sep.SeparatorType = SeparatorTypeEnum.ksSeparatorDownName;
//			
//			// угол штриховки
//			angleEdit = (PropertyEdit)ctrls.Add(ControlTypeEnum.ksControlEditReal);
//			angleEdit.Id = ANGLE_HATCH_ID;
//			angleEdit.Name = "У&гол, гр";
//			angleEdit.NameVisibility = PropertyControlNameVisibility.ksNameAlwaysVisible;
//			angleEdit.Width = 7;
//			angleEdit.Value = tmp.hatchAng;
//			angleEdit.Enable = true;
//			//par.drawType = ID_VIDSEC;
//			angleEdit.Hint = "Угол штриховки";
//			angleEdit.Tips = "Угол";
//			
//			// шаг штриховки
//			stepEdit = (PropertyEdit)ctrls.Add(ControlTypeEnum.ksControlEditReal);
//			stepEdit.Id = STEP_HATCH_ID;
//			stepEdit.Name = "&Шаг, мм";
//			stepEdit.NameVisibility = PropertyControlNameVisibility.ksNameAlwaysVisible;
//			stepEdit.Width = 7;
//			stepEdit.Value = tmp.hatchStep;
//			stepEdit.Enable = true;
//			//par.drawType = ID_VIDSEC;
//			stepEdit.Hint = "Шаг штриховки";
//			stepEdit.Tips = "Шаг";
			
			//Const PARAMS_ID = 10009;     // список параметров
			//Const VIEW_BOX_ID = 10010;   // окно просмотра

			//Set hatchPar = CreateObject("HatchControl1.HatchControl");
			//hatchPar.AngleEditValue = "10";
			
			return result;
		}
		

		/// <summary>
		/// Отрисовка гайки
		/// </summary>
		/// <param name="ls"></param>
		/// <param name="l"></param>
		/// <param name="d1"></param>
		/// <param name="s"></param>
		/// <param name="d"></param>
		/// <param name="l1"></param>
		/// <param name="H"></param>
		/// <param name="j"></param>
		/// <param name="j1"></param>
		
		public void gayka_k(double ls, double l, double d1, double s, double d, double l1,
			double H, short j, short j1, double d2, short j2)
		{
			double[] X = new double[9];
			double[] Y = new double[9];
			double c;
			double h1;
			double rb;
			double xc2;
			double yc2;
			double xcbl;
			double ycbl;
			double xcbp;
			double ycbp;
			double ycml;
			if (Kompas.Instance.Math != null)
			{
				// Координаты контура рез_части
				// j1=1 - контровочное отверстие j1=0 - контров. отв. нет
				// j2=1 исполнение 1 j2=2 исполнение 2

				d = s / Kompas.Instance.Math.ksCosD(30);
				c = (d - d2) / 2 * Kompas.Instance.Math.ksTanD(30);
				h1 = d * 0.5 * Kompas.Instance.Math.ksSinD(30);
				rb = (h1 * h1 + c * c) / 2 / c;

				X[1] = ls;
				Y[1] = 0;
				if (j2 == 1)
				{
					X[2] = ls;
					Y[2] = j * (d2 * 0.5);
					X[3] = ls + c;
					Y[3] = j * (d * 0.5);
					X[7] = ls + c;
					Y[7] = j * h1;
				}
				else
				{
					X[2] = ls;
					Y[2] = j * (d * 0.5);
					X[7] = ls;
					Y[7] = j * h1;
				}
			
				X[4] = ls + H - c;
				Y[4] = j * (d * 0.5);
				X[5] = ls + H;
				Y[5] = j * (d2 * 0.5);
				X[6] = ls + H;
				Y[6] = 0;
			
				X[8] = ls + H - c;
				Y[8] = j * h1;
				xc2 = ls + l;
				yc2 = j * (d * 0.5 - l1);
			
				xcbl = ls + rb;
				ycbl = 0;
				xcbp = ls + H - rb;
				ycbp = 0;
				ycml = j * ((d * 0.5 - h1) / 2 + h1);
			
				if (j2 == 1)
				{
					Kompas.Instance.Document2D.ksLineSeg(X[1], Y[1], X[2], Y[2], 1);
					Kompas.Instance.Document2D.ksLineSeg(X[2], Y[2], X[3], Y[3], 1);
					Kompas.Instance.Document2D.ksLineSeg(X[3], Y[3], X[4], Y[4], 1);
					Kompas.Instance.Document2D.ksLineSeg(X[4], Y[4], X[5], Y[5], 1);
					Kompas.Instance.Document2D.ksLineSeg(X[5], Y[5], X[6], Y[6], 1);
					Kompas.Instance.Document2D.ksLineSeg(X[7], Y[7], X[8], Y[8], 1);
					Kompas.Instance.Document2D.ksArcByPoint(xcbl, ycbl, rb, X[1], Y[1], X[7], Y[7], Convert.ToInt16(-j), 1);
					Kompas.Instance.Document2D.ksArcByPoint(xcbp, ycbp, rb, X[6], Y[6], X[8], Y[8], j, 1);
					Kompas.Instance.Document2D.ksArcBy3Points(ls + c * 0.5, (d * 0.5 - (d - d2) / 4) * j, ls, ycml, X[7], Y[7], 1);
					Kompas.Instance.Document2D.ksArcBy3Points(ls + H - c * 0.5, (d * 0.5 - (d - d2) / 4) * j, ls + H, ycml, X[8], Y[8], 1);
				}
				else
				{
					Kompas.Instance.Document2D.ksLineSeg(X[1], Y[1], X[2], Y[2], 1);
					Kompas.Instance.Document2D.ksLineSeg(X[2], Y[2], X[4], Y[4], 1);
					Kompas.Instance.Document2D.ksLineSeg(X[4], Y[4], X[5], Y[5], 1);
					Kompas.Instance.Document2D.ksLineSeg(X[5], Y[5], X[6], Y[6], 1);
					Kompas.Instance.Document2D.ksLineSeg(X[7], Y[7], X[8], Y[8], 1);
					Kompas.Instance.Document2D.ksArcByPoint(xcbp, ycbp, rb, X[6], Y[6], X[8], Y[8], j, 1);
					Kompas.Instance.Document2D.ksArcBy3Points(ls + H - c * 0.5, (d * 0.5 - (d - d2) / 4) * j, ls + H, ycml, X[8], Y[8], 1);
				}			
				if (j1 == 1)
				{
					Kompas.Instance.Document2D.ksCircle(xc2, yc2, d1 * 0.5, 1);
					Kompas.Instance.Document2D.ksLineSeg(xc2 - 2, yc2, xc2 + 2, yc2, 2);
					Kompas.Instance.Document2D.ksLineSeg(xc2, yc2 - 2, xc2, yc2 + 2, 2);
				}
			}
		}
	

		/// <summary>
		/// Отрисовка гайки
		/// </summary>
		/// <param name="ls"></param>
		/// <returns></returns>
		public void gayka_k_side(double ls, double s, double d, double d2, double H, short j, short j2)
		{
			double X;
			double Y;
			double x2;
			double y2;
			short c;

			if (Kompas.Instance.Math != null)
			{
				// j2 = 1 исполнение 1 j2 = 2 упрощенное
				c = Convert.ToInt16((d - d2) / 2 * Kompas.Instance.Math.ksTanD(30));
				Y = j * s * 0.5;
				if (j2 == 1)
				{
					X = ls + c;
					Kompas.Instance.Document2D.ksLineSeg(ls, 0, ls, j * d2 * 0.5, 1);
					Kompas.Instance.Document2D.ksLineSeg(ls, j * d2 * 0.5, X, Y, 1);
					Kompas.Instance.Document2D.ksArcBy3Points(ls + c, j * (s * 0.5), ls, s * 0.25 * j, ls + c, 0, 1);
				}
				else
				{
					X = ls;
					Kompas.Instance.Document2D.ksLineSeg(X, 0, X, Y, 1);
				}
			
				if (j2 == 3)
				{
					x2 = ls + H;
					y2 = Y;
				}
				else
				{
					x2 = ls + H - c;
					y2 = j * d2 * 0.5;
				}
			
				Kompas.Instance.Document2D.ksLineSeg(X, Y, x2, Y, 1);
				if (j2 != 3)
				{
					Kompas.Instance.Document2D.ksLineSeg(x2, Y, ls + H, y2, 1);
					Kompas.Instance.Document2D.ksArcBy3Points(ls + H - c, j * (s * 0.5), ls + H, s * 0.25 * j, ls + H - c, 0, 1);
				}
				Kompas.Instance.Document2D.ksLineSeg(ls + H, y2, ls + H, 0, 1);
				if (j > 0)
					Kompas.Instance.Document2D.ksLineSeg(X, 0, x2, 0, 1);
			}
		}
	

		/// <summary>
		/// Отрисовка гайки
		/// </summary>
		/// <param name="ls"></param>
		/// <returns></returns>
		public void gayka_k_y(double ls, short j)
		{
			double s = 0;
			double h1;
			double p1;
			double p2;
			if (Kompas.Instance.Math != null)
			{
				tmp.d = Convert.ToSingle(tmp.s / Kompas.Instance.Math.ksCosD(30));

				h1 = tmp.d * 0.5 * Kompas.Instance.Math.ksSinD(30);
				p1 = j * tmp.d * 0.5;
				p2 = Convert.ToSingle(s + tmp.H);
				Kompas.Instance.Document2D.ksLineSeg(ls, 0, ls, p1, 1);
				Kompas.Instance.Document2D.ksLineSeg(ls, p1, p2, p1, 1);
				Kompas.Instance.Document2D.ksLineSeg(p2, p1, p2, 0, 1);
				Kompas.Instance.Document2D.ksLineSeg(ls, j * h1, p2, j * h1, 1);
			}
		}
	

		/// <summary>
		/// Отрисовка гайки
		/// </summary>
		/// <param name="j"></param>
		/// <returns></returns>
		public void gayka_p_k(short j)
		{
			double c1 = 0;
			double c2 = 0;
			double X = 0;
			double x1 = 0;
			double x2 = 0;
			double x3 = 0;
			double y2 = 0;
			double y3 = 0;
			double c = 0;
			double Y = 0;
			double dd = 0;
			if (Kompas.Instance.Math != null)
			{
				c1 = 0;
				c = (tmp.d - tmp.d2) / 2 * Kompas.Instance.Math.ksTanD(30);
				if (tmp.perform == 0 && !tmp.simple)
					Y = tmp.d2 * 0.5 * j;
				else
					Y = tmp.d * 0.5 * j;

				Kompas.Instance.Document2D.ksLineSeg(0, 0, 0, Y, 1);

				if (tmp.perform == 0 && !tmp.simple)
				{
					X = c;
					Kompas.Instance.Document2D.ksLineSeg(0, Y, X, j * (tmp.d * 0.5), 1);
				}

				dd = tmp.dr - 2 * MODSTEP_REAL * tmp.p;

				if (!tmp.simple)
				{
					x3 = tmp.H - c;
					y3 = j * tmp.d2 * 0.5;
					y2 = j * tmp.da * 0.5;
				}
				else
				{
					x3 = tmp.H;
					y3 = j * tmp.d * 0.5;
					y2 = j * dd * 0.5;
				}

				Kompas.Instance.Document2D.ksLineSeg(X, j * tmp.d * 0.5, x3, j * tmp.d * 0.5, 1);
				if (!tmp.simple)
					Kompas.Instance.Document2D.ksLineSeg(x3, j * tmp.d * 0.5, tmp.H, y3, 1);

				Kompas.Instance.Document2D.ksLineSeg(tmp.H, y3, tmp.H, 0, 1);

				x1 = tmp.H;
				x2 = x1;
				if (!tmp.simple)
				{
					c1 = (tmp.da - dd) * 0.5;
					c2 = (tmp.da - tmp.dr) * 0.5;
					if (tmp.perform == 0)
						x2 = x2 - c2;
				}

				if (tmp.perform == 0 && !tmp.simple)
				{
					x1 = x1 - c1;
					Kompas.Instance.Document2D.ksLineSeg(tmp.H, j * tmp.da * 0.5, x1, j * 0.5 * dd, 1);
					Kompas.Instance.Document2D.ksLineSeg(x1, j * dd * 0.5, x1, 0, 1);
				}

				Kompas.Instance.Document2D.ksLineSeg(x1, j * 0.5 * dd, c1, j * 0.5 * dd, 1);

				if (!tmp.simple)
				{
					Kompas.Instance.Document2D.ksLineSeg(c1, j * 0.5 * dd, 0, j * 0.5 * tmp.da, 1);
					Kompas.Instance.Document2D.ksLineSeg(c1, j * dd * 0.5, c1, 0, 1);
				}

				Kompas.Instance.Document2D.ksHatch(0, tmp.hatchAng, tmp.hatchStep, 0, 0, 0);
				Kompas.Instance.Document2D.ksLineSeg(0, y2, 0, Y, 1);
				if (tmp.perform == 0 && !tmp.simple)
				{
					Kompas.Instance.Document2D.ksLineSeg(0, Y, X, j * (tmp.d * 0.5), 1);
					Kompas.Instance.Document2D.ksLineSeg(tmp.H, y3, tmp.H, j * tmp.da * 0.5, 1);
					Kompas.Instance.Document2D.ksLineSeg(tmp.H, j * tmp.da * 0.5, x1, j * 0.5 * dd, 1);
				}
				else
					Kompas.Instance.Document2D.ksLineSeg(tmp.H, y3, tmp.H, j * dd * 0.5, 1);

				Kompas.Instance.Document2D.ksLineSeg(X, j * (tmp.d * 0.5), x3, j * tmp.d * 0.5, 1);
				if (!tmp.simple)
				{
					Kompas.Instance.Document2D.ksLineSeg(x3, j * tmp.d * 0.5, tmp.H, y3, 1);
					Kompas.Instance.Document2D.ksLineSeg(c1, j * 0.5 * dd, 0, j * 0.5 * tmp.da, 1);
				}

				Kompas.Instance.Document2D.ksLineSeg(x1, j * 0.5 * dd, c1, j * 0.5 * dd, 1);

				Kompas.Instance.Document2D.ksEndObj();

				Kompas.Instance.Document2D.ksLineSeg(c2, j * 0.5 * tmp.dr, x2, j * 0.5 * tmp.dr, 2);
			}
		}


		/// <summary>
		/// отрисовка гайки
		/// </summary>
		/// <returns></returns>
		public void gayka_sverhu()
		{
			double h1;
			double dd;
			double d;
			double s;
			double rad;
			double x1;
			double y1;
			if (Kompas.Instance.Math != null)
			{
				s = tmp.s * 0.5;
				d = s / Kompas.Instance.Math.ksCosD(30);
					dd = tmp.dr - 2 * MODSTEP_REAL * tmp.p;
				h1 = d * Kompas.Instance.Math.ksSinD(30);
			
				Kompas.Instance.Document2D.ksLineSeg(-s, h1, 0, d, 1);
				Kompas.Instance.Document2D.ksLineSeg(0, d, s, h1, 1);
				Kompas.Instance.Document2D.ksLineSeg(s, h1, s, -h1, 1);
				Kompas.Instance.Document2D.ksLineSeg(s, -h1, 0, -d, 1);
				Kompas.Instance.Document2D.ksLineSeg(0, -d, -s, -h1, 1);
				Kompas.Instance.Document2D.ksLineSeg(-s, -h1, -s, h1, 1);
			
				if (!tmp.simple)
					Kompas.Instance.Document2D.ksCircle(0, 0, tmp.d2 * 0.5, 1);
			
				Kompas.Instance.Document2D.ksCircle(0, 0, dd * 0.5, 1);
				rad = tmp.dr * 0.5;
				x1 = rad * Kompas.Instance.Math.ksSinD(15);
				y1 = rad * Kompas.Instance.Math.ksCosD(15);
				Kompas.Instance.Document2D.ksArcByPoint(0, 0, rad, x1, y1, y1, -x1, 1, 2);
			
				if (tmp.axis_off == 0)
				{
					if (d >= 6)
					{
						Kompas.Instance.Document2D.ksLineSeg(-3 - s, 0, s + 3, 0, 3);
						Kompas.Instance.Document2D.ksLineSeg(0, -3 - d, 0, 3 + d, 3);
					}
					else
					{
						Kompas.Instance.Document2D.ksLineSeg(-1 - s, 0, s + 1, 0, 3);
						Kompas.Instance.Document2D.ksLineSeg(0, -1 - d, 0, 1 + d, 3);
					}
				}
			
			}
		}
	

		/// <summary>
		/// отрисовать позиционную линию выноски
		/// уже проверено , что объект спецификации есть
		/// Функцию нужно запускать вне Cursor и Placement
		/// </summary>
		/// <param name="spcObj"></param>
		/// <param name="spc"></param>
		public void DrawPosLeader(int spcObj, ksSpecification spc)
		{
			ksRequestInfo info = (ksRequestInfo)Kompas.Instance.KompasObject.GetParamStruct((short)StructType2DEnum.ko_RequestInfo);

			int posLeater = 0;
			int posLeader = 0;
			bool flag = false;
			double x1 = 0;
			double y1 = 0;
			int j1 = 0;
			string menu = string.Empty;

			// Создать новую линию выноски
			if (info != null)
			{
				info.Init();
				flag = false;
				posLeader = 0;
				do
				{
					info.commandsString = "!Подключить_существующую !Создать_новую_линию_выноски";
					info.prompt = "Укажите линию выноски";
					j1 = Kompas.Instance.Document2D.ksCursor(info, ref x1, ref y1, null);
				
					switch (j1)
					{
						case 2:
							posLeader = Kompas.Instance.Document2D.ksCreateViewObject(ldefin2d.POSLEADER_OBJ);
							flag = false;
							break;
						case 1:	// Подключить существующую
							info.commandsString = menu;
							if (Kompas.Instance.Document2D.ksCursor(info, ref x1, ref y1, null) != 0)
							{
								posLeader = Kompas.Instance.Document2D.ksFindObj(x1, y1, 100);	// величина стороны окошка-ловушки с центром x,y
								if (posLeader == 0 && Kompas.Instance.Document2D.ksGetObjParam(posLeader, null, 0) == ldefin2d.POSLEADER_OBJ)
								{
									Kompas.Instance.KompasObject.ksError(menu);
									posLeader = 0;
									flag = true;
								}
								else
									flag = false;
							}
							else
								flag = false;
							break;
						case -1:
							posLeater = Kompas.Instance.Document2D.ksFindObj(x1, y1, 100);	// величина стороны окошка-ловушки с центром x,y
							if (posLeader == 0 || Kompas.Instance.Document2D.ksGetObjParam(posLeater, null, 0) != ldefin2d.POSLEADER_OBJ)
							{
								Kompas.Instance.KompasObject.ksError("Ошибка! Объект не позиционная линия выноски!");
								posLeader = 0;
								flag = true;
							}
							else
								flag = false;
							break;
					}			
				}
				while (flag);
			
				// линия выноски есть, подключим ее к объекту спецификации
				if (posLeader > 0)
				{
					// ввойдем в режим редактирования объекта спецификации
					if (spc.ksSpcObjectEdit(spcObj) == 1)
					{
						// подключим линию выноски
						spc.ksSpcIncludeReference(posLeader, 1);
						// закроем объект спецификации
						spc.ksSpcObjectEnd();
					}
				}
			}
		}
	

		public bool IsSpcObjCreate()
		{
			return true;
		}
	

		/// <summary>
		/// Объект спецификации
		/// </summary>
		/// <param name="geom"></param>
		/// <returns></returns>
		public bool DrawSpcObj(int geom)
		{
			ksSpecification spc = (ksSpecification)Kompas.Instance.Document2D.GetSpecification();

			if (spc != null && IsSpcObjCreate())
			{
				if (Kompas.Instance.KompasObject.ksReturnResult() == (int)ErrorType.etError10)
				{
					//  10  "Ошибка! Вырожденный объект"
					Kompas.Instance.KompasObject.ksResultNULL();
					return false;
				}

				spcObj[0] = EditSpcObj(spcObj[0], geom);
				if (par.flagAttr == 0 && spcObj[0] > 0)
				{
					for (int i = 0; i < countObj; i ++)
					{
						if (spcObj[i] != 0)
						{
							// ввойдем в режим редактирования объекта спецификации
							if (spc.ksSpcObjectEdit(spcObj[i]) == 1)
							{
								// очистим геометрию, чтобы не удалялась в месте с объектом
								spc.ksSpcIncludeReference(0, 0);
								// закроем объект спецификации
								spc.ksSpcObjectEnd();
							}
							Kompas.Instance.Document2D.ksDeleteObj(spcObj[i]);
							spcObj[i] = 0;
						}
					}
				}
			}

			if (spcObj[0] > 0)
				return true;
			else
				return false;
		}


		/// <summary>
		/// Создание или редактирование объекта спецификации
		/// </summary>
		/// <param name="spcObj"></param>
		/// <returns></returns>
		public int EditSpcObj(int spcObj,int geom)
		{
			ksSpecification spc = (ksSpecification)Kompas.Instance.Document2D.GetSpecification();
			ksUserParam bufPar = (ksUserParam)Kompas.Instance.KompasObject.GetParamStruct((short)StructType2DEnum.ko_UserParam);
			ksLtVariant item = (ksLtVariant)Kompas.Instance.KompasObject.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
			ksDynamicArray arr = (ksDynamicArray)Kompas.Instance.KompasObject.GetDynamicArray(ldefin2d.LTVARIANT_ARR);

			int flMode = 0;
			bool create = false;
			float massa = 0;
			if (spc != null && bufPar != null && item != null && arr != null)
			{
				bufPar.Init();
				item.Init();
				bufPar.SetUserArray(arr);
				if (flagMode == 0 && par.flagAttr == 0)
				{
					return 0;
				}

				if (flagMode > 0)
				{
					spcObj = spc.ksGetSpcObjForGeomWithLimit("", 0, geom, 0, 1, STANDART_SECTION, 297327484710);
						if (par.flagAttr == 0)
						{
							return spcObj;
						}

					if (spcObj > 0)
					{
						if (spc.ksSpcObjectEdit(spcObj) == 0)
						{
							spcObj = 0;
						}
					}
				}

				flMode = spcObj;
				if (flMode == 0)
				{
					create = Convert.ToBoolean(spc.ksSpcObjectCreate("", 0, STANDART_SECTION, 0, 297327484710, 0));
					}

				// создаем объект спецификации для умолчательногостиля СП
				if (flMode > 0 || create)
				{
					arr.ksAddArrayItem(-1, item);
					// исполнение
					if (tmp.perform == 1)
					{
						spc.ksSpcVisible(SPC_NAME, 2, 0);
					}
					else
					{
						item.uIntVal = 2;
						arr.ksSetArrayItem(0, item);
						spc.ksSpcVisible(SPC_NAME, 2, 1);
						spc.ksSpcChangeValue(SPC_NAME, 2, bufPar, ldefin2d.UINT_ATTR_TYPE);
					}

					// изменим диаметр
					item.floatVal = tmp.dr;
					arr.ksSetArrayItem(0, item);
					spc.ksSpcChangeValue(SPC_NAME, 4, bufPar, ldefin2d.FLOAT_ATTR_TYPE);

					// отследим мелкий шаг
					if (!tmp.pitch)	// выключить шаг и его разделитель
					{
						spc.ksSpcVisible(SPC_NAME, 5, 0);
						spc.ksSpcVisible(SPC_NAME, 6, 0);	// щаг
					}
					else
					{
						spc.ksSpcVisible(SPC_NAME, 5, 1);
						spc.ksSpcVisible(SPC_NAME, 6, 1);	// щаг
						item.floatVal = tmp.p;
						arr.ksSetArrayItem(0, item);
						spc.ksSpcChangeValue(SPC_NAME, 6, bufPar, ldefin2d.FLOAT_ATTR_TYPE);
					}
				
					// выключим поле допуска
					if (flMode == 0)
					{
						spc.ksSpcVisible(SPC_NAME, 7, 0);
						// выключим класс прочности
						spc.ksSpcVisible(SPC_NAME, 8, 0);
						// выключим материал
						spc.ksSpcVisible(SPC_NAME, 9, 0);
						// выключим покрытие
						spc.ksSpcVisible(SPC_NAME, 10, 0);
					}
				
					// изменим ГОСТ
					item.uIntVal = tmp.gost;
				
					arr.ksSetArrayItem(0, item);
					spc.ksSpcChangeValue(SPC_NAME, 12, bufPar, ldefin2d.UINT_ATTR_TYPE);
					switch (tmp.indexMassa)
					{
						case 0: massa = tmp.massa;								break;
						case 1: massa = Convert.ToSingle(0.356 * tmp.massa);	break;
						case 2: massa = Convert.ToSingle(1.08 * tmp.massa);		break;
					}
				
					spc.ksSpcMassa(massa.ToString());// масса детали
		
						// подключим геометрию
						if (geom > 0)
							spc.ksSpcIncludeReference(geom, SPC_CLEAR_GEOM);
				
					return spc.ksSpcObjectEnd();
				}
			}
	
			return 0;
		}
	

		/// <summary>
		/// Инициализация параметров пользователя
		/// </summary>
		public void InitUserParamTmp()
		{
			if (paramTmp != null)
			{
				ksLtVariant item = (ksLtVariant )Kompas.Instance.KompasObject.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
				ksDynamicArray arr = (ksDynamicArray)Kompas.Instance.KompasObject.GetDynamicArray(ldefin2d.LTVARIANT_ARR);
				if (item != null && arr != null)
				{
					paramTmp.Init();
					paramTmp.SetUserArray(arr);
					item.Init();
					item.floatVal = tmp.dr;		// 0 - dr
					arr.ksAddArrayItem(-1, item);
					item.floatVal = tmp.p;		// 1 - p1
					arr.ksAddArrayItem(-1, item);
					item.floatVal = tmp.p;		// 2 - p2
					arr.ksAddArrayItem(-1, item);
					item.floatVal = tmp.s;		// 3 - s
					arr.ksAddArrayItem(-1, item);
					item.floatVal = tmp.d;		// 4 - D
					arr.ksAddArrayItem(-1, item);
					item.floatVal = tmp.da;		// 5 - da
					arr.ksAddArrayItem(-1, item);
					item.floatVal = tmp.H;		// 6 - h
					arr.ksAddArrayItem(-1, item);
					item.floatVal = tmp.d2;		// 7 - d2
					arr.ksAddArrayItem(-1, item);
					item.floatVal = tmp.massa;	// 8 - m
					arr.ksAddArrayItem(-1, item);
					item.floatVal = tmp.s;		// 9 - s1
					arr.ksAddArrayItem(-1, item);
					item.floatVal = tmp.d;		// 10 - D1
					arr.ksAddArrayItem(-1, item);
					item.floatVal = tmp.massa;	// 11 - m1
					arr.ksAddArrayItem(-1, item);
				}
			}
		}
	

		public void GetUserParamTmp()
		{
			if (paramTmp != null)
			{
				ksLtVariant item = (ksLtVariant)Kompas.Instance.KompasObject.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
				ksLtVariant item1 = (ksLtVariant)Kompas.Instance.KompasObject.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
				ksDynamicArray arr = (ksDynamicArray)paramTmp.GetUserArray();
				if (item != null && item1 != null && arr != null && arr.ksGetArrayCount() >= 12)
				{
					item.Init();
					arr.ksGetArrayItem(0, item);	// dr
					tmp.dr = item.floatVal;
				
					arr.ksGetArrayItem(5, item);	// da
					tmp.da = item.floatVal;
				
					arr.ksGetArrayItem(6, item);	// h
					tmp.H = item.floatVal;
				
					arr.ksGetArrayItem(7, item);	// d2
					tmp.d2 = item.floatVal;
				
					arr.ksGetArrayItem(1, item);	// p1
					arr.ksGetArrayItem(2, item1);	// p2
					if (tmp.pitch)
						tmp.p = item1.floatVal;
					else
						tmp.p = item.floatVal;
				
					if (System.Math.Abs(item1.floatVal - item.floatVal) < 0.001)
						tmp.pitch_off = 1;
					else
						tmp.pitch_off = 0;
				
					arr.ksGetArrayItem(9, item1);	// s1
					if (tmp.key_s_on == 1 && System.Math.Abs(item1.floatVal) > 0.001)
						tmp.key_s_gray = 1;
					else
						tmp.key_s_gray = 0;
					tmp.key_s = false;
				
					if (tmp.key_s_on == 1 && tmp.key_s)
					{
						tmp.s = item1.floatVal;			// размер под ключ
						arr.ksGetArrayItem(10, item1);	// D1
						tmp.d = item1.floatVal;			// диаметр описанной окружности
						arr.ksGetArrayItem(11, item1);	// m1
						tmp.massa = item1.floatVal;
					}
					else
					{
						arr.ksGetArrayItem(3, item1);	// s
						tmp.s = item1.floatVal;			// размер под ключ
						arr.ksGetArrayItem(4, item1);	// D
						tmp.d = item1.floatVal;			// диаметр описанной окружности
						arr.ksGetArrayItem(8, item1);	// m
						tmp.massa = item1.floatVal;
					}
				}
			}
		}
	
			
		/// <summary>
		/// Подключиться к базе данных
		/// </summary>
		/// <param name="bd"></param>
		/// <param name="name"></param>
		/// <returns></returns>
		public bool ConnectDB_(int bd, string name)
		{
			string buf = string.Empty;
			if (GetFullName(name, ref buf))
				return Convert.ToBoolean(data.ksConnectDB(base_.bg, buf));
			else
				return Convert.ToBoolean(data.ksConnectDB(base_.bg, name));
		}
	

		/// <summary>
		/// Получить полное имя файла
		/// </summary>
		/// <param name="name"></param>
		/// <param name="buf"></param>
		/// <returns></returns>
		public bool GetFullName(string name, ref string buf)
		{
			bool result = false;

			FileInfo fInfo = new FileInfo(Kompas.Instance.KompasObject.ksSystemPath(ldefin2d.sptLIBS_FILES) + @"\Load\" + name);
			if (fInfo.Exists)
			{
				result = true;
				buf = fInfo.FullName;
			}
			else
			{
				fInfo = new FileInfo(Assembly.GetExecutingAssembly().Location);
				fInfo = new FileInfo(fInfo.DirectoryName + @"\Load\" + name);
				if (fInfo.Exists)
				{
					result = true;
					buf = fInfo.FullName;
				}
				else
				{
					fInfo = new FileInfo(Assembly.GetExecutingAssembly().Location);
					fInfo = new FileInfo(fInfo.DirectoryName + @"\" + name);
					if (fInfo.Exists)
					{
						result = true;
						buf = fInfo.FullName;
					}
					else
						result = false;
				}
			}

			return result;
		}
	

		/// <summary>
		/// Открыть базу данных
		/// </summary>
		/// <returns></returns>
		public bool OpenGaykaBase()
		{
			bool result = false;

			InitUserParamTmp();
			if (data != null)
			{
				base_.bg = data.ksCreateDB("TXT_DB");	// TXT_DB
			
				if (ConnectDB_(base_.bg, "5915.loa"))	// 5915.loa
				{
					base_.rg = data.ksRelation(base_.bg);
					data.ksRFloat("dr");
					data.ksRFloat("");
					data.ksRFloat("");
					data.ksRFloat("");
					data.ksRFloat("");
					data.ksRFloat("");
					data.ksRFloat("");
					data.ksRFloat("");
					data.ksRFloat("");
					data.ksRFloat("");
					data.ksRFloat("");
					data.ksRFloat("");
					data.ksEndRelation();
					if (data.ksDoStatement(base_.bg, base_.rg, "") == 1)	// TXT_ALL
						result = true;
				}
			}

			return result;
		}
	

		/// <summary>
		/// Закрыть базу данных
		/// </summary>
		public void CloseGaykaBase()
		{
			if (base_.bg != 0)
				data.ksDeleteDB(base_.bg);
		}


		/// <summary>
		/// Читать параметры гайки по диаметру
		/// возвращает 1 - успех 0 - не найдено записи , ошибка связи с БД
		/// </summary>
		/// <param name="d"></param>
		/// <returns></returns>
		public bool ReadGaykaBase(float d)
		{
			bool result = false;

			string s;
			if (d > 0)
				s = string.Format("dr={0}", d);
			else
				s = string.Format("dr={0}", tmp.d);

			int i = 0;
			if (data.ksCondition(base_.bg, base_.rg, s) == 1)
			{
				i = data.ksReadRecord(base_.bg, base_.rg, paramTmp);
				if (i > 0)
				{
					GetUserParamTmp();
					result = true;
				}
				else
					result = false;
			}

			return result;
		}
	

		/// <summary>
		/// Инициализация параметров структуры гайки
		/// </summary>
		public void init()
		{
			tmp.gost = 5915;
			tmp.hatchAng = 45;
			tmp.hatchStep = 2;
			tmp.dr = 20;
			tmp.p = 2.5F;
			tmp.ver = 1;
			tmp.indexMassa = 0;
			tmp.s = 30;
			tmp.d = 33;
			tmp.da = 21.6F;
			tmp.H = 16;
			tmp.d2 = 27.7F;
			tmp.cls = 2;	// класс точности
			tmp.perform = 0;
			tmp.axis_off = 0;
			tmp.simple = false;
			tmp.massa = 71.44F;
		}
	
			
		/// <summary>
		/// Инициализация объекта класса
		/// </summary>
		private void Class_Initialize()
		{
			Param = (ksUserParam)Kompas.Instance.KompasObject.GetParamStruct((short)StructType2DEnum.ko_UserParam);
			paramTmp = (ksUserParam)Kompas.Instance.KompasObject.GetParamStruct((short)StructType2DEnum.ko_UserParam);
			data = (ksDataBaseObject)Kompas.Instance.KompasObject.DataBaseObject();

			refMacr = 0;
			flagMode = 0;
			countObj = 1;
		
			InitUserParam();
		
			if (Kompas.Instance.Document2D.ksEditMacroMode() > 0 && Kompas.Instance.Document2D.ksGetMacroParam(0, Param) > 0)
			{
				GetUserParam();	// инициализация из макропараметров
				tmp.ver = 1;
				tmp.key_s_on = 1;
				tmp.koef_mat_on = 1;
			}
			else
			{
				// инициализация по умолчанию
				par.drawType = ID_VID;
				par.ang = 0;
				par.flagAttr = 0;
				init();	// инициализация параметров структуры гайки
			}
		}

		public GaykaObj()
		{
			Class_Initialize();
		}


		private bool processParam_ChangeControlValue(IPropertyControl control)
		{
			bool redrawPhantom = true;

			float val_Renamed;
			switch (control.Id)
			{
				case DIAM_ID:	// Комбобокс диаметров резьбы
					val_Renamed = Convert.ToSingle(control.Value);
					tmp.dr = val_Renamed;	// Установим диаметр гайке
					ReadGaykaBase(val_Renamed);	// Прочитаем новые параметры из базы
					redrawPhantom = true;	// перерисовать фантом
					break;
				case SPC_CHECK_ID:	// Чекбокс объекта сп
					par.flagAttr = Convert.ToInt16(spcCheck.Value);
					break;
				case ANGLE_HATCH_ID:	// угол штриховки
					tmp.hatchAng = Convert.ToSingle(angleEdit.Value);
					redrawPhantom = true;	// перерисовать фантом
					break;
				case STEP_HATCH_ID:	// шаг штриховки
					tmp.hatchStep = Convert.ToSingle(stepEdit.Value);
					redrawPhantom = true;	// перерисовать фантом
					break;
			}
		
			if (redrawPhantom)
				RedrawPhantomProc();

			return true;
		}
	
		private bool processParam_ControlCommand(IPropertyControl control, int buttonID)
		{
			bool redrawPhantom = true;

			switch (control.Id)
			{
				case VIEWSIDE_ID:	// Кнопки отображения
					if (buttonID == BASE_VIEW)
					{
						// главный вид
						par.drawType = ID_VID;
						angleEdit.Enable = false;
						stepEdit.Enable = false;
					}
					if (buttonID == LEFT_VIEW)	// вид сбоку
					{
						par.drawType = ID_SIDEVID;
						angleEdit.Enable = false;
						stepEdit.Enable = false;
					}
					if (buttonID == TOP_VIEW)	// вид сверху
					{
						par.drawType = ID_TOPVID;
						angleEdit.Enable = false;
						stepEdit.Enable = false;
					}
					if (buttonID == SEC_VIEW)	// вид \ разрез
					{
						par.drawType = ID_VIDSEC;
						angleEdit.Enable = true;
						stepEdit.Enable = true;
					}
					redrawPhantom = true;	// перерисовать фантом
					break;
				case PERFORMANCE_ID:	// Кнопки исполнений
					if (buttonID == PERF2_VIEW)	// исполнение 2
						tmp.perform = 1;
					else
						tmp.perform = 0;	// исполнение 1
		
					redrawPhantom = true;	// перерисовать фантом
					break;
				case SIMPLES_ID:	// Кнопки упрощений
					if (buttonID == SIMPLE_VIEW)	// Упрошено
					{
						tmp.simple = simpButt.get_ButtonChecked(SIMPLE_VIEW);
						redrawPhantom = true;	// перерисовать фантом
					}
					if (buttonID == DRAW_AXIS)	// Рисовать ось
					{
						if (simpButt.get_ButtonChecked(DRAW_AXIS))
							tmp.axis_off = 0;
						else
							tmp.axis_off = 1;
					}
					redrawPhantom = true;	// перерисовать фантом
					break;
				case ADD_PARAM_ID:	// Кнопки доп. параметров
					if (buttonID == ADDSTEP_BUTT)	// Мелкий шаг
						tmp.pitch = addButt.get_ButtonChecked(ADDSTEP_BUTT);
					if (buttonID == KEY_BUTT)	// Доп. размер под ключ
						tmp.key_s = addButt.get_ButtonChecked(KEY_BUTT);
					break;
			}
			
			if (redrawPhantom)
				RedrawPhantomProc();

			return true;
		}
	
		private void RedrawPhantomProc()
		{
			ksType1 pt1 = (ksType1)phantom.GetPhantomParam();
			int gr = 0;
			GetGroup(ref gr);	// новая геометрия
			pt1.gr = gr;
		
			Kompas.Instance.Document2D.ksChangeObjectInLibRequest(null, phantom);
		}
	
		private bool userCtrl_CreateOCX(object iOcx)
		{
			hatchPar = (HatchControl)iOcx;
			hatchPar.AngleChanged += new AngleChangedHandler(hatchPar_AngleChanged);
			hatchPar.StepChanged += new StepChangedHandler(hatchPar_StepChanged);
			hatchPar.AngleEditValue = tmp.hatchAng.ToString();
      hatchPar.StepEditValue = tmp.hatchStep.ToString();
      return true;
		}
		private bool userCtrl_DestroyOCX()
		{
			hatchPar = null;
			return true;
		}

		private void hatchPar_AngleChanged()
		{
			if (hatchPar != null )
				tmp.hatchAng = (float)System.Convert.ToDouble(hatchPar.AngleEditValue);
			Kompas.Instance.KompasObject.ksMessage("Изменили угол");
		}

		private void hatchPar_StepChanged()
		{
      if (hatchPar != null)
        tmp.hatchAng = (float)System.Convert.ToDouble(hatchPar.StepEditValue);
      Kompas.Instance.KompasObject.ksMessage("Изменили шаг");
		}
	}
}
