using System;

namespace KAPITypes
{

	/// <summary>
	/// типы документов
	/// </summary>
	public enum DocType 
	{
		/// <summary>
		/// 0- нет активного документа
		/// </summary>
		lt_DocUnknown    = 0,
		/// <summary>
		/// 1- чертеж стандартный
		/// </summary>
		lt_DocSheetStandart = 1,
		/// <summary>
		/// 2- чертеж нестандартный
		/// </summary>
		lt_DocSheetUser,
		/// <summary>
		/// 3- фрагмент
		/// </summary>
		lt_DocFragment,
		/// <summary>
		/// 4- спецификация
		/// </summary>
		lt_DocSpc,
		/// <summary>
		/// 5- 3d-документ модель
		/// </summary>
		lt_DocPart3D,
		/// <summary>
		/// 6- 3d-документ сборка
		/// </summary>
		lt_DocAssemble3D,
		/// <summary>
		/// 7- текстовый документ стандартный
		/// </summary>
		lt_DocTxtStandart,
		/// <summary>
		/// 8- текстовый документ нестандартный
		/// </summary>
		lt_DocTxtUser,
		/// <summary>
		/// 9- спецификация нестандартный формат
		/// </summary>
		lt_DocSpcUser,
		/// <summary>
		/// 10- 3d-документ технологическая сборка
		/// </summary>
		lt_DocTechnologyAssemble3D,
	}


	/// <summary>
	/// система квалитета
	/// </summary>
	public enum LtQualSystem 
	{
		/// <summary>
		/// 1 - вала
		/// </summary>
		lt_qsShaft = 1,
		/// <summary>
		/// 2 - отверстия
		/// </summary>
		lt_qsHole = 2,
	}


	/// <summary>
	/// квалитеты
	/// </summary>
	public enum LtQualDir 
	{
		/// <summary>
		/// 1 - предпочтительные
		/// </summary>
		lt_qdPreferable = 1,
		/// <summary>
		/// 2 - основные
		/// </summary>
		lt_qdBasic,
		/// <summary>
		/// 3 - дополнительные
		/// </summary>
		lt_qdAdditional,
	}


	/// <summary>
	/// типы данных для LtVariant
	/// </summary>
	public enum LtVariantType 
	{
		/// <summary>
		/// 1 - символ
		/// </summary>
		ltv_Char = 1,
		/// <summary>
		/// 2 - байт
		/// </summary>
		ltv_UChar,
		/// <summary>
		/// 3 - целое
		/// </summary>
		ltv_Int,
		/// <summary>
		/// 4 - беззнаковое целое
		/// </summary>
		ltv_UInt,
		/// <summary>
		/// 5 - длинное целое
		/// </summary>
		ltv_Long,
		/// <summary>
		/// 6 - вешественное
		/// </summary>
		ltv_Float,
		/// <summary>
		/// 7 - двойное вешественное
		/// </summary>
		ltv_Double,
		/// <summary>
		/// 8 - строка 255 символов
		/// </summary>
		ltv_Str,
		/// <summary>
		/// 9 - пока не используется
		/// </summary>
		ltv_NoUsed,
		/// <summary>
		/// 10 - короткое целое
		/// </summary>
		ltv_Short,
		/// <summary>
		/// 11 - строка 255 символов Unicode
		/// </summary>
		ltv_WStr,
	}


	/// <summary>
	/// типы точек привязки текста
	/// </summary>
	public enum TextAlign 
	{
		/// <summary>
		/// точка привязки слева
		/// </summary>
		txta_Left = 0,
		/// <summary>
		/// точка привязки вцентре
		/// </summary>
		txta_Center,
		/// <summary>
		/// точка привязки справа
		/// </summary>
		txta_Right
	}


	
	/// <summary>
	/// перечисление возможных типов узла дерева библиотеки документов
	/// </summary>
	public enum LtNodeType 
	{
		/// <summary>
		/// корень дерева
		/// </summary>
		tn_root,
		/// <summary>
		/// папка (директория)
		/// </summary>
		tn_dir,
		/// <summary>
		/// документ (файл)
		/// </summary>
		tn_file
	}



	/// <summary>
	/// типы значка объекта "Выносной элемент"
	/// </summary>
	public enum LtRemoteElmSignType 
	{
		/// <summary>
		/// 0 - окружность
		/// </summary>
		re_Circle,
		/// <summary>
		/// 1 - прямоугольник
		/// </summary>
		re_Rectangle,
		/// <summary>
		/// 2 - скругленный прямоугольник
		/// </summary>
		re_Ballon
	}


	/// <summary>
	/// Тип изменения порядка объектов
	/// </summary>
	public enum ChangeOrderType 
	{
		/// <summary>
		/// Выше всех
		/// </summary>
		co_Top = 1,
		/// <summary>
		/// Ниже всех 
		/// </summary>
		co_Bottom,
		/// <summary>
		/// Перед объектом 
		/// </summary>
		co_BeforeObject,
		/// <summary>
		/// За объектом
		/// </summary>
		co_AfterObject,
		/// <summary>
		/// На уровень вперед
		/// </summary>
		co_UpLevel,
		/// <summary>
		/// На уровень назад
		/// </summary>
		co_DownLevel,
    /// <summary>
    /// Перед объектом (расположить объекты группы в порядке их добавления в группу)
    /// </summary>
    co_BeforeObjectByGroupOrder,
    /// <summary>
    /// За объектом (расположить объекты группы в порядке их добавления в группу)
    /// </summary>
    co_AfterObjectByGroupOrder

  }


  public class ldefin2d
	{
		public const int TEXT_LENGTH		= 128;
		public const int MAX_TEXT_LENGTH	= 255;

		public const int ODBC_DB	= 0;
		public const int TXT_DB		= 1;

		public const int TXT_CHAR	= 1;
		public const int TXT_USHORT = 2;
		public const int TXT_SSHORT = 3;
		public const int TXT_SLONG	= 4;
		public const int TXT_ULONG	= 5;
		public const int TXT_LONG	= 6;
		public const int TXT_FLOAT	= 7;
		public const int TXT_DOUBLE = 8;
		public const int TXT_INT	= 9;
		public const int TXT_ALL	= 0;
//		public const string TXT_INDEX  "Index1000"

		public const int stACTIVE		= 0;  //состояние для вида, слоя, документа
		public const int stREADONLY		= 1;  //состояние для вида, слоя
		public const int stINVISIBLE	= 2;  //состояние для вида, слоя
		public const int stCURRENT		= 3;  //состояние для вида, слоя
		public const int stPASSIVE		= 1;  //состояние для документа

		// Определения для функции ksSystemPath
		public const int sptSYSTEM_FILES          = 0;  // Выдать путь на каталог системных файлов
		public const int sptLIBS_FILES			      = 1;  // Выдать путь на каталог файлов библиотек
    public const int sptAPPS_FILES            = 1;  // Выдать путь на каталог файлов приложений
    public const int sptTEMP_FILES			      = 2;  // Выдать путь на каталог сохранения временных файлов
		public const int sptCONFIG_FILES		      = 3;  // Выдать путь на каталог сохранения конфигурации системы
		public const int sptINI_FILE			        = 4;  // Выдать полное имя INI-файла системы
		public const int sptBIN_FILE			        = 5;  // Выдать путь на каталог исполняемых файлов системы
		public const int sptPROJECT_FILES		      = 6;  // Выдать путь на каталог сохранения kompas.prj
		public const int sptDESKTOP_FILES		      = 7;  // Выдать путь на каталог сохранения kompas.dsk
		public const int sptTEMPLATES_FILES		    = 8;  // Выдать путь на каталог шаблонов Компас-документов
		public const int sptPROFILES_FILES		    = 9;  // Выдать путь на каталог сохранения профилей пользователя
    public const int sptWORK_FILES            = 10; // Выдать путь на каталог "Мои документы"
    public const int sptSHEETMETAL_FILES      = 11; // Выдать путь на каталог таблиц сгибов
    public const int sptPARTLIB_FILES         = 12; // Выдать путь на каталог PartLib
    public const int sptMULTILINE_FILES       = 13; // Выдать путь к каталогу шаблонов мультилинии
    public const int sptPRINTDEVICE_FILES     = 14; // Выдать путь к каталогу конфигураций плоттеров/принтеров
    public const int sptCURR_WORK_FILES       = 15; // запоминание последних директориев, с которых выполнилось открытие/сохранение файла в диалоге Open|Save
    public const int sptCURR_LIBS_FILES       = 16; // запоминание последних директориев, с которых выполнилось открытие/сохранение файла в диалоге Open|Save
    public const int sptCURR_APPS_FILES       = 16; // запоминание последних директориев, с которых выполнилось открытие/сохранение файла в диалоге Open|Save
    public const int sptCURR_SYSTEM_FILES     = 17; // запоминание последних директориев, с которых выполнилось открытие/сохранение файла в диалоге Open|Save
    public const int sptCURR_PROFILES_FILES   = 18; // запоминание последних директориев, с которых выполнилось открытие/сохранение файла в диалоге Open|Save
    public const int sptCURR_SHEETMETAL_FILES = 19; // запоминание последних директориев, с которых выполнилось открытие/сохранение файла в диалоге Open|Save
    public const int sptMULTY_APPS_FILES      = 20; // Выдать список каталогов файлов приложений
    public const int sptDOC_LIBS_FILES        = 21; // Выдать путь на каталог файлов библиотек документов
    public const int sptMULTY_DOC_LIBS_FILES  = 22; // Выдать список каталогов файлов библиотек документов
    public const int sptCUR_DOC_LIBS_FILES    = 23; // запоминание последних директориев, с которых выполнилось открытие/сохранение файла в диалоге Open|Save
    public const int sptUTILS_FILES           = 24; // Выдать путь на каталог утилит
    public const int sptMULTY_UTILS_FILES     = 25; // Выдать список каталогов утилит
    public const int sptCURR_UTILS_FILES      = 26; // запоминание последних директориев, с которых выполнилось открытие/сохранение файла в диалоге Open|Save
    public const int sptPROGRAMDATA_FILES     = 27; // Выдать путь на каталог данных приложения ( C:\ProgramData\ASCON\KOMPAS-3D\18 )


    // Определения для результата функции SystemControlStart
    public const int scsSTOPPED_FOR_MENU_COMMAND		= ( 1); // Выполнена команда меню "Остановить работу библиотеки"
		public const int scsSTOPPED_FOR_SYSTEM_STOP			= ( 0); // Идет закрытие задачи
		public const int scsSTOPPED_FOR_ITSELF				= (-1); // Вызов функции SystemControlStop из-под библиотеки
		public const int scsSTOPPED_FOR_START_THIS_LIB		= (-2); // Принудительный останов при запуске той же библиотеки		
		public const int scsSTOPPED_FOR_START_ANOTHER_LIB	= (-3); // Принудительный останов при запуске другой библиотеки		

		//  Определения для функций GetObjParam и SetObjParam
		//  '+'  отмечены объекты, для которых реализованы  GetObjParam и SetObjParam
		public const int ALLPARAM					= -1;  // все параметры объекта в СК объекта владельца
		public const int SHEET_ALLPARAM              = -2;  // тоже что и  ALLPARAM  но параметры объекта в СК листа
		public const int NURBS_CLAMPED_ALLPARAM      = -5;  // параметры нурбса, преобразовать узловой вектор в зажатый  
		public const int NURBS_CLAMPED_SHEET_ALLPARAM= -6;  // параметры нурбса в СК листа, преобразовать узловой вектор в зажатый
		public const int VIEW_ALLPARAM               = -7;  // все параметры объекта в СК вида

		public const int ANGLE_ARC_PARAM       =  0;   // параметры дуги по углам ( для дуги и эллиптической дуги ) в СК объекта владельца
		public const int POINT_ARC_PARAM       =  1;   // параметры дуги по точкам ( для дуги и эллиптической дуги ) в СК объекта владельца
		public const int ANGLE_ARC_SHEET_PARAM =  2;   // параметры дуги по углам ( для дуги и эллиптической дуги ) в СК листа
		public const int POINT_ARC_SHEET_PARAM =  3;   // параметры дуги по точкам ( для дуги и эллиптической дуги ) в СК листа
		public const int ANGLE_ARC_VIEW_PARAM  =  4;   // параметры дуги по углам ( для дуги и эллиптической дуги ) в СК вида
		public const int POINT_ARC_VIEW_PARAM  =  5;   // параметры дуги по точкам ( для дуги и эллиптической дуги ) в СК вида

		public const int VIEW_LAYER_STATE			= 1;   // состояние слоя ,вида
		public const int DOCUMENT_STATE				= 1;   // состояние документа
		public const int DOCUMENT_SIZE				= 0;   // размер листа
		public const int DIM_TEXT_PARAM				= 0;   // параметры текста для размеров
		public const int DIM_SOURSE_PARAM			= 1;   // параметры привязки размера
		public const int DIM_DRAW_PARAM				= 2;   // параметры отрисовки размера
		public const int DIM_VALUE					= 3;   // значение размера - double
		public const int DIM_PARTS					= 4;   // составляющие части для размеров struct DimensionPartsParam
		public const int SHEET_DIM_PARTS			= 5;   // составляющие части для размеров struct DimensionPartsParam в СК листа
		public const int TECHNICAL_DEMAND_PAR		= -1;  // параметры технических требований -
		public const int TT_FIRST_STR				= 1000;// начало отсчета для получения или замены текста ТТ по строкам
		public const int CONIC_PARAM				= 2;   // параметры для построения конического сечения ( для эллипса и эллтптической дуги )
		public const int SPC_TUNING_PARAM			= 0;   // параметры настроек для стиля СП
		public const int HATCH_PARAM_EX				= 0;   // параметры штриховки расширенные
		public const int ASSOCIATION_VIEW_PARAM		= 0;   // параметры ассоциативного вида
		public const int DIM_SOURSE_VIEWPARAM     = 7;   // параметры привязки размера в ситеме координат вида
		public const int DIM_DRAW_VIEWPARAM       = 8;   // параметры отрисовки размера в ситеме координат вида
		public const int DIM_SOURSE_SHEETPARAM    = 9;   // параметры привязки размера в ситеме координат листа
		public const int DIM_DRAW_SHEETPARAM      = 10;  // параметры отрисовки размера в ситеме координат листа

		public const int ALL_OBJ				= 0;         // все объекты,кроме вспомогательных, которые могут входить в вид                    -
		public const int LINESEG_OBJ			= 1;         // отрезок                        +
		public const int CIRCLE_OBJ				= 2;         // окружность                     +
		public const int ARC_OBJ				= 3;         // дуга                           +
		public const int TEXT_OBJ				= 4;         // текст                          +
		public const int POINT_OBJ				= 5;         // точка                          +
		public const int HATCH_OBJ				= 7;         // штриховка                      +
		public const int BEZIER_OBJ				= 8;         // bezier сплайн                  +
		public const int LDIMENSION_OBJ			= 9;         // линейный размер                +
		public const int ADIMENSION_OBJ			= 10;        // угловой размер                 +
		public const int DDIMENSION_OBJ			= 13;        // диаметральный размер           +
		public const int RDIMENSION_OBJ			= 14;        // радиальный размер              +
		public const int RBREAKDIMENSION_OBJ	= 15;        // радиальный размер с изломом    +
		public const int ROUGH_OBJ				= 16;        // шероховатость                  +
		public const int BASE_OBJ				= 17;        // база                           +
		public const int WPOINTER_OBJ			= 18;        // стрелка вида                   +
		public const int CUT_OBJ				= 19;        // линия разреза                  +
		public const int LEADER_OBJ				= 20;        // простая линия выноски          +
		public const int POSLEADER_OBJ			= 21;        // линия выноски для обозначения позиции      +
		public const int BRANDLEADER_OBJ		= 22;        // линия выноски для обозначения клеймения    +
		public const int MARKERLEADER_OBJ		= 23;        // линия выноски для обозначения маркирования +
		public const int TOLERANCE_OBJ			= 24;        // допуск формы                   +
		public const int TABLE_OBJ              = 25;        // таблица                        -     //тексты
		public const int CONTOUR_OBJ            = 26;        // контур                         +     //стиль
		public const int MACRO_OBJ              = 27;        // нетипизированный макроэлемент  -
		public const int LINE_OBJ               = 28;        // линия                          +
		public const int LAYER_OBJ              = 29;        // слой                           +
		public const int FRAGMENT_OBJ           = 30;        // вставной фрагмент              +
		public const int POLYLINE_OBJ           = 31;        // полилиния                      +
		public const int ELLIPSE_OBJ            = 32;        // эллипс                         +
		public const int NURBS_OBJ              = 33;        // nurbs сплайн                   +
		public const int ELLIPSE_ARC_OBJ        = 34;        // дуга эллипса                   +
		public const int RECTANGLE_OBJ          = 35;        // прямоугольник                  +
		public const int REGULARPOLYGON_OBJ     = 36;        // многоугольник                  +
		public const int EQUID_OBJ              = 37;        // эквидистанта                   +
		public const int LBREAKDIMENSION_OBJ    = 38;        // линейный размер с обрывом      +
		public const int ABREAKDIMENSION_OBJ    = 39;        // угловой размер с обрывом       +
		public const int ORDINATEDIMENSION_OBJ  = 40;        // размер высоты
		public const int COLORFILL_OBJ          = 41;        // фоновая заливка цветом         +
		public const int CENTREMARKER_OBJ       = 42;        // обозначение центра             +
		public const int ARCDIMENSION_OBJ       = 43;        // размер длины дуги
		public const int SPC_OBJ                = 44;        // объект спецификации            +
		public const int RASTER_OBJ             = 45;        // растровый объект               +
		public const int CHANGE_LEADER_OBJ      = 46;        // Обозначение изменения          -
		public const int REMOTE_ELEMENT_OBJ     = 47;        // выносной элемент               +
		public const int AXISLINE_OBJ           = 48;        // Осевая линия                   +
		public const int OLEOBJECT_OBJ          = 49;        // вставка ole объекта            -
    public const int KNOTNUMBER_OBJ         = 50;        // объект номер узла              -
    public const int BRACE_OBJ              = 51;        // объект фигурная скобка         -
    public const int POSNUM_OBJ             = 52;        // марка/позиционное обозначение с линией-выноской - 
    public const int MARKONLDR_OBJ          = 53;        // марка/позиционное обозначение на линии          -
    public const int MARKWOLDR_OBJ          = 54;        // марка/позиционное обозначение без линии-выноски -
    public const int WAVELINE_OBJ           = 55;        // волнистая линия                -
    public const int DIRAXIS_OBJ            = 56;        // прямая ось                     -
    public const int BROKENLINE_OBJ         = 57;        // линия обрыва с изломами        -
    public const int CIRCLEAXIS_OBJ         = 58;        // круговая ось                   -
    public const int ARCAXIS_OBJ            = 59;        // дуговая ось                    -
    public const int CUTUNITMARKING         = 60;        // Обозначение узла в сечении     -
    public const int UNITMARKING            = 61;        // Обозначение узла      -
    public const int MULTITEXTLEADER        = 62;        // Выносная надпись к многослойным конструкциям.      -
    public const int EXTERNALVIEW_OBJ       = 63;        // Вставка внешнего вида                              -
    public const int ANNLINESEG_OBJ         = 64;        // Аннотационный отрезок                 +- Для GetObjParam используется структура LineSegParam
    public const int ANNCIRCLE_OBJ          = 65;        // Аннотационная окружность              +- Для GetObjParam используется структура CircleParam
    public const int ANNELLIPSE_OBJ         = 66;        // Аннотационный эллипс                  +- Для GetObjParam используется структура EllipseParam
    public const int ANNARC_OBJ             = 67;        // Аннотационная дуга                    +- Для GetObjParam используется структура ArcParam
    public const int ANNELLIPSE_ARC_OBJ     = 68;        // Аннотационная дуга эллипса            +- Для GetObjParam используется структура EllipseArcParam
    public const int ANNPOLYLINE_OBJ        = 69;        // Аннотационная полилиния               +- Для GetObjParam используется структура PolylineParam
    public const int ANNPOINT_OBJ           = 70;        // Аннотационная точка                   +- Для GetObjParam используется структура PointParam
    public const int ANNTEXT_OBJ            = 71;        // Текст с аннотационной точкой привязки +- Для GetObjParam используется структура TextParam
    public const int MULTILINE_OBJ          = 72;        // Мультилиния                    -
    public const int BUILDINGCUTLINE_OBJ    = 73;        // Линия разреза/сечения для СПДС + используется структура CutLineParam
    public const int ATTACHED_LEADER_OBJ    = 74;        // Присоединенная линия выноски ( не имеет текстов )  +
    public const int CONDITIONCROSSING_OBJ  = 75;        // Условное пересечение           -
    public const int REPORTTABLE_OBJ        = 76;        // ассоциативная таблица отчета
    public const int EMBODIMENTSTABLE_OBJ   = 77;        // таблица исполнений
    public const int SPECIALCURVE_OBJ       = 78;        // Кривая общего вида
    public const int ARRAYPARAMTABLE_OBJ    = 79;        // Таблица параметров массива
    public const int NURBS_BY_POINTS_OBJ    = 80;        // NURBS-кривая по точкам                        +- Для GetObjParam используется структура NurbsParam
    public const int CONIC_CURVE_OBJ        = 81;        // Коническая кривая                             +- Для GetObjParam используется структура NurbsParam
    public const int CONIC_CURVE_4POINT_OBJ = 82;        // Коническая кривая  по 4 точками и направлению +- Для GetObjParam используется структура NurbsParam
    public const int CONIC_CURVE_5POINT_OBJ = 83;        // Коническая кривая  по 5 точками               +- Для GetObjParam используется структура NurbsParam
    public const int CIRCULAR_CENTRES_OBJ   = 84;        // Круговая сетка центров
    public const int LINEAR_CENTRES_OBJ     = 85;        // Линейная сетка центров

    public const int MAX_VIEWTIP_SEARCH     = 85;        // верхняя граница типов поиска для объектов вида  -

		public const int SPECIFICATION_OBJ      = 121;       // спецификация на листе
		public const int SPECROUGH_OBJ          = 122;       // неуказанная шероховатость      +
		public const int VIEW_OBJ               = 123;       // вид                            +
		public const int DOCUMENT_OBJ           = 124;       // документ  графический          +   (лист или фрагмент)
		public const int TECHNICALDEMAND_OBJ    = 125;       // технические требования         +
		public const int STAMP_OBJ              = 126;       // штамп                          -  //тексты
		public const int SELECT_GROUP_OBJ       = 127;       // группа селектирования          -
		public const int NAME_GROUP_OBJ         = 128;       // именная группа                 -
		public const int WORK_GROUP_OBJ         = 129;       // рабочая группа                 -
		public const int SPC_DOCUMENT_OBJ       = 130;       // документ  спецификация         +
		public const int D3_DOCUMENT_OBJ        = 131;       // 3d документ  модели или сборки +
		public const int CHANGE_LIST_OBJ        = 132;       // таблица изменений              -
		public const int TXT_DOCUMENT_OBJ       = 133;       // текстовый документ
		public const int ALL_DOCUMENTS          = 134;       // документы всех типов

		public const int MAX_TIP_SEARCH         = 134;       // верхняя граница типов поиска   -
    public const int ALL_OBJ_SHOW_ORDER     = -1000;     // все объекты которые могут входить в вид в порядке отрисовки


    // Поле тип линии имеет значения( системные стили ) см ksCurveStyleEnum:
		//	1  - основная,
		//  2  - тонкая,
		//  3  - осевая,
		//  4  - штриховая,
		//  5  - волнистая линия
		//	6  - вспомогательная,
		//  7  - утолщенная,
		//  8  - штрих-пунктир с 2 точками,
		//  9  - штриховая толстая
		//  10 -осевая толстая
		//  11 -тонкая, включаемая в штриховку
		//  12 - ISO штриховая линия
		//  13 - ISO штриховая линия (дл. пробел)
		//  14 - ISO штрихпунктирная линия (дл. штрих)
		//  15 - ISO штрихпунктирная линия (дл. штрих 2 пунктира)
		//  16 - ISO штрихпунктирная линия (дл. штрих 3 пунктира)
		//  17 - ISO пунктирная линия
		//  18 - ISO штрихпунктирная линия (дл. и кор. штрихи)
		//  19 - ISO штрихпунктирная линия (дл. и 2 кор. штриха) 
		//  20 - ISO штрихпунктирная линия
		//  21 - ISO штрихпунктирная линия (2 штриха)
		//  22 - ISO штрихпунктирная линия (2 пунктира)
		//  23 - ISO штрихпунктирная линия (3пунктира)
		//  24 - ISO штрихпунктирная линия (2 штриха 2 пунктира)
		//  25 - ISO штрихпунктирная линия (2 штриха 3 пунктира)

		// Поле тип точки для точки( системные стили ) :
		//0 - точка
		//1 - крестик
		//2 - х-точка
		//3	-	квадрат
		//4	-	треугольник
		//5	-	окружность
		//6	-	звезда
		//7	-	перечеркнутый квадрат
		//8	-	утолщенный плюс

    // Поле тип штриховки для штриховки( системные стили ) см ksHatchStyleEnum:
		// 0  - металл
		// 1  - неметалл 
		// 2  - дерево
		// 3  - камень естественный
		// 4  - керамика
		// 5  - бетон
		// 6  - стекло
		// 7  - жидкость
		// 8  - естественный грунт
		// 9  - насыпной грунт
		// 10 - камень искусственный
		// 11 - железобетон
		// 12 - напряженный железобетон
		// 13 - дерево в продольном сечении
		// 14 - песок

		// Определения флагов для текста
		public const int INVARIABLE			= 0;    //не менять флаги текста

		public const int NUMERATOR			= 0x1;    //числитель
		public const int DENOMINATOR		= 0x2;    //знаменатель
		public const int END_FRACTION       = 0x3;    //конец дроби
		public const int UPPER_DEVIAT       = 0x4;    //верхнее отклонение
		public const int LOWER_DEVIAT       = 0x5;    //нижнее отклонение
		public const int END_DEVIAT         = 0x6;    //конец  отклонений
		public const int S_BASE             = 0x7;    //основание выражения типа суммы
		public const int S_UPPER_INDEX      = 0x8;    //верхний индекс выражения типа суммы
		public const int S_LOWER_INDEX      = 0x9;    //нижний индекс выражения типа суммы
		public const int S_END              = 0x10;   //конец выражения типа суммы
		public const int SPECIAL_SYMBOL     = 0x11;   //спецзнак
		public const int SPECIAL_SYMBOL_END = 0x12;   //для спецзнаков с текстом
		public const int RETURN_BEGIN       = 0x13;   //начало для ввода следующих строк в спецзнаке с текстом, дробях, отклонениях
		public const int RETURN_DOWN        = 0x14;   //для ввода следующих строк в спецзнаке с текстом, дробях, отклонениях
		public const int RETURN_RIGHT       = 0x15;   //для ввода строк справа в спецзнаке с текстом, дробях, отклонениях
		public const int TAB                = 0x16;   //табуляция по текущему стилю
		public const int FONT_SYMBOL        = 0x17;   //символ фонта
    public const int MARK_SEPARATOR     = 0x18;   //разделитель в обозначении
    public const int FONT_SYMBOL_W      = 0x2017; //символ фонта Unicode
    public const int HYPER_TEXT         = 0x2000; //ссылка на текст или положение объекта

		public const int ITALIC_ON      = 0x40;   //включить наклон
		public const int ITALIC_OFF     = 0x80;   //выключть наклон
		public const int BOLD_ON        = 0x100;  //включить толщину
		public const int BOLD_OFF       = 0x200;  //выключть толщину
		public const int UNDERLINE_ON   = 0x400;  //включить подчеркивание
		public const int UNDERLINE_OFF  = 0x800;  //выключть подчеркивание
		public const int NEW_LINE       = 0x1000; //новая строка в параграфе

		public const int FONT_NAME       = 1;       //имя фонта
		public const int NARROWING       = 2;       //коэффициент сужения фонта
		public const int HEIGHT          = 3;       //высота фонта
		public const int COLOR           = 4;       //цвет текста
		public const int SPECIAL         = 5;       //спецзнак
		public const int FRACTION_TYPE   = 6;       //высота дроби по отношению к тексту 1-полная высота 2-в 1.5 раза меньше 3-в 2 раза меньше
		public const int SUM_TYPE        = 7;       //высота выражения типа суммы по отношению к тексту 1-полная высота 2-в 1.5 раза больше

		//Определения для динамических массивов
		public const int CHAR_STR_ARR          = 1;  // динамический массив указателей на строки символов
		public const int POINT_ARR             = 2;  // динамический массив указателей на математические точки -структура MathPointParam
		public const int CURVE_PATTERN_ARR     = 2;  // динамический массив указателей на участки штриховой линии -структура CurvePattern
		public const int TEXT_LINE_ARR         = 3;  // динамический массив строк текста - структура TextLineParam
		public const int TEXT_ITEM_ARR         = 4;  // динамический массив компонент строк текста структура TextItemParam
		public const int ATTR_COLUMN_ARR       = 5;  // динамический массив колонок атрибутов- структура  ColumnInfo
		public const int USER_ARR              = 6;  // динамический пользовательский массив
		public const int POLYLINE_ARR          = 7;  // динамический массив полилиний-(указателей массивов POINT_ARR)
		public const int RECT_ARR              = 8;  // динамический массив габаритных прямоугольников-(структура RectParam)
		public const int LIBRARY_STYLE_ARR     = 9;  // динамический массив структур параметров для стиля в библиотеке стилей( LibraryStyleParam )
		public const int VARIABLE_ARR          = 10; // динамический массив структур параметров параметрических переменных( VariableParam )
		public const int CURVE_PATTERN_ARR_EX  = 11; // динамический массив указателей на участки штриховой линии -структура CurvePatternEx
		public const int LIBRARY_ATTR_TYPE_ARR = 12; // динамический массив структур параметров для типа атрибута в библиотеке типов атрибутов( LibraryAttrTypeParam )
		public const int NURBS_POINT_ARR       = 13; // динамический массив структур NurbsPointParam
		public const int DOUBLE_ARR            = 14; // динамический массив duuble
		public const int CONSTRAINT_ARR        = 15; // динамический массив параметрических ограничений - структура ConstraintParam
		public const int CORNER_ARR            = 16; // динамический массив структур параметров углов CornerParam для прямоугольников и многоугольников
		public const int DOC_SPCOBJ_ARR        = 17; // динамический массив структур параметров прикрепленных документов к объекту спецификации DocAttachedSpcParam
		public const int SPCSUBSECTION_ARR     = 18; // динамический массив структур параметров подраздела спецификации SpcSubSectionParam
		public const int SPCTUNINGSEC_ARR      = 19; // динамический массив структур параметров настройки раздела спецификации SpcTuningSectionParam
		public const int SPCSTYLECOLUMN_ARR    = 20; // динамический массив структур параметров стиля колонки таблицы спецификации SpcStyleColumnParam
		public const int SPCSTYLESEC_ARR       = 21; // динамический массив структур параметров стиля разделa спецификации SpcStyleSectionParam
		public const int QUALITYITEM_ARR       = 22; // динамический массив структур QualityItemParam - запись об одном интервале для какого-то квалитета
		public const int LTVARIANT_ARR         = 23; // динамический массив структур LtVariant
		public const int TOLERANCEBRANCH_ARR   = 24; // динамический массив структур ToleranceBranch
		public const int HATCHLINE_ARR         = 25; // динамический массив структур HatchLineParam
		public const int TREENODEPARAM_ARR     = 26; // динамический массив структур узла дерева TreeNodeParam

    // Поле style для текста( системные стили ) см  ksTextStyleEnum:
    // 0 -умолчательный стиль для данного типа объекта
    // 1  текст на чертеже
    // 2  текст для технических требований
    // 3  текст размерных надписей
    // 4  текст шероховатости
    // 5  текст для линии выноски  ( позиционной )
    // 6  текст для линии выноски  ( над\под полкой )
    // 7  текст для линии выноски  ( сбоку )
    // 8  текст для допуска формы
    // 9  текст для таблицы ( заголовок )
    // 10 текст для таблицы ( ячейка )
    // 11 текст для линии разреза
    // 12 текст для стрелки вида
    // 13 текст для для неуказанной шероховатости
    // 14 текст для обозначения изменения
    // 15 текст для фигурной скобки
    // 16 текст для номера узла
    // 17 текст для выносной надписи
    // 18 текст для обозначения узла
    // 19 текст для марки координационной оси
    // 20 текст для МПО(марка/позиционное обозначение с линией-выноской)
    // 21 текст для МПО(марка/позиционное обозначение) на линии
    // 22 текст для МПО(марка/позиционное обозначение) без линии выноски
    // 23 текст для заголовков спецификации
    // 24 текст для линия разреза для СПДС
    // 25 Текст для таблицы отчета ( заголовок ).
    // 26 Текст для таблицы отчета ( ячейка ).

		// Структуры для работы с табличными атрибутами  
		public const int CHAR_ATTR_TYPE    = 1;
		public const int UCHAR_ATTR_TYPE   = 2;
		public const int INT_ATTR_TYPE     = 3;
		public const int UINT_ATTR_TYPE    = 4;
		public const int LINT_ATTR_TYPE    = 5;
		public const int FLOAT_ATTR_TYPE   = 6;
		public const int DOUBLE_ATTR_TYPE  = 7;
		public const int STRING_ATTR_TYPE  = 8;   //строка фиксированной длины MAX_TEXT_LENGTH
		public const int RECORD_ATTR_TYPE  = 9;

		// признаки формирования размерной надписи
		public const int _AUTONOMINAL       = 0x1;   // >0 автоматическое определение номинального значения размера
		public const int _RECTTEXT          = 0x2;   // >0 текст в рамочке
		public const int _PREFIX            = 0x4;   // >0 есть текст до номинала
		public const int _NOMINALOFF        = 0x8;   // >0 нет  номинала
		public const int _TOLERANCE         = 0x10;  // >0 квалитет проставлять
		public const int _DEVIATION         = 0x20;  // >0 отклонения проставлять
		public const int _UNIT              = 0x40;  // >0 единица измерения
		public const int _SUFFIX            = 0x80;  // >0 есть текст после номинала
		public const int _DEVIATION_INFORM  = 0x100; // >0 при включенном _DEVIATION, отклонения есть в массиве текстов( даже если не ручное проставление отклонений).
		//    Появляется после  функции GetObjParam, чтобы пользователь мог получить отклонения.
		public const int _UNDER_LINE        = 0x200; // >0 размер с подчеркиванием
		public const int _BRACKETS          = 0x400; // >0 размер в скобках
		public const int _SQUARE_BRACKETS   = 0x800; // >0 размер в квадратных скобках, используется вместе с _BRACKETS
		//    _BRACKETS                    - размер в круглых скобках
		//    _BRACKETS | _SQUARE_BRACKETS - размер в квадратных скобках

		public const int   INDICATIN_TEXT_LINE_ARR        = 0xFFFF;  //для шероховаиости, позиционной линии выноски, маркировки и клеймения
		//признак, что для текста используется динамический массив TEXT_LINE_ARR

		// типы стилей
		public const int CURVE_STYLE    = 1;  // стиль криивых
		public const int HATCH_STYLE    = 2;  // стиль штриховок
		public const int TEXT_STYLE     = 3;  // стиль текста
		public const int STAMP_STYLE    = 4;  // стиль штампа
		public const int CURVE_STYLE_EX = 5;  // стиль криивых расширенный

		// curveType | LIKE_BASIC_LINE - параметры пера как у  основной линии
		public const int  LIKE_BASIC_LINE = 0x10; // параметры пера как у  основной линии
		public const int  LIKE_THIN_LINE  = 0x20; // параметры пера как у  тонкой линии
		public const int  LIKE_HEAVY_LINE = 0x30; // параметры пера как у  утолщенной линии

		// Определения для функций Get/SetDocOptions и ksGet/SetSysOptions
		public const int DIMENTION_OPTIONS            = 1; // Настройки размера
		public const int SNAP_OPTIONS                 = 1; // Настройки привязок
		public const int ARROWFILLING_OPTIONS         = 2; // Зачернять стрелки ?
		public const int SHEET_OPTIONS                = 3; // Параметры листа для новых документов
		public const int SHEET_OPTIONS_EX             = 4; // Параметры листа документа
		public const int LENGTHUNITS_OPTIONS          = 5; // Настройки единиц измерений
		public const int SNAP_OPTIONS_EX              = 6; // Настройки привязок документа
		public const int VIEWCOLOR_OPTIONS            = 7; // Настройки цвета фона рабочего поля 2d - документов
		public const int TEXTEDIT_VIEWCOLOR_OPTIONS   = 8; // Настройки цвета фона редактирования текста
		public const int MODEL_VIEWCOLOR_OPTIONS      = 9; // Настройки цвета фона для моделей
		public const int OVERLAP_OBJECT_OPTIONS      = 10; // Настройки перекрывающихся объектов
		public const int DIMENTION_OPTIONS_EX        = 11; // Настройки размера

		//типы колонок для спецификации
		public const int   SPC_CLM_FORMAT   = 1;   // формат
		public const int   SPC_CLM_ZONE     = 2;   // зона
		public const int   SPC_CLM_POS      = 3;   // позиция
		public const int   SPC_CLM_MARK     = 4;   // обозначение
		public const int   SPC_CLM_NAME     = 5;   // наименование
		public const int   SPC_CLM_COUNT    = 6;   // количество
		public const int   SPC_CLM_NOTE     = 7;   // примечание
		public const int   SPC_CLM_MASSA    = 8;   // масса
		public const int   SPC_CLM_MATERIAL = 9;   // материал
		public const int   SPC_CLM_USER     = 10;  // пользовательская
		public const int   SPC_CLM_KOD      = 11;  // код
		public const int   SPC_CLM_FACTORY  = 12;  // завод изготовитель

		//типы значений для колонки спецификации
// констаны не используются нужно использовать LtVariantType
//		public const int   SPC_INT      = 1;   // целый
//		public const int   SPC_DOUBLE   = 2;   // вещественный
//		public const int   SPC_STRING   = 3;   // строка
//		public const int   SPC_RECORD   = 4;   // запись

		//типы блиотек стилей
		public const int CURVE_STYLE_LIBRARY               = 1; // библиотека стилей кривых (*.lcs)
		public const int HATCH_STYLE_LIBRARY               = 2; // библиотека стилей штриховок (*.lhs)
		public const int TEXT_STYLE_LIBRARY                = 3; // библиотека стилей текстов   (*.lts)
		public const int STAMP_LAYOUT_STYLE_LIBRARY        = 4; // библиотека стилей описаний штампов (*.lyt)
		public const int GRAPHIC_LAYOUT_STYLE_LIBRARY      = 5; // библиотека стилей оформлений графических документов (*.lyt)
		public const int TEXT_LAYOUT_STYLE_LIBRARY         = 6; // библиотека стилей оформлений текстовых документов (*.lyt)
		public const int SPC_LAYOUT_STYLE_LIBRARY          = 7; // библиотека стилей оформлений спецификаций (*.lyt)

		//размерности и типы детали для рассчета массо-ценровочных характеристик
		public const int  ST_MIX_MM      = 0x1;  // миллиметры
		public const int  ST_MIX_SM      = 0;    // сантиметры
		public const int  ST_MIX_DM      = 0x2;  // дециметры
		public const int  ST_MIX_M       = 0x3;  // метры
		public const int  ST_MIX_GR      = 0;    // граммы
		public const int  ST_MIX_KG      = 0x10; // килограммы
		public const int  ST_MIX_EXT     = 0;    // выдавливание
		public const int  ST_MIX_RV      = 0x20; // вращение

		// тип локальной привязки
		public const int  SN_NEAREST_POINT    = 1;    // Ближайшая точка
		public const int  SN_NEAREST_MIDDLE   = 2;    // Середина
		public const int  SN_CENTRE           = 3;    // Центр
		public const int  SN_INTERSECT        = 4;    // Пересечение
		public const int  SN_GRID             = 5;    // По сетке
		public const int  SN_XY_ALIGN         = 6;    // Выравнивание
		public const int  SN_ANGLE            = 7;    // Угловая привязка
		public const int  SN_POINT_CURVE      = 8;    // Точка на кривой

		// типы общих настроек для привязок
		public const int  SN_DYNAMICALLY               = 0x1;  // Привязки отслеживать динамически
		public const int  SN_ASSISTANT                 = 0x2;  // Писать текст
		public const int  SN_BACKGROUND_LAYER          = 0x4;  // Учитывать фоновые слои и виды
		public const int  SN_SUSPENDED                 = 0x8;  // Подавить привязки
		public const int  SN_VISIBLE_GRID_POINTS_ONLY  = 0x10; // Привязка только к видимым точкам сетки


		// Типы параметрических ограничений
		public const int CONSTRAINT_FIXED_POINT           = 1;  // фиксировать точку
		public const int CONSTRAINT_POINT_ON_CURVE        = 2;  // точка на кривой
		public const int CONSTRAINT_HORIZONTAL            = 3;  // горизонталь
		public const int CONSTRAINT_VERTICAL              = 4;  // вертикаль
		public const int CONSTRAINT_PARALLEL              = 5;  // параллельность двух прямых или отрезков
		public const int CONSTRAINT_PERPENDICULAR         = 6;  // перпендикулярность двух прямых или отрезков
		public const int CONSTRAINT_EQUAL_LENGTH          = 7;  // равенство длин двух отрезков
		public const int CONSTRAINT_EQUAL_RADIUS          = 8;  // равенство радиусов двух дуг/окружностей
		public const int CONSTRAINT_HOR_ALIGN_POINTS      = 9;  // выравнивать две точки по горизонтали
		public const int CONSTRAINT_VER_ALIGN_POINTS      = 10; // выравнивать две точки по вертикали
		public const int CONSTRAINT_MERGE_POINTS          = 11; // совпадение двух точек
		public const int CONSTRAINT_TANGENT_TWO_CURVES    = 15; // касание двух кривых
    public const int CONSTRAINT_SYMMETRY_TWO_POINTS   = 16; // симетрия двух точек
    public const int CONSTRAINT_COLLINEAR             = 17; // колинеарность отрезков
    public const int CONSTRAINT_FIXED_ANGLE           = 18; // фиксированный угол
    public const int CONSTRAINT_FIXED_LENGHT          = 19; // фиксированная длина
    public const int CONSTRAINT_POINT_ON_CURVE_MIDDLE = 20; // точка на серидине кривой
    public const int CONSTRAINT_BISECTOR              = 21; // биссектриса
    public const int CONSTRAINT_CONCENTRICITY         = 22; // совпадение центров окружностей, дуг, эллипсов и точек

		// типы объектов спецификации
		public const int  SPC_BASE_OBJECT     = 1;    // базовый объект (редактируется пользователем)
		public const int  SPC_COMMENT         = 2;    // комментарий    (редактируется пользователем)
		public const int  SPC_SECTION_NAME    = 3;    // наименование раздела ( создается по стилю СП автоматически )
		public const int  SPC_BLOCK_NAME      = 4;    // наименование блока исполнений ( создается по стилю СП автоматически )
		public const int  SPC_RESERVE_STR     = 5;    // резервная строка ( создается по стилю СП автоматически )
		public const int  SPC_EMPTY_STR       = 6;    // пустая строка ( создается по стилю СП автоматически )

		// типы сортировки
		public const int SPC_SORT_OFF        = 0;   // нет сортировки
		public const int SPC_SORT_COMPOS     = 1;   // составная сортировка
		public const int SPC_SORT_ALPHABET   = 2;   // сортировка по алфавиту (1.06.01 - больше не используется)
		public const int SPC_SORT_UP         = 3;   // сортировка по возрастанию колонок
		public const int SPC_SORT_DOCUMENT   = 4;   // сортировка раздела документация
		public const int SPC_SORT_DOWN       = 5;   // сортировка по убыванию колонок
		public const int SPC_SORT_COMPOSDOWN = 6;   // составная сортировка по убыванию 

		////////////////////////////////////////////////////////////////////////////////
		//
		//  типы специальных символов ( аннотационный объект )
		//
		////////////////////////////////////////////////////////////////////////////////
		public const int  ARROW_INSIDE_SYMBOL       = 1;  // стрелка (ласточкин хвост) изнутри
		public const int  ARROW_OUT_SIDE_SYMBOL     = 2;  // стрелка (ласточкин хвост) снаружи
		public const int  TICK_TAIL_SYMBOL          = 3;  // засечка с продолжением кривой (с хвостиком)
		public const int  UP_HALF_ARROW_SYMBOL      = 4;  // верхняя половина стрелки изнутри
		public const int  DOWN_HALF_ARROW_SYMBOL    = 5;  // нижняя половина стрелки изнутри
		public const int  BIG_ARROW_INSIDE_SYMBOL   = 6;  // большая стрелка изнутри (7мм)
		public const int  ARROW_ORDINATE_DIM_SYMBOL = 7;  // стрелка для размера высоты(штрихи длиной 4 мм под углом 45 гр)
		public const int  TRIANGLE_SYMBOL           = 8;  // треугольник по напр-нию кривой
		public const int  CIRCLE_RAD2_SYMBOL        = 9;  // окружность радиусом 2 мм тонкой линией - для шер-сти и линии-выноски
		public const int  CENTRE_MARKER_SYMBOL      = 10; // обозначение фиктивного центра в виде большого креста
		public const int  GLUE_SIGN_SYMBOL          = 11; // знак склеивания
		public const int  SOLDER_SIGN_SYMBOL        = 12; // знак пайки
		public const int  SEWING_SIGN_SYMBOL        = 13; // знак сшивания
		public const int  CRAMP_SIGN_SYMBOL         = 14; // знак соединения внахлестку металл.скобами
		public const int  CORNER_CRAMP_SIGN_SYMBOL  = 15; // знак углового соединения металл.скобами
		public const int  MONTAGE_JOINT_SYMBOL      = 16; // знак монтажного шва
		public const int  TICK_SYMBOL               = 17; // засечка без продолжения кривой (без хвостика)
		public const int  TRIANGLE_CURR_CS          = 18; // треугольник по текущей СК - для базы
		public const int  ARROW_CLOSED_INSIDE       = 19; // закрытая стрелка изнутри
		public const int  ARROW_CLOSED_OUTSIDE      = 20; // закрытая стрелка снаружи
		public const int  ARROW_OPEN_INSIDE         = 21; // открытая стрелка изнутри
		public const int  ARROW_OPEN_OUTSIDE        = 22; // открытая стрелка снаружи
		public const int  ARROW_RIGHTANGLE_INSIDE   = 23; // стрелка 90 град изнутри
		public const int  ARROW_RIGHTANGLE_OUTSIDE  = 24; // стрелка 90 град снаружи
		public const int  SYMBOL_DOT                = 25; // точка (диаметр равен длины стрелки размера)
		public const int  SYMBOL_SMALLDOT           = 26; // точка маленькая (диаметр равен 0.6 длины стрелки размера)
    public const int  AUXILIARY_POINT           = 27; // вспомогательная точка
    public const int  LEFT_TICK_SYMBOL          = 28; // засечка с наклоном влево
          

		//------------------------------------------------------------------------------
		// Битовые флаги для функции ksSetMacroParam;
		// указание какой тип редактирования поддерживает макро
		// ---
		public const int MP_DBL_CLICK_OFF  = 0x01; //>0 редактирование по двойному нажанию выключено
		public const int MP_HOTPOINTS      = 0x02; //>0 интерфейс hot точек включен
		public const int MP_EXTERN_EDIT    = 0x04; //>0 интерфейс внешнего управления

		//-----------------------------------------------------------------------------
		//определения для конвертации в растровый формат
		// ---
		public const int FORMAT_BMP   = 0;
		public const int FORMAT_GIF   = 1;
		public const int FORMAT_JPG   = 2;
		public const int FORMAT_PNG   = 3;
		public const int FORMAT_TIF   = 4;
		public const int FORMAT_TGA   = 5;
		public const int FORMAT_PCX   = 6;
    public const int FORMAT_WMF   = 16;
    public const int FORMAT_EMF   = 17;

		//-----------------------------------------------------------------------------
		//определения для настройки цвета растрового формата
		// ---
		public const int BLACKWHITE   = 0;   //цвет черный
		public const int COLORVIEW    = 1;   //цвет установленный для вида
		public const int COLORLAYER   = 2;   //цвет установленный для слоя
		public const int COLOROBJECT  = 3;   //цвет установленный для объекта

		//-----------------------------------------------------------------------------
		// орределения бит на пиксел для конвертации в растровый формат
		// ---
		public const int BPP_COLOR_01 = 1;  //"Черный" 
		public const int BPP_COLOR_02 = 2;  //"4 цвета" 
		public const int BPP_COLOR_04 = 4;  //"16 цветов" 
		public const int BPP_COLOR_08 = 8;  //"256 цветов"
		public const int BPP_COLOR_16 = 16; //"16 разрядов"
		public const int BPP_COLOR_24 = 24; //"24 разряда"
		public const int BPP_COLOR_32 = 32; //"32 разряда"

		//------------------------------------------------------------------------------
		// Типы стандартных видов
		// ---
		public const int VIEW_FRONT       = 0x1; //  Спереди
		public const int VIEW_REAR        = 0x2; //  Сзади
		public const int VIEW_UP          = 0x4; //  Сверху
		public const int VIEW_DOWN        = 0x8; //  Снизу
		public const int VIEW_LEFT        = 0x10; //  Слева
		public const int VIEW_RIGHT       = 0x20; //  Справа
		public const int VIEW_ISO         = 0x40; //  Изометрия

    //------------------------------------------------------------------------------
    // Стандартые курсоры Компас
    // ---
    public const int OCR_SELECT  = 0xFFFE; //  Курсор ввиде SELECT 
    public const int OCR_SNAP    = 0xFFFD; //  Курсор ввиде SNAP 
    public const int OCR_CATCH   = 0xFFFC; //  Курсор ввиде CATCH
    public const int OCR_DEFAULT = 0;      //  Курсор в виде креста

    public const int OCR_DEDAULT = 0;      //  Курсор в виде креста

    //-----------------------------------------------------------------------------
    // Неопределенный цвет для TextItemFont.color
    // В стиле может быть выставлен по умолчанию цвет отличный он 0
    // В этом случае если TextItemFont.color будет значение 0 то создастся
    // модификатор на цвет и он не будет отображаться как цвет по умолчанию
    // для того чтобы модификатор цвета не создался нужно
    // или прислать цвет из настроек или константу FREE_COLOR
    // ---
    public const uint FREE_COLOR     = 0xff000000; //  Неопределенный цвет

	}
}