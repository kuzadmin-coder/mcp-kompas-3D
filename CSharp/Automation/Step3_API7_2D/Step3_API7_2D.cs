using Kompas6API5;
using KompasAPI7;
using Kompas6Constants;
using KAPITypes;
using System;
using System.Resources;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Windows.Forms;


namespace Step3_API7_2D
{
	public class SamplesSymbols
	{
    private KompasObject      kompas;
    private IApplication      appl;         // Интерфейс приложения
    private IKompasDocument2D doc;          // Интерфейс документа 2D в API7 
    private ksDocument2D      doc2D;        // Интерфейс документа 2D в API5
    private ResourceManager   resMng = new ResourceManager( typeof(SamplesSymbols) );       // Менеджер ресурсов


    //-------------------------------------------------------------------------------
    // Загрузить строку из ресурса
    // ---
    string LoadString( string name )
    {
      return resMng.GetString( name );
    }
    
    //-------------------------------------------------------------------------------
    // Имя библиотеки
    // ---
    [return: MarshalAs(UnmanagedType.BStr)] public string GetLibraryName()
    {
      return LoadString( "IDS_LIBNAME" );
    }

    //-------------------------------------------------------------------------------
    // Меню библиотеки
    // ---
    [return: MarshalAs(UnmanagedType.BStr)] public string ExternalMenuItem(short number, ref short itemType, ref short command)
    {
      string result = string.Empty;
      itemType = 1; // "MENUITEM"
      switch (number)
      {
        case 1:
          result = "Создание и редактирование простой линии выноски";
          command = 1;
        break;

        case 2:
          result = "Создание и редактирование знака маркировки";
          command = 2;
        break;

        case 3:
          result = "Создание и редактирование знака изменения";
          command = 3;
        break;

        case 4:
          result = "Создание и редактирование знака клеймения";
          command = 4;
        break;

        case 5:
          result = "Создание и редактирование обозначения шероховатости";
          command = 5;
        break;

        case 6:
          result = "Создание и редактирование обозначения базы";
          command = 6;
        break;

        case 7:
          result = "Создание и редактирование линии разреза/сечения";
          command = 7;
        break;

        case 8:
          result = "Создание и редактирование стрелки направления взгляда";
          command = 8;
        break;

        case 9:
          result = "Создание и редактирование допуска формы";
          command = 9;
        break;

        case 10:
          result = "Навигация по массиву объектов вида";
          command = 10;
        break;

        case 11:
          command = -1;
          itemType = 3; // "ENDMENU"
        break;
      }
      return result;
    }


    //-------------------------------------------------------------------------------
    // Головная функция библиотеки
    // ---
    public void ExternalRunCommand([In] short command, [In] short mode, [In, MarshalAs(UnmanagedType.IDispatch)] object kompas_)
    {
      kompas = (KompasObject)kompas_;

      if ( kompas == null )
        return;
      
      // Получаем интерфейс приложения
      appl = (IApplication)kompas.ksGetApplication7();

      if ( appl == null )
        return;

      // Получаем интерфейс активного документа 2D в API7
      doc = (IKompasDocument2D)appl.ActiveDocument;

      if ( doc == null )
        return;

      // Получаем интерфейс активного документа 2D в API5
      doc2D = (ksDocument2D)kompas.ActiveDocument2D();

      if ( doc2D == null )
        return;

      switch ( command )
      {
        case 1: LeaderWork();       break;  // Создание и редактирование простой линии выноски
        case 2: MarkLeaderWork();   break;  // Создание и редактирование знака маркировки
        case 3: ChangeLeaderWork(); break;  // Создание и редактирование знака изменения
        case 4: BrandLeaderWork();  break;  // Создание и редактирование знака клеймения
        case 5: RoughWork();        break;  // Создание и редактирование обозначения шероховатости
        case 6: BaseWork();         break;  // Создание и редактирование обозначения базы
        case 7: CutLineWork();      break;  // Создание и редактирование линии разреза/сечения
        case 8: ViewPointerWork();  break;  // Создание и редактирование стрелки направления взгляда
        case 9: ToleranceWork();    break;  // Создание и редактирование допуска формы
        case 10: ObjectsNavigation(); break;  // Навигация по массиву объектов вида
      }
    }

    #region Вспомогательные функции
    //-------------------------------------------------------------------------------
    //  Получить контейнер обозначений 2D
    // ---
    ISymbols2DContainer GetSymbols2DContainer()
    {
      if ( doc != null )
      {
        // Получим менеджер для работы с видами и слоями
        ViewsAndLayersManager viewsMng = doc.ViewsAndLayersManager;
    		
        if ( viewsMng != null )
        {
          // Получим коллекцию видов
          IViews views = viewsMng.Views;
    			
          if ( views != null )
          {
            // Получаем контейнер у активного вида
            IView view = views.ActiveView;

            if ( view != null )
              return (ISymbols2DContainer)view;
          }
        }
      }
      return null;
    }


    //-------------------------------------------------------------------------------
    // Создать текст с заданным размером шрифта
    // ---
    void SetTextSmallFont( ref IText txt, string str, double size )
    {
	    if ( txt != null )
	    {
		    // Получить интерфейс строки текста
		    ITextLine line = txt.Add();

		    if ( line != null )
		    {
			    // Получить интерфейс компоненты текста
			    ITextItem item = line.Add();

			    if ( item != null )
			    {
				    // Получить интерфейс шрифта
				    ITextFont font = (ITextFont)item;

				    if ( font != null )
					    // Задать высоту шрифта
					    font.Height = size;

				    // Задать значение строки
				    item.Str = str;
				    // Применить параметры
				    item.Update();
			    }
		    }
	    }
    }
    #endregion


    #region Простая линия выноски
    //-------------------------------------------------------------------------------
    // Создание простой линии выноски
    // ---
    bool CreateLeader( ref ILeader leader )
    {
      bool res = false;

	    if ( leader != null )
	    {
		    // Направление полки - вправо
		    leader.ShelfDirection = Kompas6Constants.ksShelfDirectionEnum.ksLSRight;

		    // Получаем интерфейс ответвлений
		    IBranchs branchs = (IBranchs)leader;

		    if ( branchs != null )
		    {
			    // Координаты начала полки или точка привязки
			    branchs.X0 = 100;
			    branchs.Y0 = 150;
			    // Добавить прямолинейные ответвления
			    branchs.AddBranchByPoint( -1, 60, 120 );
			    branchs.AddBranchByPoint( -1, 65, 105 );
		    }

		    // Получить интерфейс текста над полкой
		    IText txtOnSh = leader.TextOnShelf;

		    if ( txtOnSh != null )
			    // Изменить текст
			    txtOnSh.Str = "1";

		    // Получить интерфейс текста под полкой
		    IText txtUnderSh = leader.TextUnderShelf;

		    if ( txtUnderSh != null )
			    // Изменить текст
			    txtUnderSh.Str = "2";

		    // Получить интерфейс текста над ножкой
		    IText txtOnBr = leader.TextOnBranch;
    		
		    if ( txtOnBr != null )
			    // Изменить текст
			    txtOnBr.Str = "3";

		    // Получить интерфейс текста под ножкой
		    IText txtUnderBr = leader.TextUnderBranch;
    		
		    if ( txtUnderBr != null )
			    // Изменить текст
			    txtUnderBr.Str = "4";

		    // Получить интерфейс текста за полкой
		    IText txtAfterSh = leader.TextAfterShelf;
    		
		    if ( txtAfterSh != null )
			    // Изменить текст
			    txtAfterSh.Str = "5";

		    IBaseLeader baseLeader = (IBaseLeader)leader;

		    if ( baseLeader != null )
			    // Применить параметры
			    res = baseLeader.Update();
	    }
      return res;
    }


    //-------------------------------------------------------------------------------
    // Задать компоненту текста
    // ---
    void AddTextItem( ref ITextLine line, string str, Kompas6Constants.ksTextItemEnum type )
    {
	    if ( line != null )
	    {
		    // Получить интерфейс компоненты текста
		    ITextItem item = line.Add();
    		
		    if ( item != null )
		    {
			    // Задать строковое значение
			    item.Str = str;
			    // Задать тип
			    item.ItemType = type;
			    // Применить параметры
			    item.Update();
		    }
	    }
    }


    //-------------------------------------------------------------------------------
    // Редактирование линии выноски
    // ---
    void EditLeader( ref ILeader leader )
    {
	    if ( leader != null )
	    {
		    // Тип значка - знак склеивания
		    leader.SignType = Kompas6Constants.ksLeaderSignEnum.ksLGlueSign;
		    // Включить признак обоработки по контуру
		    leader.Arround = true;

		    // Получаем интерфейс ответвлений
		    IBranchs branchs = (IBranchs)leader;
    		
		    if ( branchs != null )
			    // Добавить прямолинейное ответвление
			    branchs.AddBranchByPoint( -1, 140, 120 );

		    // Начало созданного ответвления - от конца полки
		    leader.set_BranchBegin( 2, false );

		    // Получить текст над полкой
		    IText txt = leader.TextOnShelf;

		    if ( txt != null )
		    {
			    txt.Str = "";

			    ITextLine line = txt.Add();

			    if ( line != null )
			    {
				    // Добавить строку
				    AddTextItem( ref line, "1", Kompas6Constants.ksTextItemEnum.ksTItString );
				    // Добавить числитель дроби
				    AddTextItem( ref line, "1", Kompas6Constants.ksTextItemEnum.ksTItNumerator );
				    // Добавить знаментатель дроби
				    AddTextItem( ref line, "1", Kompas6Constants.ksTextItemEnum.ksTItDenominator );
				    // Закончить дробь
				    AddTextItem( ref line, "1", Kompas6Constants.ksTextItemEnum.ksTItFractionEnd );
			    }
		    }

		    IBaseLeader baseLeader = (IBaseLeader)leader;
    		
		    if ( baseLeader != null )
			    // Применить параметры
			    baseLeader.Update();
	    }
    }


    //-------------------------------------------------------------------------------
    // Создание и редактирование простой линии выноски
    // ---
    void LeaderWork()
    {
	    // Получить контейнер условных обозначений
	    ISymbols2DContainer symbCont = GetSymbols2DContainer();

	    if ( symbCont != null )
	    {
		    // Получить коллекцию линий выноски
		    ILeaders leadersCol = symbCont.Leaders;

		    if ( leadersCol != null )
		    {
			    // Добавить простую линию выноски
			    ILeader leader = (ILeader)leadersCol.Add( Kompas6Constants.DrawingObjectTypeEnum.
                                           ksDrLeader );

			    if ( leader != null && doc2D != null )
			    {
				    // Создать линию выноски
				    bool create = CreateLeader( ref leader );

				    // Получить референс объекта
				    IBaseLeader bLeader = (IBaseLeader)leader;
				    int refr = 0;

				    if ( bLeader != null )
					    refr = bLeader.Reference; 
    				
				    // Подсветить созданный объект
            doc2D.ksLightObj( refr, 1/*включить подсветку*/ );

            if ( create && kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
				    {		
					    // Получить объект из коллекции по референсу
					    ILeader lead = (ILeader)leadersCol.get_Leader( refr );				
					    // Редактировать линию выноски
					    EditLeader( ref lead );
				    }
            doc2D.ksLightObj( refr, 0/*выключить подсветку*/ );
			    }
		    }
	    }
    }
    #endregion


    #region Знак маркировки
    //-------------------------------------------------------------------------------
    // Создание знака маркировки
    // ---
    bool CreateMarkLeader( ref IMarkLeader markLeader )
    {
      bool res = false;

	    if ( markLeader != null )
	    {
		    // Получить интерфейс ответвлений
		    IBranchs branchs = (IBranchs)markLeader;
    		
		    if ( branchs != null )
		    {
			    // Точка привязки 
			    branchs.X0 = 100;
			    branchs.Y0 = 190;
			    // Добавить прямолинейное ответвление
			    branchs.AddBranchByPoint( -1, 60, 120 );
		    }

		    // Получить интерфейс текста обозначения
		    IText des = markLeader.Designation;

		    if ( des != null )
			    // Изменить текст
			    des.Str = LoadString( "IDS_MARK" );

		    // Получить интерфейс текста над ножкой
		    IText textOnBranch = markLeader.TextOnBranch;

		    if ( textOnBranch != null )
			    // Изменить текст
			    textOnBranch.Str = "2";

		    // Получить интерфейс текста под ножкой
		    IText textUnderBranch = markLeader.TextUnderBranch;
    		
		    if ( textUnderBranch != null )
			    // Изменить текст
			    textUnderBranch.Str = "3";

		    // Получить базовый интерфейс линии выноски
		    IBaseLeader baseLeader = (IBaseLeader)markLeader;

		    if ( baseLeader != null )
			    // Применить параметры
			    res = baseLeader.Update();
	    }
      return res;
    }


    //-------------------------------------------------------------------------------
    // Редактирование знака маркировки
    // ---
    void EditMarkLeader( ref IMarkLeader markLeader )
    {
	    if ( markLeader != null )
	    {
		    // Получить интерфейс ответвлений
		    IBranchs branchs = (IBranchs)markLeader;

		    if ( branchs != null )
			    // Добавить прямолинейное ответвление
			    branchs.AddBranchByPoint( -1, 70, 110 );

		    // Получить базовый интерфейс линии выноски
		    IBaseLeader baseLeader = (IBaseLeader)markLeader;

		    if ( baseLeader != null )
		    {
			    // Тип стрелки
			    baseLeader.ArrowType = Kompas6Constants.ksArrowEnum.ksLeaderArrow;
			    // Применить изменения
			    baseLeader.Update();
		    }
	    }
    }


    //-------------------------------------------------------------------------------
    // Создание и редактирование знака маркировки
    // ---
    void MarkLeaderWork()
    {
	    // Получить контейнер условных обозначений
	    ISymbols2DContainer symbCont = GetSymbols2DContainer();
    	
	    if ( symbCont != null )
	    {
		    // Получить коллекцию линий выноски
		    ILeaders leadersCol = symbCont.Leaders;
    		
		    if ( leadersCol != null )
		    {
			    // Добавить знак маркировки
			    IMarkLeader mLeader = (IMarkLeader)leadersCol.Add( Kompas6Constants.
                                 DrawingObjectTypeEnum.ksDrMarkerLeader);
    			
			    if ( mLeader != null && doc2D != null )
			    {
				    // Создать знак маркировки
				    bool create = CreateMarkLeader( ref mLeader );
    				
				    // Получить референс объекта
				    IBaseLeader bLeader = (IBaseLeader)mLeader;
				    int refr = 0;
    				
				    if ( bLeader != null )
					    refr = bLeader.Reference;
    				
				    // Подсветить созданный объект
            doc2D.ksLightObj( refr, 1/*включить подсветку*/ );
    				
            if ( create && kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
				    {
					    // Получить объект из коллекции по референсу
					    IMarkLeader markLead = (IMarkLeader)leadersCol.get_Leader( refr );				
					    // Редактировать знак маркировки
					    EditMarkLeader( ref markLead );
				    }
            doc2D.ksLightObj( refr, 0/*выключить подсветку*/ );
			    }
		    }
	    }
    }
    #endregion


    #region Знак изменения
    //-------------------------------------------------------------------------------
    // Создание знака изменения
    // ---
    bool CreateChangeLeader( ref IChangeLeader changeLeader )
    {
      bool res = false;

	    if ( changeLeader != null )
	    {
		    // Тип значка - квадрат
		    changeLeader.SignType = Kompas6Constants.ksChangeLeaderSignEnum.ksCLSSquare;

		    // Получить интерфейс ответвлений
		    IBranchs branchs = (IBranchs)changeLeader;

		    if ( branchs != null )
		    {
			    // Точка привязки
			    branchs.X0 = 70;
			    branchs.Y0 = 150;
			    // Добавить ответвление
			    branchs.AddBranchByPoint( -1, 40, 130 );
		    }

		    // Получить интерфейс текста обозначения
		    IText des = changeLeader.Designation;
    		
		    if ( des != null )
			    // Изменить текст
			    des.Str = "1";

		    // Получить базовый интерфейс линии выноски
		    IBaseLeader baseLeader = (IBaseLeader)changeLeader;

		    if ( baseLeader != null )
			    // Применить параметры
			    res = baseLeader.Update();
	    }
      return res;
    }


    //-------------------------------------------------------------------------------
    // Редактирование знака изменения
    // ---
    void EditChangeLeader( ref IChangeLeader changeLeader )
    {
	    if ( changeLeader != null )
	    {
		    // Тип значка - круг
		    changeLeader.SignType = Kompas6Constants.ksChangeLeaderSignEnum.ksCLSCircle;
		    // Полная длина выноски отключена
		    changeLeader.FullLeaderLength = false;
		    // Длина выноски
		    changeLeader.LeaderLength = 5;

		    // Получить базовый интерфейс линии выноски
		    IBaseLeader baseLeader = (IBaseLeader)changeLeader;
    		
		    if ( baseLeader != null )
			    // Применить изменения
			    baseLeader.Update();
	    }
    }

    //-------------------------------------------------------------------------------
    // Создание и редактирование знака изменения
    // ---
    void ChangeLeaderWork()
    {
	    // Получить контейнер условных обозначений
	    ISymbols2DContainer symbCont = GetSymbols2DContainer();
    	
	    if ( symbCont != null )
	    {
		    // Получить коллекцию линий выноски
		    ILeaders leadersCol = symbCont.Leaders;
    		
		    if ( leadersCol != null )
		    {
			    // Добавить знак изменения
			    IChangeLeader chLeader = (IChangeLeader)leadersCol.Add( Kompas6Constants.
                                    DrawingObjectTypeEnum.ksDrChangeLeader );
    			
			    if ( chLeader != null && doc2D != null )
			    {
				    // Создать знак изменения
				    bool create = CreateChangeLeader( ref chLeader );
    				
				    // Получить референс объекта
				    IBaseLeader bLeader = (IBaseLeader)chLeader;
				    int refr = 0;
    				
				    if ( bLeader != null )
					    refr = bLeader.Reference;
    				
				    // Подсветить созданный объект
            doc2D.ksLightObj( refr, 1/*включить подсветку*/ );
    				
            if ( create && kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
				    {
					    // Получить объект из коллекции по референсу
					    IChangeLeader changeLead = (IChangeLeader)leadersCol.get_Leader( refr );				
					    // Редактировать знак изменения
					    EditChangeLeader( ref changeLead );
				    }
            doc2D.ksLightObj( refr, 0/*выключить подсветку*/ );
			    }
		    }
	    }
    }
    #endregion


    #region Знак клеймения
    //-------------------------------------------------------------------------------
    // Создание знака клеймения
    // ---
    bool CreateBrandLeader( ref IBrandLeader brandLeader )
    {
      bool res = false;

	    if ( brandLeader != null )
	    {
		    // Направление
		    brandLeader.Direction = false;

		    // Получить интерфейс текста обозначения
		    IText des = brandLeader.Designation;

		    if ( des != null )
			    // Изменить текст
			    des.Str = LoadString( "IDS_MARK2" );

		    // Получить интерфейс ответвлений
		    IBranchs branchs = (IBranchs)brandLeader;

		    if ( branchs != null )
		    {
			    // Точка привязки
			    branchs.X0 = 100;
			    branchs.Y0 = 150;
			    // Добавить ответвление
			    branchs.AddBranchByPoint( -1, 60, 110 );
		    }

		    // Получить базовый интерфейс линии выноски
		    IBaseLeader baseLeader = (IBaseLeader)brandLeader;
    		
		    if ( baseLeader != null )
			    // Применить параметры
			    res = baseLeader.Update();
	    }
      return res;
    }


    //-------------------------------------------------------------------------------
    // Редактирование знака изменения
    // ---
    void EditBrandLeader( ref IBrandLeader brandLeader )
    {
	    if ( brandLeader != null )
	    {
		    // Направление
		    brandLeader.Direction = true;

		    // Получить интерфейс текста над ножкой
		    IText textOn = brandLeader.TextOnBranch;
    		
		    if ( textOn != null )
			    // Изменить текст
			    textOn.Str = "2";
    		
		    // Получить интерфейс текста под ножкой
		    IText textUnder = brandLeader.TextUnderBranch;
    		
		    if ( textUnder != null )
			    // Изменить текст
			    textUnder.Str = "3";

		    // Получить интерфейс текста обозначения
		    IText des = brandLeader.Designation;
    		
		    if ( des != null )
			    // Изменить текст
			    des.Str = LoadString( "IDS_MARK3" );

		    // Получить базовый интерфейс линии выноски
		    IBaseLeader baseLeader = (IBaseLeader)brandLeader;
    		
		    if ( baseLeader != null )
			    // Применить параметры
			    baseLeader.Update();
	    }
    }


    //-------------------------------------------------------------------------------
    // Создание и редактирование знака клеймения
    // ---
    void BrandLeaderWork()
    {
	    // Получить контейнер условных обозначений
	    ISymbols2DContainer symbCont = GetSymbols2DContainer();
    	
	    if ( symbCont != null )
	    {
		    // Получить коллекцию линий выноски
		    ILeaders leadersCol = symbCont.Leaders;
    		
		    if ( leadersCol != null )
		    {
			    // Добавить знак клеймения
			    IBrandLeader brLeader = (IBrandLeader)leadersCol.Add( Kompas6Constants.DrawingObjectTypeEnum.ksDrBrandLeader );
    			
			    if ( brLeader != null && doc2D != null )
			    {
				    // Создать знак клеймения
				    bool create = CreateBrandLeader( ref brLeader );
    				
				    // Получить референс объекта
				    IBaseLeader bLeader = (IBaseLeader)brLeader;
				    int refObj = 0;
    				
				    if ( bLeader != null )
					    refObj = bLeader.Reference;
    				
				    // Подсветить созданный объект
            doc2D.ksLightObj( refObj, 1/*включить подсветку*/ );
    				
            if ( create && kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
				    {
					    // Получить объект из коллекции по референсу
					    IBrandLeader brandLead = (IBrandLeader)leadersCol.get_Leader( refObj );				
					    // Редактировать знак клеймения
					    EditBrandLeader( ref brandLead );
				    }
            doc2D.ksLightObj( refObj, 0/*выключить подсветку*/ );
			    }
		    }
	    }
    }
    #endregion


    #region Обозначение шероховатости
    //-------------------------------------------------------------------------------
    // Создание обозначения шероховатости
    // ---
    bool CreateRough( ref IRough rough )
    {
	    bool res = false;

	    if ( rough != null && doc2D != null )
	    {
		    // Структура параметров запроса к системе
		    ksRequestInfo info = (ksRequestInfo)kompas.GetParamStruct( (int)Kompas6Constants.
                              StructType2DEnum.ko_RequestInfo );
		    info.Init();
		    info.commandsString = LoadString( "IDS_COMMAND1" );
        info.cursorId = ldefin2d.OCR_CATCH;
		    double x = 0, y = 0;

		    // Отключить привязки
		    SnapOptions sOpt = (SnapOptions)kompas.GetParamStruct( (int)Kompas6Constants.
                            StructType2DEnum.ko_SnapOptions );
        kompas.ksGetSysOptions( ldefin2d.SNAP_OPTIONS, sOpt );
		    sOpt.commonOpt = 0;
		    kompas.ksSetSysOptions( ldefin2d.SNAP_OPTIONS, sOpt );

		    // Указать базовый объект
        if ( doc2D.ksCursorEx(info, ref x, ref y, null, null) != 0 )
		    {
			    // Найти объект по указанным координатам
          int pObj = doc2D.ksFindObj( x, y, doc2D.ksGetCursorLimit() );
			    // Подсветить объект
          doc2D.ksLightObj( pObj, 1/*включить подсветку*/ );
			    // Получить интерфейс графического объекта из референса
			    IDrawingObject baseObj = (IDrawingObject)kompas.TransferReference( pObj, doc2D.reference );
			    info.commandsString = LoadString( "IDS_COMMAND2" );
			    info.cursorId = 0;

			    // Если привязки выключены, включить
			    if ( sOpt.commonOpt == 0 )
			    {
				    sOpt.commonOpt = 1;
				    kompas.ksSetSysOptions( ldefin2d.SNAP_OPTIONS, sOpt );
			    }

			    // Указать точку привязки
          if ( doc2D.ksCursorEx(info, ref x, ref y, null, null) != 0 )
			    {
				    if ( baseObj != null )
					    // Базовый объект
					    rough.BaseObject = baseObj;

				    // Точка привязки
				    rough.BranchX0 = x;
				    rough.BranchY0 = y;
    				
				    IRoughParams roughPar = (IRoughParams)rough;

				    if ( roughPar != null )
				    {
					    roughPar.ShelfDirection = Kompas6Constants.ksShelfDirectionEnum.ksLSRight;
					    rough.Update();
					    // Обработка по контуру включена
					    roughPar.ProcessingByContour = true;
					    // Длина линии выноски
					    roughPar.LeaderLength = 20;
					    // Угол наклона линии выноски
					    roughPar.LeaderAngle = 45;

					    // Получить интерфейс текста параметров шероховатости
					    IText txt1 = roughPar.RoughParamText;

					    if ( txt1 != null )
						    txt1.Str = "1";
    					
					    // Получить интерфейс текста способа обработки поверхности
					    IText txt2 = roughPar.ProcessText;
    					
					    if ( txt2 != null )
						    txt2.Str = "2";

					    // Получить интерфейс текста базовой длины
					    IText txt3 = roughPar.BaseLengthText;
    					
					    if ( txt3 != null )
						    txt3.Str = "3";
				    }

				    // Применить параметры
				    res = rough.Update();
			    }
			    else
            kompas.ksMessage( LoadString("IDS_NOPOINT") );

			    doc2D.ksLightObj( pObj, 0/*выключить подсветку*/ );
		    }
		    else
          kompas.ksMessage( LoadString("IDS_NOOBJ") );

		    // Если привязки выключены, включить
		    if ( sOpt.commonOpt == 0 )
		    {
			    sOpt.commonOpt = 1;
			    kompas.ksSetSysOptions( ldefin2d.SNAP_OPTIONS, sOpt );
		    }
	    }
	    return res;
    }


    //-------------------------------------------------------------------------------
    // Редактирование обозначения шероховатости
    // ---
    void EditRough( ref IRough rough )
    {
	    if ( rough != null )
	    {
		    // Получить интерфейс параметров шероховатости
		    IRoughParams roughPar = (IRoughParams)rough;

		    if ( roughPar != null )
		    {
			    // Убрать стрелку
			    roughPar.ArrowType = Kompas6Constants.ksArrowEnum.ksWithoutArrow;
			    // Обработка по контуру отключена
			    roughPar.ProcessingByContour = false;
		    }
		    // Применить изменения
		    rough.Update();
	    }
    }


    //-------------------------------------------------------------------------------
    // Создание и редактирование обозначения шероховатости
    // ---
    void RoughWork()
    {
	    // Получить контейнер условных обозначений
	    ISymbols2DContainer symbCont = GetSymbols2DContainer();
    	
	    if ( symbCont != null )
	    {
		    // Получить коллекцию обозначений шероховатости
		    IRoughs roughsCol = symbCont.Roughs;
    		
		    if ( roughsCol != null )
		    {
			    // Добавить обозначение шероховатости
			    IRough newRough = roughsCol.Add();
    			
			    if ( newRough != null && doc2D != null )
			    {
				    // Создать обозначение шероховатости
				    if ( CreateRough(ref newRough) )
				    {			
					    // Получить референс объекта
					    int objRef = newRough.Reference;
					    // Подсветить созданный объект
              doc2D.ksLightObj( objRef, 1/*включить подсветку*/ );
    					
              if ( kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
					    {
						    // Получить объект из коллекции по референсу
						    IRough rough = roughsCol.get_Rough( objRef );				
						    // Редактировать обозначение шероховатости
						    EditRough( ref rough );
					    }
              doc2D.ksLightObj( objRef, 0/*выключить подсветку*/ );
				    }
			    }
		    }
	    }
    }
    #endregion


    #region Обозначение базы
    //-------------------------------------------------------------------------------
    // Создание обозначения базы
    // ---
    bool CreateBase( ref IBase pBase )
    {
	    bool res = false;

	    if ( pBase != null )
	    {
        // Структура параметров запроса к системе
        ksRequestInfo info = (ksRequestInfo)kompas.GetParamStruct( (int)Kompas6Constants.
                              StructType2DEnum.ko_RequestInfo );
        info.Init();
        info.commandsString = LoadString( "IDS_COMMAND1" );
        info.cursorId = ldefin2d.OCR_CATCH;
        double x = 0, y = 0;

        // Отключить привязки
        SnapOptions sOpt = (SnapOptions)kompas.GetParamStruct( (int)Kompas6Constants.
                            StructType2DEnum.ko_SnapOptions );
        kompas.ksGetSysOptions( ldefin2d.SNAP_OPTIONS, sOpt );
        sOpt.commonOpt = 0;
        kompas.ksSetSysOptions( ldefin2d.SNAP_OPTIONS, sOpt );
    		
		    // Указать базовый объект
        if ( doc2D.ksCursorEx(info, ref x, ref y, null, null) != 0 )
		    {
			    // Найти объект по указанным координатам
          int pObj = doc2D.ksFindObj( x, y, doc2D.ksGetCursorLimit() );
			    // Подсветить объект
          doc2D.ksLightObj( pObj, 1/*включить подсветку*/ );
			    // Получить интерфейс графического объекта из референса
			    IDrawingObject baseObj = (IDrawingObject)kompas.TransferReference( pObj, doc2D.reference );
			    info.commandsString = LoadString( "IDS_COMMAND2" );
			    info.cursorId = 0;

			    // Если привязки выключены, включить
			    if ( sOpt.commonOpt == 0 )
			    {
				    sOpt.commonOpt = 1;
				    kompas.ksSetSysOptions( ldefin2d.SNAP_OPTIONS, sOpt );
			    }
    			
			    // Указать положение знака
          if ( doc2D.ksCursorEx(info, ref x, ref y, null, null) != 0 )
			    {
				    if ( baseObj != null )
					    // Базовый объект
					    pBase.BaseObject = baseObj;

				    // Точка положения знака
				    pBase.X0 = x;
				    pBase.Y0 = y;
				    // Способ отрисовки - перпендикулярно к объекту
				    pBase.DrawType = true;
				    // Применить параметры
				    res = pBase.Update();
			    }
			    else
            kompas.ksMessage( LoadString("IDS_NOPOINT") );

          doc2D.ksLightObj( pObj, 0/*выключить подсветку*/ );
		    }
		    else
          kompas.ksMessage( LoadString("IDS_NOOBJ") );

		    // Если привязки выключены, включить
		    if ( sOpt.commonOpt == 0 )
		    {
			    sOpt.commonOpt = 1;
			    kompas.ksSetSysOptions( ldefin2d.SNAP_OPTIONS, sOpt );
		    }
	    }
	    return res;
    }


    //-------------------------------------------------------------------------------
    // Редактирование обозначения базы
    // ---
    void EditBase( ref IBase pBase )
    {
	    if ( pBase != null )
	    {
		    // Тип отрисовки - произвольно
		    pBase.DrawType = false;
		    // Отключить автосортировку
		    pBase.AutoSorted = false;
		    // Конечная точка выноски
		    pBase.BranchX = pBase.X0 + 10;
		    pBase.BranchY = pBase.Y0 + 10;

		    // Получить интерфейс текста обозначения базы
		    IText txt = pBase.Text;

		    if ( txt != null )
			    // Изменить текст
			    txt.Str = "B";

		    // Применить параметры
		    pBase.Update();
	    }
    }


    //-------------------------------------------------------------------------------
    // Создание и редактирование обозначения базы
    // ---
    void BaseWork()
    {
	    // Получить контейнер условных обозначений
	    ISymbols2DContainer symbCont = GetSymbols2DContainer();
    	
	    if ( symbCont != null )
	    {
		    // Получить коллекцию обозначений базы
		    IBases basesCol = symbCont.Bases;
    		
		    if ( basesCol != null )
		    {
			    // Добавить обозначение базы
			    IBase newBase = basesCol.Add();
    			
			    if ( newBase != null )
			    {
				    // Создать обозначение базы
				    if ( CreateBase( ref newBase) )
				    {			
					    // Получить референс объекта
					    int objRef = newBase.Reference;
					    // Подсветить созданный объект
              doc2D.ksLightObj( objRef, 1/*включить подсветку*/ );
    					
              if ( kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
					    {
						    // Получить объект из коллекции по референсу
						    IBase pBase = basesCol.get_Base( objRef );				
						    // Редактировать обозначение базы
						    EditBase( ref pBase );
					    }
              doc2D.ksLightObj( objRef, 0/*выключить подсветку*/ );
				    }
			    }
		    }
	    }
    }
    #endregion


    #region Линия разреза/сечения
    //-------------------------------------------------------------------------------
    // Создание линии разреза/сечения
    // ---
    bool CreateCutLine( ref ICutLine cutLine )
    {
      bool res = false;
 
	    if ( cutLine != null )
	    {
		    // Координаты начального текста
		    cutLine.X1 = 80;
		    cutLine.Y1 = 165;
		    // Координаты конечного текста
		    cutLine.X2 = 120;
		    cutLine.Y2 = 200;
		    // Расположение стрелок - слева
		    cutLine.ArrowPos = true;
		    // Размещение дополнительного текста - у первой стрелки
		    cutLine.AdditionalTextPos = true;

		    // Создать массив точек линии разреза
		    double[] points = new double[4];
		    points[0] = 80;
		    points[1] = 165;
		    points[2] = 120;
		    points[3] = 200;
			  // Задать массив точек линии разреза
			  cutLine.Points = points;

		    // Получить интерфейс дополнительного текста
		    IText adText = cutLine.AdditionalText;
    		
		    if ( adText != null )
			    SetTextSmallFont( ref adText, "(1)", 7 );

		    // Применить параметры
		    res = cutLine.Update();
	    }
      return res;
    }


    //-------------------------------------------------------------------------------
    // Редактирование линии разреза/сечения
    // ---
    void EditCutLine( ref ICutLine cutLine )
    {
	    if ( cutLine != null )
	    {
		    // Расположение стрелок - справа по направлению ломаной
		    cutLine.ArrowPos = false;
		    // Расположение дополнительного текста - у второй стрелки
		    cutLine.AdditionalTextPos = false;
		    // Отключить автосортировку
		    cutLine.AutoSorted = false;

		    // Создать массив точек линии разреза
		    double[] points = new double[6];
		    points[0] = 80;
		    points[1] = 165;
		    points[2] = 115;
		    points[3] = 165;
		    points[4] = 120;
		    points[5] = 200;
			  cutLine.Points = points;

		    // Применить изменения
		    cutLine.Update();
	    }
    }


    //-------------------------------------------------------------------------------
    // Создание и редактирование линии разреза/сечения
    // ---
    void CutLineWork()
    {
	    // Получить контейнер условных обозначений
	    ISymbols2DContainer symbCont = GetSymbols2DContainer();
    	
	    if ( symbCont != null )
	    {
		    // Получить коллекцию линий разреза/сечения
		    ICutLines cutCol = symbCont.CutLines;
    		
		    if ( cutCol != null && doc2D != null )
		    {
			    // Добавить линию разреза/сечения
			    ICutLine newCut = cutCol.Add();
    			
			    if ( newCut != null )
			    {
				    // Создать линию разреза/сечения
				    bool create = CreateCutLine( ref newCut );
    				
				    // Получить референс объекта
				    int objRef = newCut.Reference;
				    // Подсветить созданный объект
            doc2D.ksLightObj( objRef, 1/*включить подсветку*/ );
    				
            if ( create && kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
				    {
					    // Получить объект из коллекции по референсу
					    ICutLine cutLine = cutCol.get_CutLine( objRef );				
					    // Редактировать линию разреза/сечения
					    EditCutLine( ref cutLine );
				    }
            doc2D.ksLightObj( objRef, 0/*выключить подсветку*/ );
			    }
		    }
	    }
    }
    #endregion


    #region Стрелка направления взгляда
    //-------------------------------------------------------------------------------
    // Создание стрелки направления взгляда
    // ---
    bool CreateViewPointer( ref IViewPointer viewPointer )
    {
      bool res = false;

	    if ( viewPointer != null )
	    {
		    // Начальная точка стрелки
		    viewPointer.X1 = 100;
		    viewPointer.Y1 = 150;
		    // Конечная точка стрелки
		    viewPointer.X2 = 120;
		    viewPointer.Y2 = 160;
		    // Применить параметры
		    res = viewPointer.Update();
	    }
      return res;
    }


    //-------------------------------------------------------------------------------
    // Редактирование стрелки направления взгляда
    // ---
    void EditViewPointer( ref IViewPointer viewPointer )
    {
	    if ( viewPointer != null )
	    {
		    // Конечная точка стрелки
		    viewPointer.X2 = 90;
		    viewPointer.Y2 = 140;
		    // Включить автосортировку
		    viewPointer.AutoSorted = true;

		    // Получить интерфейс текста обозначения направления взгляда
		    IText adText = viewPointer.AdditionalText;
    		
		    if ( adText != null )
			    // Изменить текст
			    SetTextSmallFont( ref adText, "(1)", 7 );
    		
		    // Применить параметры
		    viewPointer.Update();
	    }
    }


    //-------------------------------------------------------------------------------
    // Создание и редактирование стрелки направления взгляда
    // ---
    void ViewPointerWork()
    {
	    // Получить контейнер условных обозначений
	    ISymbols2DContainer symbCont = GetSymbols2DContainer();
    	
	    if ( symbCont != null )
	    {
		    // Получить коллекцию стрелок направления взгляда
		    IViewPointers viewPointsCol = symbCont.ViewPointers;
    		
		    if ( viewPointsCol != null )
		    {
			    // Добавить стрелку направления взгляда
			    IViewPointer newPointer = viewPointsCol.Add();
    			
			    if ( newPointer != null && doc2D != null )
			    {
				    // Создать стрелку направления взгляда
				    bool create = CreateViewPointer( ref newPointer );
    				
				    // Получить референс объекта
				    int objRef = newPointer.Reference;
				    // Подсветить созданный объект
            doc2D.ksLightObj( objRef, 1/*включить подсветку*/ );
    				
            if ( create && kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
				    {
					    // Получить объект из коллекции по референсу
					    IViewPointer viewPointer = viewPointsCol.get_ViewPointer( objRef );				
					    // Редактировать стрелку направления взгляда
					    EditViewPointer( ref viewPointer );
				    }
            doc2D.ksLightObj( objRef, 0/*выключить подсветку*/ );
			    }
		    }
	    }
    }
    #endregion


    #region Допуск формы
    //-------------------------------------------------------------------------------
    // Задать текст в ячейках допуска формы
    // ---
    void SetToleranceText( ref IToleranceParam tolPar )
    {
	    if ( tolPar != null )
	    {
		    // Получить интерфейс таблицы с текстом допуска формы
		    ITable tolTable = tolPar.Table;
    		
		    if ( tolTable != null )
		    {
			    // Добавить 3 столбца (1 уже есть)
			    tolTable.AddColumn( -1, true/*справа*/ );
			    tolTable.AddColumn( -1, true/*справа*/ );
			    tolTable.AddColumn( -1, true/*справа*/ );
    			
			    // Записать текст в 1-ю ячейку
			    ITableCell cell = tolTable.get_Cell( 0, 0 );
			    ITextLine txt = null;
    			
			    if ( cell != null )
			    {
				    txt = (ITextLine)cell.Text;
    				
				    if ( txt != null )
					    txt.Str = "@22~";
			    }
    			
			    // Записать текст во 2-ю ячейку
			    cell = tolTable.get_Cell( 0, 1 );
    			
			    if ( cell != null )
			    {
				    txt = (ITextLine)cell.Text;
    				
				    if ( txt != null )
					    txt.Str = "@2~";
			    }

			    // Записать текст в 3-ю ячейку
			    cell = tolTable.get_Cell( 0, 2 );
    			
			    if ( cell != null )
			    {
				    txt = (ITextLine)cell.Text;
    				
				    if ( txt != null )
					    txt.Str = "B";
			    }

			    // Записать текст в 4-ю ячейку
			    cell = tolTable.get_Cell( 0, 3 );
    			
			    if ( cell != null )
			    {
				    txt = (ITextLine)cell.Text;
    				
				    if ( txt != null )
					    txt.Str = "@30~";
			    }
		    }
	    }
    }


    //-------------------------------------------------------------------------------
    // Создание допуска формы
    // ---
    bool CreateTolerance( ref ITolerance tolerance )
    {
      bool res = false;

	    if ( tolerance != null )
	    {
		    // Получить интерфейс ответвления
		    IBranchs branchs = (IBranchs)tolerance;

		    if ( branchs != null )
		    {
			    // Задать точку привязки
			    branchs.X0 = 100;
			    branchs.Y0 = 150;
			    // Добавить 2 ответвления
			    branchs.AddBranchByPoint( -1, 100, 120 );
			    branchs.AddBranchByPoint( -1, 50, 155 );
		    }

		    // Получить интерфейс параметров допуска формы
		    IToleranceParam tolPar = (IToleranceParam)tolerance;
    		
		    if ( tolPar != null )
		    {
			    // Создать текст в ячейках
			    SetToleranceText( ref tolPar );
			    // Положение базовой точки относительно таблицы - внизу посередине
			    tolPar.BasePointPos = Kompas6Constants.ksTablePointEnum.ksTPBottomCenter;
		    }
		    // Тип стрелки 1-го ответвления - треугольник
		    tolerance.set_ArrowType( 0, false );
		    // Положение 1-го ответвления относительно таблицы - внизу посередине
		    tolerance.set_BranchPos( 0, Kompas6Constants.ksTablePointEnum.ksTPBottomCenter );
		    // Тип стрелки 2-го ответвления - стрелка
		    tolerance.set_ArrowType( 1, true );
		    // Положение 2-го ответвления относительно таблицы - слева посередине
		    tolerance.set_BranchPos( 1, Kompas6Constants.ksTablePointEnum.ksTPLeftCenter );
		    // Применить параметры
		    res = tolerance.Update();
	    }
      return res;
    }


    //-------------------------------------------------------------------------------
    // Редактирование допуска формы
    // ---
    void EditTolerance( ref ITolerance tolerance )
    {	
	    if ( tolerance != null )
	    {
		    // Получитьинтерфейс параметров допуска формы
		    IToleranceParam tolPar = (IToleranceParam)tolerance;

		    if ( tolPar != null )
		    {
			    // Получить интерфейс таблицы с текстом допуска формы
			    ITable tolTable = tolPar.Table;
    			
			    if ( tolTable != null )
			    {
				    // Изменить текст во 2-й ячейке
				    ITableCell cell = tolTable.get_Cell( 0, 1 );
    				
				    if ( cell != null )
				    {
					    ITextLine txt = (ITextLine)cell.Text;
    					
					    if ( txt != null )
						    txt.Str = "@2~15";
				    }
			    }
			    // Задать признак вертикальности
			    tolPar.Vertical = true;
		    }

		    // Получить интерфейс ответвлений
		    IBranchs branchs = (IBranchs)tolerance;

		    if ( branchs != null )
		    {
			    // Удалить ответвление
			    branchs.DeleteBranch( 0 );
			    // Добавить новое ответвление
			    branchs.AddBranchByPoint( -1, 130, 120 );
		    }
		    tolerance.set_ArrowType( 1, false );
		    tolerance.set_BranchPos( 1, Kompas6Constants.ksTablePointEnum.ksTPBottomCenter );
		    // Применить изменения
		    tolerance.Update();
	    }
    }


    //-------------------------------------------------------------------------------
    // Создание и редактирование допуска формы
    // ---
    void ToleranceWork()
    {
	    // Получить контейнер условных обозначений
	    ISymbols2DContainer symbCont = GetSymbols2DContainer();
    	
	    if ( symbCont != null )
	    {
		    // Получить коллекцию допусков формы
		    ITolerances tolerancesCol = symbCont.Tolerances;
    		
		    if ( tolerancesCol != null )
		    {
			    // Добавить допуск формы
			    ITolerance newTol = tolerancesCol.Add();
    			
			    if ( newTol != null )
			    {
				    // Создать допуск формы
				    bool create = CreateTolerance( ref  newTol );
    				
				    // Получить референс объекта
				    int objRef = newTol.Reference;
				    // Подсветить созданный объект
            doc2D.ksLightObj( objRef, 1/*включить подсветку*/ );
    				
            if ( create && kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
				    {
					    // Получить объект из коллекции по референсу
					    ITolerance tolerance = tolerancesCol.get_Tolerance( objRef );				
					    // Редактировать допуск формы
					    EditTolerance( ref tolerance );
				    }
            doc2D.ksLightObj( objRef, 0/*выключить подсветку*/ );
			    }
		    }
	    }
    }
    #endregion


    #region Навигация по объектам вида
    //-------------------------------------------------------------------------------
    // Получить контейнер графических объектов
    // ---
    IDrawingContainer GetDrawingContainer()
    {
	    if ( doc != null )
	    {
		    // Получить менеджер видов и слоев
		    IViewsAndLayersManager mng = doc.ViewsAndLayersManager;

		    if ( mng != null )
		    {
			    // Получить коллекцию видов
			    IViews viewsCol = mng.Views;

			    if ( viewsCol != null )
			    {
				    // Получить активный вид
				    IView view = viewsCol.ActiveView;

				    if ( view != null )
				    {
					    // Получить контейнер графических объектов
					    IDrawingContainer drawCont = (IDrawingContainer)view;
					    return drawCont;
				    }
			    }
		    }
	    }
	    return null;
    }


    //-------------------------------------------------------------------------------
    // Выдать сообщение с параметрами простой линии выноски
    // ---
    void GetLeaderPar( ref ILeader leader )
    {
	    if ( leader != null )
	    {
		    double x0 = 0, y0 = 0;

		    // Получить координаты точки привязки
		    IBranchs branchs = (IBranchs)leader;

		    if ( branchs != null )
		    {
			    x0 = branchs.X0;
			    y0 = branchs.Y0;
		    }

		    // Сформировать сообщение для выдачи
		    // "Простая линия выноски"
		    string buf = LoadString( "IDS_LEADER" );
        buf += "\n\r" + LoadString( "IDS_POINT" ) + "\n\r";
        buf += string.Format( LoadString("IDS_COORDS"), x0, y0 );
		    // "\nПризнак обработки по контуру: "
        buf += "\n\r" + LoadString( "IDS_ARROUND" );

		    if ( leader.Arround )
			    // "включен"
          buf += LoadString( "IDS_ON" );
		    else
			    // "отключен"
          buf += LoadString( "IDS_OFF" );

		    // "\nНачало ответвления: "
        buf += "\n\r" + LoadString( "IDS_BRANCHBEGIN" );

		    if ( leader.get_BranchBegin(0) )
			    // "от начала полки"
          buf += LoadString( "IDS_BEGINSHELF" );
		    else
			    // "от конца полки"
          buf = buf + LoadString( "IDS_ENDSHELF" );

		    // "\nПризнак параллельности ответвлений: "
        buf += "\n\r" + LoadString( "IDS_PARALLEL" );

		    if ( leader.ParallelBranch )
			    // "включен"
          buf += LoadString( "IDS_ON" );
		    else
			    // "отключен"
          buf += LoadString( "IDS_OFF" );
    		
		    // Выдать сообщение 
		    kompas.ksMessage( buf );
	    }
    }


    //-------------------------------------------------------------------------------
    // Выдать сообщение с параметрами знака маркировки
    // ---
    void GetMarkLeaderPar( ref IMarkLeader markLeader )
    {
	    if ( markLeader != null )
	    {
		    double x0 = 0, y0 = 0;
		    long count = 0;
    	
		    IBranchs branchs = (IBranchs)markLeader;
    		
		    if ( branchs != null )
		    {
			    // Получить координаты точки привязки
			    x0 = branchs.X0;
			    y0 = branchs.Y0;
			    // Получить количество ответвлений
			    count = branchs.BranchCount;
		    }

		    // Сформировать сообщение для выдачи
		    // "Знак маркировки"
        string buf = LoadString( "IDS_MARKLEADER" );
		    // "\nТочка привязки:\nX0 = %4.2f, Y0 = %4.2f"
        buf += "\n\r" + LoadString( "IDS_POINT" ) + "\n\r";
        buf += string.Format( LoadString("IDS_COORDS"), x0, y0 );
		    // "\nКоличество ответвлений: %d"
        buf += "\n\r" + string.Format( LoadString("IDS_BRANCHCOUNT"), count );

		    // Выдать сообщение 
        kompas.ksMessage( buf );
	    }
    }


    //-------------------------------------------------------------------------------
    // Выдать сообщение с параметрами знака клеймения
    // ---
    void GetBrandLeaderPar( ref IBrandLeader brandLeader )
    {
	    if ( brandLeader != null )
	    {
		    // Сформировать сообщение
        string buf = LoadString( "IDS_BRANDLEADER" );
		    // Выдать сообщение 
		    kompas.ksMessage( buf );
	    }
    }


    //-------------------------------------------------------------------------------
    // Выдать сообщение с параметрами знака изменения
    // ---
    void GetChangeLeaderPar( ref IChangeLeader changeLeader )
    {
	    if ( changeLeader != null )
	    {
		    // Сформировать сообщение
        string buf = LoadString( "IDS_CHANLEADER" );
		    // Выдать сообщение 
		    kompas.ksMessage( buf );
	    }
    }


    //-------------------------------------------------------------------------------
    // Выдать сообщение с параметрами обозначения шероховатости
    // ---
    void GetRoughPar( ref IRough rough )
    {
	    if ( rough != null )
	    {
		    // Сформировать сообщение
        string buf = LoadString( "IDS_ROUGH" );
		    // Выдать сообщение 
		    kompas.ksMessage( buf );
	    }
    }


    //-------------------------------------------------------------------------------
    // Выдать сообщение с параметрами обозначения базы
    // ---
    void GetBasePar( ref IBase pBase )
    {
	    if ( pBase != null )
	    {
		    // Сформировать сообщение
        string buf = LoadString( "IDS_BASE" );
		    // Выдать сообщение 
		    kompas.ksMessage( buf );
	    }
    }


    //-------------------------------------------------------------------------------
    // Выдать сообщение с параметрами линии разреза/сечения
    // ---
    void GetCutLinePar( ref ICutLine cutLine )
    {
	    if ( cutLine != null )
	    {
		    // Сформировать сообщение
        string buf = LoadString( "IDS_CUTLINE" );
		    // Выдать сообщение 
		    kompas.ksMessage( buf );
	    }
    }


    //-------------------------------------------------------------------------------
    // Выдать сообщение с параметрами стрелки направления взгляда
    // ---
    void GetViewPointerPar( ref IViewPointer viewPointer )
    {
	    if ( viewPointer != null )
	    {
		    // Сформировать сообщение
        string buf = LoadString( "IDS_VIEWPOINTER" );
		    // Выдать сообщение 
		    kompas.ksMessage( buf );
	    }
    }


    //-------------------------------------------------------------------------------
    // Выдать сообщение с параметрами допуска формы
    // ---
    void GetTolerancePar( ref ITolerance tolerance )
    {
	    if ( tolerance != null )
	    {
		    // Сформировать сообщение
        string buf = LoadString( "IDS_TOLERANCE" );
		    // Выдать сообщение 
		    kompas.ksMessage( buf );
	    }
    }


    //-------------------------------------------------------------------------------
    // Навигация по массиву объектов вида
    // ---
    void ObjectsNavigation()
    {
	    // Получить контейнер графических объектов
	    IDrawingContainer drawCont = GetDrawingContainer();

	    if ( drawCont != null )
	    {
		    // Получить массив объектов 
        Array arr = (Array)drawCont.get_Objects( 0/*все объекты*/ );

        // Если массив есть
		    if ( arr != null )
		    {    	
			    for ( long j = 0; j < arr.Length; j++ )
			    {
				    // Получить элемент из массива
				    stdole.IDispatch pObj = (stdole.IDispatch)arr.GetValue( j );
    				
				    if ( pObj != null )
				    {
					    // Получить интерфейс графического объекта
					    IDrawingObject pDrawObj = (IDrawingObject)pObj;
    					
					    if ( pDrawObj != null )
					    {
						    // Получить тип объекта
						    long type = (int)pDrawObj.DrawingObjectType;
						    // Получить референс объекта
						    int objRef = pDrawObj.Reference;
						    // Подсветить объект
						    doc2D.ksLightObj( objRef, 1/*включить*/ );

						    // В зависимости от типа вывести сообщение для данного типа объектов
						    switch( type )
						    {
							    // Простая линия выноски
							    case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrLeader:
							    {
								    ILeader leader = (ILeader)pDrawObj;
								    GetLeaderPar( ref leader );
								    break;
							    }

							    // Знак маркировки
							    case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrMarkerLeader:
							    {
								    IMarkLeader markLeader = (IMarkLeader)pDrawObj;
								    GetMarkLeaderPar( ref markLeader );
								    break;
							    }

							    // Знак клеймения
							    case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrBrandLeader:
							    {
								    IBrandLeader brandLeader = (IBrandLeader)pDrawObj;
								    GetBrandLeaderPar( ref brandLeader );
								    break;
							    }

							    // Знак изменения
							    case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrChangeLeader:
							    {
								    IChangeLeader changeLeader = (IChangeLeader)pDrawObj;
								    GetChangeLeaderPar( ref changeLeader );
								    break;
							    }

							    // Обозначение шероховатости
							    case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrRough:
							    {
								    IRough rough = (IRough)pDrawObj;
								    GetRoughPar( ref rough );
								    break;
							    }

							    // Обозначение базы
							    case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrBase:
							    {
								    IBase pBase = (IBase)pDrawObj;
								    GetBasePar( ref pBase );
								    break;
							    }

							    // Линия разреза/сечения
							    case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrCut:
							    {
								    ICutLine cutLine = (ICutLine)pDrawObj;
								    GetCutLinePar( ref cutLine );
								    break;
							    }

							    // Стрелка направления взгляда
							    case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrWPointer:
							    {
								    IViewPointer viewPointer = (IViewPointer)pDrawObj;
								    GetViewPointerPar( ref viewPointer );
								    break;
							    }

							    // Допуск формы
							    case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrTolerance:
							    {
								    ITolerance tolerance = (ITolerance)pDrawObj;
								    GetTolerancePar( ref tolerance );
								    break;
							    }
						    }
						    // Убрать подсветку
						    doc2D.ksLightObj( objRef, 0/*выключить*/ );
					    } 
				    }
			    }
		    }
	    }
    }
    #endregion


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
