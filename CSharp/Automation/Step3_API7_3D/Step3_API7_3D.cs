using Kompas6API5;
using KompasAPI7;
using Kompas6Constants;
using Kompas6Constants3D;
using System;
using System.Resources;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Windows.Forms;


namespace Step3_API7_3D
{
	public class Symbols3D
	{
    private KompasObject      kompas;
    private IApplication      appl;         // Интерфейс приложения
    private IKompasDocument3D doc;          // Интерфейс документа 3D в API7 
    private ksDocument3D      doc3D;        // Интерфейс документа 3D в API5
    private ResourceManager   resMng = new ResourceManager( typeof(Symbols3D) );   // Менеджер ресурсов
    private int               oType = (int)Kompas6Constants3D.Obj3dType.o3d_face;


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
          result = LoadString( "IDS_MENU1" );
          command = 1;
        break;

        case 2:
          result = LoadString( "IDS_MENU2" );
          command = 2;
        break;

        case 3:
          result = LoadString( "IDS_MENU3" );
          command = 3;
        break;

        case 4:
          result = LoadString( "IDS_MENU4" );
          command = 4;
        break;

        case 5:
          result = LoadString( "IDS_MENU5" );
          command = 5;
        break;

        case 6:
          result = LoadString( "IDS_MENU6" );
          command = 6;
        break;

        case 7:
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

      // Получаем интерфейс активного документа 3D в API7
      doc = (IKompasDocument3D)appl.ActiveDocument;

      if ( doc == null )
        return;

      // Получаем интерфейс активного документа 3D в API5
      doc3D = (ksDocument3D)kompas.ActiveDocument3D();

      if ( doc3D == null )
        return;

      switch ( command )
      {
        case 1: Rough3DWork();        break;  // Создание и редактирование шероховатости 3D
        case 2: Base3DWork();         break;  // Создание и редактирование обозначения базы 3D
        case 3: Leader3DWork();       break;  // Создание и редактирование линии-выноски 3D
        case 4: BrandLeader3DWork();  break;  // Создание и редактирование знака клеймения 3D
        case 5: MarkLeader3DWork();   break;  // Создание и редактирование знака маркировки 3D
        case 6: Tolerance3DWork();    break;  // Создание и редактирование допуска формы 3D
      }
    }


    #region Вспомогательные функции
    //-------------------------------------------------------------------------------
    // Получить контейнер обозначений 3D
    // ---
    ISymbols3DContainer GetSymbols3DContainer()
    {
      if ( doc != null )
      {
        // Получаем контейнер обозначений 3D
        return (ISymbols3DContainer)doc.TopPart;
      }
      return null;
    }


    //-----------------------------------------------------------------------------
    // Функция фильтрации
    // ---
    public bool UserFilterProc( object e )
    {
      ksEntity entity = (ksEntity)e;

      if( e != null && (oType == 0 || entity.type == oType) )
      {
        return true;
      }
      else
        return false;
    }


    //-------------------------------------------------------------------------------
    // Задать плоскость обозначения по вершине базового объекта
    // ---
    void SetPosition( ref IRough3D rough, ref IModelObject face )
    {
	    if ( rough != null && face != null )
	    {
		    // Преобразовать интерфейс объекта дерева из API7 в API5
        ksEntity faceEnt = (ksEntity)kompas.TransferInterface( face, 
                           (int)Kompas6Constants.ksAPITypeEnum.ksAPI5Auto,
                           (int)Kompas6Constants3D.ksObj3dTypeEnum.o3d_entity );

		    if ( faceEnt != null )
		    { 
			    // Получить интерфейс параметров грани
			    ksFaceDefinition faceDef = (ksFaceDefinition)faceEnt.GetDefinition();

			    if ( faceDef != null )
			    {
				    // Получить коллекцию ребер
				    ksEdgeCollection edgeCol = (ksEdgeCollection)faceDef.EdgeCollection();

				    if ( edgeCol != null )
				    {
					    // Получить интерфейс параметров ребра
					    ksEdgeDefinition edgeDef = (ksEdgeDefinition)edgeCol.First();

					    if ( edgeDef != null )
					    {
						    // Получить интерфейс параметров вершины
						    ksVertexDefinition vertex = (ksVertexDefinition)edgeDef.GetVertex( true/*начальная вершина*/ );

						    if ( vertex != null )
						    {
							    // Преобразовать интерфейс объекта из API5 в API7
                  IModelObject mObj = (IModelObject)kompas.TransferInterface( vertex, 
                                      (int)Kompas6Constants.ksAPITypeEnum.ksAPI7Dual, 0 );

							    if ( mObj != null )
								    // Установить положение плоскости обозначения по найденной вершине
								    rough.PositionObject = mObj;
						    }
					    }
				    }
			    }
		    }
	    }
    }


    //-------------------------------------------------------------------------------
    // Получить новое имя объекта
    // ---
    string GetNewName( ref IFeature7 featObj, string strID )
    {
	    string res = "";

	    if ( featObj != null )
	    {
		    // Сформировать новое имя объекта в дереве
		    string newName =	LoadString( strID );
		    string name = featObj.Name;
        int pos = name.IndexOf( ":", 0, name.Length );
    		
		    if ( pos > 0 )
			    newName += name.Substring( pos + 1, name.Length - pos - 1 );

        res = newName;
	    }
	    return res;
    }


    //-------------------------------------------------------------------------------
    // 
    // ---
    Int32 RGB( Int32 r, Int32 g, Int32 b )
    {
      return ( (r | (g << 8)) | (b << 16) );
    }


    //-------------------------------------------------------------------------------
    // Задать компоненту текста
    // ---
    void AddTextItem( ref ITextLine line, string str, 
                      Kompas6Constants.ksTextItemEnum type )
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
    #endregion


    #region Шероховатость 3D
    //-------------------------------------------------------------------------------
    // Установить параметры шероховатости
    // ---
    void SetRoughPars( ref IRough3D rough )
    {
	    if ( rough != null )
	    {
		    // Получить интерфейс параметров шероховатости
		    IRoughParams roughPars = (IRoughParams)rough;
    		
		    if ( roughPars != null )
		    {
			    // Признак обработки по контуру
			    roughPars.ProcessingByContour = true;
			    // Длина линии выноски
			    roughPars.LeaderLength = 20;
			    // Угол наклона линии выноски
			    roughPars.LeaderAngle = 45;
    			
			    // Получить интерфейс текста параметров шероховатости
			    IText txt1 = roughPars.RoughParamText;
    			
			    if ( txt1 != null )
				    txt1.Str = "1";
    			
			    // Получить интерфейс текста способа обработки поверхности
			    IText txt2 = roughPars.ProcessText;
    			
			    if ( txt2 != null )
				    txt2.Str = "2";
    			
			    // Получить интерфейс текста базовой длины
			    IText txt3 = roughPars.BaseLengthText;
    			
			    if ( txt3 != null )
				    txt3.Str = "3";
		    }
	    }
    }


    //-------------------------------------------------------------------------------
    // Создать шероховатость 3D
    // ---
    bool CreateRough3D( ref IRough3D rough )
    {
	    bool res = false;

	    if ( rough != null && doc3D != null )
	    {
			  // Указать в документе опорный объект для построения шероховатости
        ksEntity baseObj = (ksEntity)doc3D.UserSelectEntity( null, "UserFilterProc", 
                                           LoadString("IDS_OBJ"), 0, this );

			  if ( baseObj != null )
			  {
          double x = 0, y = 0, z = 0;
				  // Указать точку привязки
          doc3D.UserGetCursor( LoadString("IDS_POINT"), out x, out y, out z );

				  // Преобразовать интерфейс объекта из API5 в API7
          IModelObject mObj = (IModelObject)kompas.TransferInterface( baseObj, 
                              (int)Kompas6Constants.ksAPITypeEnum.ksAPI7Dual, 0 );

				  if ( mObj != null )
					  // Установить привязку шероховатости
					  rough.SetBasePosition( x, y, z, mObj );

				  // Задать базовую плоскость
				  rough.BasePlane = Kompas6Constants3D.ksObj3dTypeEnum.o3d_planeXOZ;

				  // Установить параметры шероховатости
				  SetRoughPars( ref rough );

				  // Применить параметры
				  res = rough.Update();
			  }
		  }
	    return res;
    }
    

    //-------------------------------------------------------------------------------
    // Редактировать шероховатость 3D
    // ---
    void EditRough3D( ref IRough3D rough )
    {
	    if ( rough != null )
	    {
		    // Получить интерфейс параметров шероховатости
		    IRoughParams roughPars = (IRoughParams)rough;

		    if ( roughPars != null )
			    // Убрать признак обработки по контуру
			    roughPars.ProcessingByContour = false;

		    // Получить интерфейс объекта дерева
		    IFeature7 featObj = (IFeature7)rough;

		    if ( featObj != null )
		    {	
			    // Изменить имя объекта в дереве	
			    rough.Name = GetNewName( ref featObj, "IDS_ROUGH" ); 
			    // Получить базовый объект
			    IModelObject baseObj = rough.BaseObject;
			    // Задать плоскость обозначения по вершине базового объекта
			    SetPosition( ref rough, ref baseObj );
		    }

		    // Получить интерфейс свойств цвета объекта
		    IColorParam7 colorPars = (IColorParam7)rough;

		    // Изменить цвет на синий
		    if ( colorPars != null )
			    colorPars.Color = RGB( 0, 0, 255 );

		    // Применить изменения
		    rough.Update();
	    }
    }


    //-------------------------------------------------------------------------------
    // Создание и редактирование шероховатости 3D
    // ---
    void Rough3DWork()
    {
	    // Получить контейнер обозначений 3D
	    ISymbols3DContainer symbCont = GetSymbols3DContainer();
    	
	    if ( symbCont != null )
	    {
		    // Получить коллекцию шероховатостей 3D
		    IRoughs3D roughsCol = symbCont.Roughs3D;
    		
		    if ( roughsCol != null )
		    {
			    // Добавить новое обозначение шероховатости 3D
			    IRough3D newRough = roughsCol.Add();
    			
			    if ( newRough != null )
			    {
				    // Создать обозначение шероховатости 3D
				    if ( CreateRough3D(ref newRough) )
				    {	
					    string name = "";
					    // Получить интерфейс объекта дерева построения
					    IFeature7 feature = (IFeature7)newRough;
    					
					    // Получить имя объекта
					    if ( feature != null )
						    name = feature.Name;
    					
					    if ( kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
					    {
						    // Получить объект из коллекции по имени
						    IRough3D rough = roughsCol.get_Rough3D( name );
						    // Редактировать шероховатость
						    EditRough3D( ref rough );
					    }
				    }
				    else
					    // Выдать сообщение "Объект не создан"
              kompas.ksMessage( LoadString("IDS_NOCREATE") );
			    }
		    }
	    }
    }
    #endregion


    #region Обозначение базы 3D
    //-------------------------------------------------------------------------------
    // Создать обозначение базы
    // ---
    bool CreateBase3D( ref IBase3D pBase ) 
    {
	    bool res = false;

	    if ( pBase != null && doc3D != null )
	    {
			  // Указать в документе опорный объект для построения базы
        ksEntity baseObj = (ksEntity)doc3D.UserSelectEntity( null, "UserFilterProc", 
                                           LoadString("IDS_OBJ"), 0, this );
    		
			  if ( baseObj != null )
			  {
				  double x = 0, y = 0, z = 0;
				  // Указать точку привязки
				  doc3D.UserGetCursor( LoadString("IDS_POINT"), out x, out y, out z );
    			
				  // Преобразовать интерфейс объекта из API5 в API7
          IModelObject mObj = (IModelObject)kompas.TransferInterface( baseObj, 
                              (int)Kompas6Constants.ksAPITypeEnum.ksAPI7Dual, 0 );
    			
				  if ( mObj != null )
					  // Установить привязку базы
					  pBase.SetBranchBeginPoint( x, y, z, mObj );
    			
				  // Задать базовую плоскость
				  pBase.BasePlane = Kompas6Constants3D.ksObj3dTypeEnum.o3d_planeXOZ;
    			
				  // Применить параметры
				  res = pBase.Update();
			  }
		  }
	    return res;
    }


    //-------------------------------------------------------------------------------
    // Редактирование обозначения базы
    // ---
    void EditBase3D( ref IBase3D pBase )
    {
	    if ( pBase != null )
	    {
		    // Отменить автосортировку
		    pBase.AutoSorted = false;

		    // Получить интерфейс текста обозначения базы
		    IText txt = pBase.Text;
    		
		    if ( txt != null )
		    {
			    txt.Str = "";
    			
			    ITextLine line = txt.Add();
    			
			    if ( line != null )
			    {
				    // Добавить строку
				    AddTextItem( ref line, "A", Kompas6Constants.ksTextItemEnum.ksTItString );
				    // Добавить выражение типа суммы (нижний индекс может существовать
				    // только в таком выражении
				    AddTextItem( ref line, "", Kompas6Constants.ksTextItemEnum.ksTItSBase );
				    // Добавить нижний индекс
				    AddTextItem( ref line, "1", Kompas6Constants.ksTextItemEnum.ksTItSLowerIndex );
				    // Закончить выражение типа суммы
				    AddTextItem( ref line, "", Kompas6Constants.ksTextItemEnum.ksTItSEnd );
			    }
		    }

		    // Тип - произвольное положение
		    pBase.DrawType = false;
		    double x = 0, y = 0, z = 0;
		    // Сместить конечную точку обозначения базы
		    pBase.GetBranchEndPoint( out x, out y, out z );
		    pBase.SetBranchEndPoint( x + 10, y + 10, z );

		    // Получить интерфейс объекта дерева
		    IFeature7 featObj = (IFeature7)pBase;

		    if ( featObj != null )
			    // Изменить имя объекта в дереве
			    pBase.Name = GetNewName( ref featObj, "IDS_BASE" );

		    // Получить интерфейс свойств цвета объекта
		    IColorParam7 colorPars = (IColorParam7)pBase;
    		
		    // Изменить цвет на черный
		    if ( colorPars != null )
			    colorPars.Color = RGB( 0, 0, 0 );


		    pBase.Update();
	    }
    }


    //-------------------------------------------------------------------------------
    // Создание и редактирование обозначения базы
    // ---
    void Base3DWork()
    {
	    // Получить контейнер обозначений 3D
	    ISymbols3DContainer symbCont = GetSymbols3DContainer();
    	
	    if ( symbCont != null )
	    {
		    // Получить коллекцию обозначений базы 3D
		    IBases3D basesCol = symbCont.Bases3D;
    		
		    if ( basesCol != null )
		    {
			    // Добавить новое обозначение базы 3D
			    IBase3D newBase = basesCol.Add();
    			
			    if ( newBase != null )
			    {
				    // Создать обозначение базы 3D
				    if ( CreateBase3D(ref newBase) )
				    {	
					    string name = "";
					    // Получить интерфейс объекта дерева построения
					    IFeature7 feature = (IFeature7)newBase;
    					
					    // Получить имя объекта
					    if ( feature != null )
						    name = feature.Name;
    					
					    if ( kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
					    {
						    // Получить объект из коллекции по имени
						    IBase3D pBase = basesCol.get_Base3D( name );
						    // Редактировать обозначение базы
						    EditBase3D( ref pBase );
					    }
				    }
				    else
					    // Выдать сообщение "Объект не создан"
              kompas.ksMessage( LoadString("IDS_NOCREATE") );
			    }
		    }
	    }
    }
    #endregion


    #region Линия-выноска 3D
    //-------------------------------------------------------------------------------
    // Установить тексты линии-выноски
    // ---
    void SetLeader3DTexts( ref IBaseLeader3D baseLeader )
    {
	    ILeader leader = (ILeader)baseLeader;

	    if ( leader != null )
	    {
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
	    }
    }


    //-------------------------------------------------------------------------------
    // Создать ответвление линии-выноски
    // ---
    void CreateBranchsLeader( ref IBaseLeader3D leader )
    {
	    if ( leader != null && doc3D != null )
	    {
			  // Указать в документе опорный объект
        ksEntity baseObj = (ksEntity)doc3D.UserSelectEntity( null, "UserFilterProc", 
                                           LoadString("IDS_OBJ"), 0, this );
    		
			  double x =0 , y = 0, z = 0, x1 = 0, y1 = 0, z1 = 0;
			  // Указать точку начала ножки.
			  // Точка должна лежать на опорном объекте
			  doc3D.UserGetCursor( LoadString("IDS_BEGINBRANCH"), out x, out y, out z );
    		
			  // Указать объект, задающий плоскость обозначения
        ksEntity posObj = (ksEntity)doc3D.UserSelectEntity( null, "UserFilterProc", 
                                          LoadString("IDS_POSOBJ"), 0, this );
    		
			  if ( baseObj != null && posObj != null )
			  {
				  // Указать точку, задающую положние объектов
				  // Точка должна лежать в плоскости обозначения
				  doc3D.UserGetCursor( LoadString("IDS_BEGINSHELF"), out x1, out y1, out z1 );
				  // Преобразовать интерфейс объекта из API5 в API7
          IModelObject mObj1 = (IModelObject)kompas.TransferInterface( baseObj, 
                               (int)Kompas6Constants.ksAPITypeEnum.ksAPI7Dual, 0 );
          IModelObject mObj2 = (IModelObject)kompas.TransferInterface( posObj, 
                               (int)Kompas6Constants.ksAPITypeEnum.ksAPI7Dual, 0 );
    			
				  if ( mObj1 != null && mObj2 != null )
				  {	
					  // Получить интерфейс ответвлений
					  IBranchs3D branchs3D = (IBranchs3D)leader;
    				
					  // Добавить ответвление
					  if ( branchs3D != null )
						  branchs3D.AddBranchByPoint( x, y, z, mObj1 );
    				
					  // Установить объект, задающий плоскость обозначения
					  leader.PositionObject = mObj2;
					  leader.SetPosition( x1, y1, z1 );
				  }
			  }
		  }
    }


    //-------------------------------------------------------------------------------
    // Создать линию-выноску
    // ---
    bool CreateLeader3D( ref IBaseLeader3D leader ) 
    {
	    bool res = false;

	    if ( leader != null )
	    {
		    // Создать ответвления линии-выноски
		    CreateBranchsLeader( ref leader );
		    // Установить тексты линии-выноски
		    SetLeader3DTexts( ref leader );
		    // Применить параметры
		    res = leader.Update();
	    }
	    return res;
    }


    //-------------------------------------------------------------------------------
    // Изменить параметры линии-выноски
    // ---
    void EditLeaderPars( ref IBaseLeader3D baseLeader )
    {
	    if ( baseLeader != null )
	    {
		    ILeader leader = (ILeader)baseLeader;

		    if ( leader != null )
		    {
			    // Тип значка - знак склеивания
			    leader.SignType = Kompas6Constants.ksLeaderSignEnum.ksLGlueSign;
			    // Включить признак обоработки по контуру
			    leader.Arround = true;
    			
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
		    }
	    }
    }


    //-------------------------------------------------------------------------------
    // Добавить прямолинейное ответвление
    // ---
    void LeaderAddBranch( ref IBaseLeader3D leader ) 
    {
	    if ( leader != null && doc3D != null )
	    {
			  // Указать в документе опорный объект
        ksEntity baseObj = (ksEntity)doc3D.UserSelectEntity( null, "UserFilterProc", 
                                           LoadString("IDS_OBJ"), 0, this );
    		
			  double x = 0, y = 0, z = 0;
			  // Указать точку начала ножки
			  doc3D.UserGetCursor( LoadString("IDS_BEGINBRANCH"), out x, out y, out z );
    		
			  if ( baseObj != null )
			  {
				  // Преобразовать интерфейс объекта из API5 в API7
          IModelObject mObj = (IModelObject)kompas.TransferInterface( baseObj, 
                              (int)Kompas6Constants.ksAPITypeEnum.ksAPI7Dual, 0 );
    			
				  if ( mObj != null )
				  {	
					  // Получить интерфейс ответвлений
					  IBranchs3D branchs3D = (IBranchs3D)leader;
    				
					  // Добавить ответвление
					  if ( branchs3D != null )
						  branchs3D.AddBranchByPoint( x, y, z, mObj );
				  }
			  }
		  }
    }


    //-------------------------------------------------------------------------------
    // Редактировать линию-выноску
    // ---
    void EditLeader3D( ref IBaseLeader3D leader ) 
    {
	    if ( leader != null )
	    {
		    // Добавить прямолинейное ответвление
		    LeaderAddBranch( ref leader );
		    // Изменить параметры линии-выноски
		    EditLeaderPars( ref leader );

		    // Получить интерфейс объекта дерева
		    IFeature7 featObj = (IFeature7)leader;
    		
		    if ( featObj != null )
			    // Изменить имя объекта в дереве
			    leader.Name = GetNewName( ref featObj, "IDS_LEADER" );
    		
		    // Получить интерфейс свойств цвета объекта
		    IColorParam7 colorPars = (IColorParam7)leader;
    		
		    // Изменить цвет на коричневый
		    if ( colorPars != null )
			    colorPars.Color = RGB( 123, 40, 0 );

		    // Применить изменения
		    leader.Update();
	    }
    }


    //-------------------------------------------------------------------------------
    // Создание и редактирование линии-выноски 3D
    // ---
    void Leader3DWork()
    {
	    // Получить контейнер обозначений 3D
	    ISymbols3DContainer symbCont = GetSymbols3DContainer();
    	
	    if ( symbCont != null )
	    {
		    // Получить коллекцию линий-выносок 3D
		    ILeaders3D leadsCol = symbCont.Leaders3D;
    		
		    if ( leadsCol != null )
		    {
			    // Добавить новую линию-выноску 3D
			    IBaseLeader3D newLeader = leadsCol.Add( Kompas6Constants3D.ksObj3dTypeEnum.
                                                  o3d_leader3D );
    			
			    if ( newLeader != null )
			    {
				    // Создать линию-выноску 3D
				    if ( CreateLeader3D( ref newLeader) )
				    {	
					    string name = "";
					    // Получить интерфейс объекта дерева построения
					    IFeature7 feature = (IFeature7)newLeader;
    					
					    // Получить имя объекта
					    if ( feature != null )
						    name = feature.Name;
    					
					    if ( kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
					    {
						    // Получить объект из коллекции по имени
						    IBaseLeader3D leader = leadsCol.get_Leader3D( name );
						    // Редактировать линию-выноску
						    EditLeader3D( ref leader );
					    }
				    }
				    else
					    // Выдать сообщение "Объект не создан"
              kompas.ksMessage( LoadString("IDS_NOCREATE") );
			    }
		    }
	    }
    }
    #endregion


    #region Знак клеймения 3D
    //-------------------------------------------------------------------------------
    // Создать знак клеймения 3D
    // ---
    bool CreateBrandLeader3D( ref IBaseLeader3D baseLeader )
    {
	    bool res = false;
    	
	    if ( baseLeader != null )
	    {
		    // Создать ответвления
		    CreateBranchsLeader( ref baseLeader );

		    // Получить интерфейс знака клеймения
		    IBrandLeader brandLeader = (IBrandLeader)baseLeader;

		    if ( brandLeader != null )
			    // Направление
			    brandLeader.Direction = false;
    			
		    // Применить параметры
		    res = baseLeader.Update();
	    }
	    return res;
    }


    //-------------------------------------------------------------------------------
    // Редактировать знак клеймения
    // ---
    void EditBrandLeader3D( ref IBaseLeader3D leader ) 
    {
	    if ( leader != null )
	    {
		    // Добавить прямолинейное ответвление
		    LeaderAddBranch( ref leader );

		    // Установить параметры знака клеймения
		    IBrandLeader brandLeader = (IBrandLeader)leader;
    		
		    if ( brandLeader != null )
		    {
			    // Получить интерфейс текста обозначения
			    IText des = brandLeader.Designation;
    			
			    if ( des != null )
				    // Изменить текст
				    des.Str = LoadString( "IDS_MARK1" );
		    }

		    // Получить интерфейс объекта дерева
		    IFeature7 featObj = (IFeature7)leader;
    		
		    if ( featObj != null )
			    // Изменить имя объекта в дереве
			    leader.Name = GetNewName( ref featObj, "IDS_BRAND" );
    		
		    // Получить интерфейс свойств цвета объекта
		    IColorParam7 colorPars = (IColorParam7)leader;
    		
		    // Изменить цвет на фиолетовый
		    if ( colorPars != null )
			    colorPars.Color = RGB( 128, 0, 128 );

		    // Применить изменения
		    leader.Update();
	    }
    }


    //-------------------------------------------------------------------------------
    // Создание и редактирование знака клеймения
    // ---
    void BrandLeader3DWork()
    {
	    // Получить контейнер обозначений 3D
	    ISymbols3DContainer symbCont = GetSymbols3DContainer();
    	
	    if ( symbCont != null )
	    {
		    // Получить коллекцию линий-выносок 3D
		    ILeaders3D leadsCol = symbCont.Leaders3D;
    		
		    if ( leadsCol != null )
		    {
			    // Добавить новый знак клеймения 3D
			    IBaseLeader3D newBrand = leadsCol.Add( Kompas6Constants3D.ksObj3dTypeEnum.
                                                 o3d_brandLeader3D );
    			
			    if ( newBrand != null )
			    {
				    // Создать знак клеймения 3D
				    if ( CreateBrandLeader3D(ref newBrand) )
				    {	
					    string name = "";
					    // Получить интерфейс объекта дерева построения
					    IFeature7 feature = (IFeature7)newBrand;
    					
					    // Получить имя объекта
					    if ( feature != null )
						    name = feature.Name;
    					
					    if ( kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
					    {
						    // Получить объект из коллекции по имени
						    IBaseLeader3D brandLeader = leadsCol.get_Leader3D( name );
						    // Редактировать знак клеймения
						    EditBrandLeader3D( ref brandLeader );
					    }
				    }
				    else
					    // Выдать сообщение "Объект не создан"
              kompas.ksMessage( LoadString("IDS_NOCREATE") );
			    }
		    }
	    }
    }
    #endregion


    #region Знак маркировки 3D
    //-------------------------------------------------------------------------------
    // Создать знак маркировки
    // ---
    bool CreateMarkLeader3D( ref IBaseLeader3D baseLeader )
    {
	    bool res = false;

	    if ( baseLeader !=null )
	    {
		    // Создать ответвление
		    CreateBranchsLeader( ref baseLeader );

		    // Получить интерфейс знака маркировки
		    IMarkLeader markLeader = (IMarkLeader)baseLeader;

		    if ( markLeader != null )
		    {	
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
		    }
		    // Применить параметры
		    res = baseLeader.Update();
	    }
	    return res;
    }


    //-------------------------------------------------------------------------------
    // Редактировать знак маркировки
    // ---
    void EditMarkLeader3D( ref IBaseLeader3D baseLeader )
    {
	    if ( baseLeader != null )
	    {
		    // Добавить прямолинейное ответвление
		    LeaderAddBranch( ref baseLeader );

		    // Получить интерфейс знака маркировки
		    IMarkLeader markLeader = (IMarkLeader)baseLeader;
    		
		    if ( markLeader != null )
		    {
			    // Получить интерфейс текста обозначения
			    IText des = markLeader.Designation;
    			
			    if ( des != null )
				    // Изменить текст
				    des.Str = "10";
		    }

		    // Получить интерфейс объекта дерева
		    IFeature7 featObj = (IFeature7)baseLeader;
    		
		    if ( featObj != null )
			    // Изменить имя объекта в дереве
			    baseLeader.Name = GetNewName( ref featObj, "IDS_BRAND" );
    		
		    // Получить интерфейс свойств цвета объекта
		    IColorParam7 colorPars = (IColorParam7)baseLeader;
    		
		    // Изменить цвет на сиреневый
		    if ( colorPars != null )
			    colorPars.Color = RGB( 204, 153, 255 );
    		
		    // Применить изменения
		    baseLeader.Update();
	    }
    }


    //-------------------------------------------------------------------------------
    // Создание и редактирование знака маркировки 3D
    // ---
    void MarkLeader3DWork()
    {
	    // Получить контейнер обозначений 3D
	    ISymbols3DContainer symbCont = GetSymbols3DContainer();
    	
	    if ( symbCont != null )
	    {
		    // Получить коллекцию линий-выносок 3D
		    ILeaders3D leadsCol = symbCont.Leaders3D;
    		
		    if ( leadsCol != null )
		    {
			    // Добавить новый знак маркировки 3D
			    IBaseLeader3D newMark = leadsCol.Add( Kompas6Constants3D.ksObj3dTypeEnum.
                                                o3d_markLeader3D );
    			
			    if ( newMark != null )
			    {
				    // Создать знак маркировки 3D
				    if ( CreateMarkLeader3D(ref newMark) )
				    {	
					    string name = "";
					    // Получить интерфейс объекта дерева построения
					    IFeature7 feature = (IFeature7)newMark;
    					
					    // Получить имя объекта
					    if ( feature != null )
						    name = feature.Name;

					    if ( kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
					    {
						    // Получить объект из коллекции по имени
						    IBaseLeader3D markLeader = leadsCol.get_Leader3D( name );
						    // Редактировать знак маркировки
						    EditMarkLeader3D( ref markLeader );
					    }
				    }
				    else
					    // Выдать сообщение "Объект не создан"
              kompas.ksMessage( LoadString("IDS_NOCREATE") );
			    }
		    }
	    }
    }
    #endregion


    #region Допуск формы 3D
    //-------------------------------------------------------------------------------
    // Создать ответвление для доопуска формы
    // ---
    void CreateToleranceBranch( ref ITolerance3D tolerance )
    {
	    if ( tolerance != null && doc3D != null )
	    {
			  // Указать в документе объект, на который будет указывать ножка
        ksEntity baseObj = (ksEntity)doc3D.UserSelectEntity( null, "UserFilterProc", 
                                           LoadString("IDS_OBJ"), 0, this );
    		
			  double x = 0, y = 0, z = 0, x1 = 0, y1 = 0, z1 = 0;
			  // Указать точку конца ножки
			  // Точка должна лежать на опорном объекте
			  doc3D.UserGetCursor( LoadString("IDS_ENDPOINT"), out x, out y, out z );
    		
			  // Указать объект, задающий плоскость обозначения
        ksEntity posObj = (ksEntity)doc3D.UserSelectEntity( null, "UserFilterProc", 
                                          LoadString("IDS_POSOBJ"), 0, this );
    		
			  if ( baseObj != null && posObj != null )
			  {
				  // Указать точку вставки таблицы
				  doc3D.UserGetCursor( LoadString("IDS_TABLE"), out x1, out y1, out z1 );
				  // Преобразовать интерфейс объекта из API5 в API7
          IModelObject mObj1 = (IModelObject)kompas.TransferInterface( baseObj, 
                               (int)Kompas6Constants.ksAPITypeEnum.ksAPI7Dual, 0 );
          IModelObject mObj2 = (IModelObject)kompas.TransferInterface( posObj, 
                               (int)Kompas6Constants.ksAPITypeEnum.ksAPI7Dual, 0 );
    			
				  if ( mObj1 != null && mObj2 != null )
				  {	
					  // Получить интерфейс ответвлений
					  IBranchs3D branchs3D = (IBranchs3D)tolerance;
    				
					  // Добавить ответвление
					  if ( branchs3D != null )
						  branchs3D.AddBranchByPoint( x, y, z, mObj1 );
    				
					  // Установить объект, задающий плоскость обозначения
					  tolerance.PositionObject = mObj2;
					  // Установить точку, задающую положение объекта
					  tolerance.SetPosition( x1, y1, z1 );
				  }
			  }
		  }
    }


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
    				
				    if ( txt != null)
					    txt.Str = "@30~";
			    }
		    }
	    }
    }


    //-------------------------------------------------------------------------------
    // Создать допуск формы
    // ---
    bool CreateTolerance3D( ref ITolerance3D tolerance ) 
    {
	    bool res = false;

	    if ( tolerance != null )
	    {
		    // Создать ответвление
		    CreateToleranceBranch( ref tolerance );
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
		    // Применить параметры
		    res = tolerance.Update();
	    }
	    return res;
    }


    //-------------------------------------------------------------------------------
    // Добавить ответвление для допуска формы
    // ---
    void ToleranceAddBranch( ref ITolerance3D tolerance ) 
    {
	    if ( tolerance != null && doc3D != null )
	    {
			  // Указать в документе объект, на который будет указывать ножка линии-выноски
        ksEntity baseObj = (ksEntity)doc3D.UserSelectEntity( null, "UserFilterProc", 
                                           LoadString("IDS_OBJ"), 0, this );
    		
			  double x = 0, y = 0, z = 0;
			  // Указать точку конца ответвления
        // Точка должна лежать на опорном объекте
			  doc3D.UserGetCursor( LoadString("IDS_ENDPOINT"), out x, out y, out z );
    		
			  if ( baseObj != null )
			  {
				  // Преобразовать интерфейс объекта из API5 в API7
          IModelObject mObj = (IModelObject)kompas.TransferInterface( baseObj, 
                               (int)Kompas6Constants.ksAPITypeEnum.ksAPI7Dual, 0 );
    			
				  if ( mObj != null )
				  {	
					  // Получить интерфейс ответвлений
					  IBranchs3D branchs3D = (IBranchs3D)tolerance;
    				
					  // Добавить ответвление
					  if ( branchs3D != null )
						  branchs3D.AddBranchByPoint( x, y, z, mObj );

					  // Ответвление со стрелкой
					  tolerance.set_ArrowType( 1, true );
				  }
			  }
		  }
    }


    //-------------------------------------------------------------------------------
    // Редактировать допуск формы
    // ---
    void EditTolerance3D( ref ITolerance3D tolerance ) 
    {
	    if ( tolerance != null )
	    {
		    ToleranceAddBranch( ref tolerance );

		    // Получитьинтерфейс параметров допуска формы
		    IToleranceParam tolPar = (IToleranceParam)tolerance;

		    if ( tolPar != null )
		    {
			    // Задать признак вертикальности
			    tolPar.Vertical = true;
			    // Положение базовой точки относительно таблицы - слева посередине
			    tolPar.BasePointPos = Kompas6Constants.ksTablePointEnum.ksTPLeftBottom;
		    }

		    // Получить интерфейс объекта дерева
		    IFeature7 featObj = (IFeature7)tolerance;
    		
		    if ( featObj != null )
			    // Изменить имя объекта в дереве
			    tolerance.Name = GetNewName( ref featObj, "IDS_TOLERANCE" );
    		
		    // Получить интерфейс свойств цвета объекта
		    IColorParam7 colorPars = (IColorParam7)tolerance;
    		
		    // Изменить цвет на темно-красный
		    if ( colorPars != null )
			    colorPars.Color = RGB( 186, 12, 34 );

		    // Применить изменения
		    tolerance.Update();
	    }
    }


    //-------------------------------------------------------------------------------
    // Создание и редактирование допуска формы 3D
    // ---
    void Tolerance3DWork()
    {
	    // Получить контейнер обозначений 3D
	    ISymbols3DContainer symbCont = GetSymbols3DContainer();
    	
	    if ( symbCont != null )
	    {
		    // Получить коллекцию допусков формы 3D
		    ITolerances3D tolerCol = symbCont.Tolerances3D;
    		
		    if ( tolerCol != null )
		    {
			    // Добавить новый допуск формы 3D
			    ITolerance3D newTolerance = tolerCol.Add();
    			
			    if ( newTolerance != null )
			    {
				    // Создать допуск формы 3D
				    if ( CreateTolerance3D(ref newTolerance) )
				    {	
					    string name = "";
					    // Получить интерфейс объекта дерева построения
					    IFeature7 feature = (IFeature7)newTolerance;
    					
					    // Получить имя объекта
					    if ( feature != null )
						    name = feature.Name;
    					
					    if ( kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
					    {
						    // Получить объект из коллекции по имени
						    ITolerance3D tolerance = tolerCol.get_Tolerance3D( name );
						    // Редактировать допуск формы
						    EditTolerance3D( ref tolerance );
					    }
				    }
				    else
					    // Выдать сообщение "Объект не создан"
              kompas.ksMessage( LoadString("IDS_NOCREATE") );
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
