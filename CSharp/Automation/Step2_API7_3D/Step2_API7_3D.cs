using Kompas6API5;
using KompasAPI7;
using Kompas6Constants;
using Kompas6Constants3D;
using System;
using System.Resources;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Windows.Forms;


namespace Step2_API7_3D
{
  [ClassInterface(ClassInterfaceType.AutoDual)]
	public class Dimensions3D
	{
    private KompasObject      kompas;
    private IApplication      appl;         // Интерфейс приложения
    private IKompasDocument3D doc;          // Интерфейс документа 3D в API7 
    private ksDocument3D      doc3D;        // Интерфейс документа 3D в API5
    private ResourceManager   resMng = new ResourceManager( typeof(Dimensions3D) );   // Менеджер ресурсов
    private int               oType = (int)Kompas6Constants3D.Obj3dType.o3d_edge;

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
        case 1: CreateLineDimension3D();      break;  // Создание линейного размера 3D
        case 2: LineDimension3DNavigation();  break;  // Навигация по коллекции линейных размеров 3D
        case 3: EditLineDimension3D();        break;  // Редактирование линейного размера 3D
        case 4: RadialDimension3DWork();      break;  // Создание и редактирование радиального размера 3D
        case 5: DiametralDimension3DWork();   break;  // Создание и редактирование диаметрального размера 3D
        case 6: AngleDimension3DWork();       break;  // Создание и редактирование углового размера 3D
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
    // Установить объект привязки и базовую плоскость
    // ---
    bool SetLineDimObjectPlane( ref ILineDimension3D dim )
    {
	    bool res = false;
	    // Фильтрация объектов - ребра
	    oType = (int)Kompas6Constants3D.Obj3dType.o3d_edge;
    	
	    if ( doc3D != null )
	    {
		    // Указать в документе объект - ребро
        ksEntity obj1 = (ksEntity)doc3D.UserSelectEntity( null, "UserFilterProc", 
                         LoadString("IDS_OBJ1"), 0, this );

		    if ( obj1 != null )
		    {
			    // Тип фильтрации - грани
			    oType = (int)Kompas6Constants3D.Obj3dType.o3d_face;
			    // Указать в документе объект - грань
          ksEntity plane = (ksEntity)doc3D.UserSelectEntity( null, "UserFilterProc", 
                            LoadString("IDS_PLANE"), 0, this );

			    if ( plane != null )
			    {
				    // Преобразовать интерфейсы ребра и грани из API5 в API7
            IModelObject mObj1 = (IModelObject)kompas.TransferInterface( obj1, 
                                 (int)Kompas6Constants.ksAPITypeEnum.ksAPI7Dual, 0 );
            IModelObject mPlane = (IModelObject)kompas.TransferInterface( plane, 
                                  (int)Kompas6Constants.ksAPITypeEnum.ksAPI7Dual, 0 );
    				
				    if ( mObj1 != null && mPlane != null && dim != null )
				    {
					    // Установить первый объект размера (Второй объект устанавливается 
					    // только в том случае, если измеряется расстояние между точками)
					    dim.Object1 = mObj1;
					    // Установить базовую плоскость размера
					    dim.Plane = mPlane;
					    res = true;
				    }
			    }
		    }
	    }
	    return res;
    }


    //-------------------------------------------------------------------------------
    // Установить новую базовую плоскость
    // ---
    bool SetNewPlane( ref ILineDimension3D dim )
    {
	    bool res = false;

	    if ( dim != null && doc3D != null )
	    {
		    oType = (int)Kompas6Constants3D.Obj3dType.o3d_face;
		    // Указать в документе объект
		    ksEntity obj1 = (ksEntity)doc3D.UserSelectEntity( null, "UserFilterProc", 
                                        LoadString("IDS_NEWPLANE"), 0, this );

		    if ( obj1 != null )
		    {
			    // Преобразовать интерфейс объекта из API5 в API7
			    IModelObject plane = (IModelObject)kompas.TransferInterface( obj1, 
                               (int)Kompas6Constants.ksAPITypeEnum.ksAPI7Dual, 0 );

			    if ( plane != null )
			    {
				    // Установить базовую плоскость
				    dim.Plane = plane;
				    res = true;
			    }
		    }
	    }
	    return res;
    }
    #endregion


    #region Линейный размер 3D
    //-------------------------------------------------------------------------------
    // Создание линейного размера 3D
    // ---
    void CreateLineDimension3D()
    {
	    // Получить контейнер обозначений 3D
	    ISymbols3DContainer symbCont = GetSymbols3DContainer();

	    if ( symbCont != null )
	    {
		    // Получить коллекцию линейных размеров 3D
		    ILineDimensions3D dimsCol = symbCont.LineDimensions3D;

		    if ( dimsCol != null )
		    {
			    // Добавить новый линейный размер 3D
			    ILineDimension3D newDim = (ILineDimension3D)dimsCol.Add( Kompas6Constants3D.
                                     ksObj3dTypeEnum.o3d_lineDimension3D );

			    if ( newDim != null )
			    {
				    bool create = false;

				    // Если удалось установить объект привязки и базовую плоскость
				    if ( SetLineDimObjectPlane(ref newDim) )
				    {
					    // Установить длину размера
					    newDim.Length = 30;
					    // Применить параметры
					    create = (bool)!!newDim.Update();		
    					
				    }
				    else
					    // Выдать сообщение "Объект не создан"
              kompas.ksMessage( LoadString("IDS_NOCREATE") );

				    // Если объект не создался, удалить
				    if ( !create )
				    {
					    IFeature7 obj = (IFeature7)newDim;
    					
					    if ( obj != null )
						    obj.Delete();
				    }
			    }
		    }
	    }
    }


    //-------------------------------------------------------------------------------
    // Выдать сообщение с параметрами линейного размера 3D
    // ---
    void GetLineDimensionPar( ref ILineDimension3D dim )
    {
	    if ( dim != null )
	    {
		    // Получить длину размера
		    double lenght = dim.Length;
		    double val = 0;

		    // Получить значение размера
		    IDimensionText dimText = (IDimensionText)dim;

		    if ( dimText != null )
			    val = dimText.NominalValue;

		    // Сформировать сообщение
        string buf = LoadString( "IDS_LINEDIM3D" ) + "\n\r";
        buf += string.Format( LoadString("IDS_LENGTH"), lenght ) + "\n\r";
        buf += string.Format( LoadString("IDS_DIMVAL"), val );
        kompas.ksMessage( buf );
	    }
    }


    //-------------------------------------------------------------------------------
    // Навигация по коллекции линейных размеров 3D
    // ---
    void LineDimension3DNavigation()
    {
	    if ( doc != null )
	    {
		    // Получить контейнер обозначений 3D
		    ISymbols3DContainer symbCont = GetSymbols3DContainer();
    		
		    if ( symbCont != null )
		    {
			    // Получить коллекцию линейных размеров 3D
			    ILineDimensions3D dimsCol = symbCont.LineDimensions3D;
    			
			    if ( doc3D != null )
			    {
				    // Получить менеджер селектирования
				    ksSelectionMng selectMng = (ksSelectionMng)doc3D.GetSelectionMng();

				    if ( dimsCol != null && selectMng != null )
				    {
					    // Цикл по коллекции 
					    for ( long i = 0; i < dimsCol.Count; i++ )
					    {	
						    // Получить размер из коллекции по индексу
						    ILineDimension3D lineDim = (ILineDimension3D)dimsCol.get_LineDimension3D( i );

						    if ( lineDim != null )
						    {
							    // Преобразовать интерфейс объекта 3D из API7 в API5
							    ksEntity dimObj = (ksEntity)kompas.TransferInterface( lineDim,
                                    (int)Kompas6Constants.ksAPITypeEnum.ksAPI5Auto,
                                    (int)Kompas6Constants3D.Obj3dType.o3d_entity );

                  if ( dimObj != null )
							    {
								    // Подсветить объект
								    selectMng.Select( dimObj );
								    // Выдать сообщение с параметрами размера
								    GetLineDimensionPar( ref lineDim );
								    // Убрать подсветку
								    selectMng.Unselect( dimObj );
							    }
						    }
					    }
				    }
			    }
		    }
	    }
    }


    //-------------------------------------------------------------------------------
    // Изменить параметры линейного размера 3D
    // ---
    void ChangeLineDimensionPar( ref ILineDimension3D dim )
    {
	    if ( dim != null )
	    {
		    // Получить интерфейс параметров размера
		    IDimensionParams dimPars = (IDimensionParams)dim;

		    if ( dimPars != null )
		    {
			    // Тип стрелок - засечка
			    dimPars.ArrowType1 = Kompas6Constants.ksArrowEnum.ksNotch;
			    dimPars.ArrowType2 = Kompas6Constants.ksArrowEnum.ksNotch;
			    // Расположение стрелок - снаружи
			    dimPars.ArrowPos = Kompas6Constants.ksDimensionArrowPosEnum.ksDimArrowOutside;
			    // Направление полки - вправо
			    dimPars.ShelfDirection = Kompas6Constants.ksShelfDirectionEnum.ksLSRight;
			    // Длина полки
			    dimPars.ShelfLength = 10;
			    // Угол наклона полки
			    dimPars.ShelfAngle = 45;
		    }
	    }
    }


    //-------------------------------------------------------------------------------
    // Редактирование линейного размера 3D
    // ---
    void EditLineDimension3D()
    {
	    if ( doc3D != null )
	    {
		    oType = (int)Kompas6Constants3D.Obj3dType.o3d_lineDimension3D;
		    // Указать в документе объект
		    ksEntity obj1 = (ksEntity)doc3D.UserSelectEntity( null, "UserFilterProc", 
                                        LoadString("IDS_DIM"), 0, this );

		    if ( obj1 != null )
		    {
			    // Преобразовать интерфейс объекта из API5 в API7
			    IModelObject mObj1 = (IModelObject)kompas.TransferInterface( obj1, 
                               (int)Kompas6Constants.ksAPITypeEnum.ksAPI7Dual, 0 );

			    if ( mObj1 != null && (int)mObj1.ModelObjectType == (int)Kompas6Constants3D.
                                 Obj3dType.o3d_lineDimension3D )
			    {
				    // Получить интерфейс линейного размера 3D
				    ILineDimension3D lineDim = (ILineDimension3D)mObj1;

				    if ( lineDim != null )
				    {
					    // Установить новую базовую плоскость
					    if ( !SetNewPlane(ref lineDim) )
						    // Выдать сообщение "Не удалось установить новую базовую плоскость"
                kompas.ksMessage( LoadString("IDS_NOTSETPLANE") );

					    // Изменить параметры размера
					    ChangeLineDimensionPar( ref lineDim );
					    // Уменьшить длину размера
					    lineDim.Length = lineDim.Length - 10;
					    // Применить изменения
					    lineDim.Update();
				    }
			    }
			    else
				    // Выдать сообщение "Объект не является линейным размером"
            kompas.ksMessage( LoadString("IDS_NOTDIM") );
		    }
	    }
    }
    #endregion


    #region Радиальный размер 3D
    //-------------------------------------------------------------------------------
    // Создать радиальный размер 3D
    // ---
    bool CreateRadDimension3D( ref IRadialDimension3D dim )
    {
	    bool res = false;

	    if ( dim != null && doc3D != null )
	    {
			  oType = (int)Kompas6Constants3D.Obj3dType.o3d_edge;
			  // Указать в документе объект - ребро
        ksEntity edge = (ksEntity)doc3D.UserSelectEntity( null, "UserFilterProc", 
                                        LoadString("IDS_OBJ1"), 0, this );

			  if ( edge != null )
			  {
				  // Получить интерфейс свойств ребра
				  ksEdgeDefinition edgeDef = (ksEdgeDefinition)edge.GetDefinition();

				  // Проверить, является ли ребро круговым
				  if ( edgeDef != null && edgeDef.IsCircle() )
				  {
					  // Преобразовать интерфейс объекта из API5 в API7
					  IModelObject mObj1 = (IModelObject)kompas.TransferInterface( edge, 
                                 (int)Kompas6Constants.ksAPITypeEnum.ksAPI7Dual, 0 );

					  if ( mObj1 != null )
						  // Установить объект
						  dim.Object1 = mObj1;

					  // Получить интерфейс параметров объекта
					  IDimensionParams dimPars = (IDimensionParams)dim;

					  if ( dimPars != null )
					  {
						  // Направление полки - влево
						  dimPars.ShelfDirection = Kompas6Constants.ksShelfDirectionEnum.ksLSLeft;
						  // Длина полки
						  dimPars.ShelfLength = 15;
						  // Угол наклона полки
						  dimPars.ShelfAngle = 30;
					  }

					  // Тип размера - не от центра
					  dim.DimensionType = false;
					  // Применить параметры
					  res = dim.Update();
				  }
				  else
					  // Выдать сообщение "Ребро не является круговым"
            kompas.ksMessage( LoadString("IDS_NOTCIRCLE") );
			  }
		  }
	    return res;
    }


    //-------------------------------------------------------------------------------
    // Редактировать радиальный размер 3D
    // ---
    void EditRadDimension3D( ref IRadialDimension3D dim )
    {
	    if ( dim != null )
	    {
		    // Тип размера - от центра
		    dim.DimensionType = true;

		    // Получить интерфейс параметров размера
		    IDimensionParams dimPars = (IDimensionParams)dim;

		    if ( dimPars != null )
			    // Отключить полку
			    dimPars.ShelfDirection = Kompas6Constants.ksShelfDirectionEnum.ksLSNone;

		    // Получить интерфейс текстов размера
		    IDimensionText dimText = (IDimensionText)dim;

		    if ( dimText != null )
		    {
			    // Задать квалитет
			    dimText.Tolerance = "h6";

			    // Включить квалитет
			    dimText.ToleranceOn = true;

			    // Получить интерфейс текста нижнего отклонения
			    ITextLine lowDev = dimText.LowDeviation;
    			
			    // Задать текстовую строку
			    if ( lowDev != null )
				    lowDev.Str = "+0.021";
		    }
		    // Применить изменения
		    dim.Update();
	    }
    }


    //-------------------------------------------------------------------------------
    // Создание и редактирование радиального размера 3D
    // ---
    void RadialDimension3DWork()
    {
	    // Получить контейнер обозначений 3D
	    ISymbols3DContainer symbCont = GetSymbols3DContainer();

	    if ( symbCont != null )
	    {
		    // Получить коллекцию радиальных размеров 3D
		    IRadialDimensions3D dimsCol = symbCont.RadialDimensions3D;

		    if ( dimsCol != null )
		    {
			    // Добавить новый радиальный размер 3D
			    IRadialDimension3D newDim = dimsCol.Add();

			    if ( newDim != null )
			    {
				    // Создать радиальный размер 3D
				    if ( CreateRadDimension3D(ref newDim) )
				    {	
					    string name = "";
					    // Получить интерфейс объекта дерева построения
					    IFeature7 feature = (IFeature7)newDim;
              				
					    // Получить имя объекта
					    if ( feature != null )
						    name = feature.Name;

					    if ( kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
					    {
						    // Получить объект из коллекции по имени
						    IRadialDimension3D radDim = dimsCol.get_RadialDimension3D( name );
						    // Редактировать размер
						    EditRadDimension3D( ref radDim );
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


    #region Диаметральный размер 3D
    //-------------------------------------------------------------------------------
    // Создать димаметральный размер
    // ---
    bool CreateDiamDimension3D( ref IDiametralDimension3D dim )
    {
	    bool res = false;
    	
	    if ( dim != null && doc3D != null )
	    {
			  oType = (int)Kompas6Constants3D.Obj3dType.o3d_edge;
			  // Указать в документе объект - ребро
        ksEntity edge = (ksEntity)doc3D.UserSelectEntity( null, "UserFilterProc", 
                                        LoadString("IDS_OBJ1"), 0, this );
			  
        if ( edge != null )
			  {
				  // Получить интерфейс свойств ребра
				  ksEdgeDefinition edgeDef = (ksEdgeDefinition)edge.GetDefinition();
    			
				  // Проверить, является ли ребро круговым
				  if ( edgeDef != null && edgeDef.IsCircle() )
				  {
					  // Преобразовать интерфейс объекта из API5 в API7
            IModelObject mObj1 = (IModelObject)kompas.TransferInterface( edge, 
                                 (int)Kompas6Constants.ksAPITypeEnum.ksAPI7Dual, 0 );
    				
					  if ( mObj1 != null )
						  // Установить объект
						  dim.Object1 = mObj1;
    				
					  // Применить параметры
					  res = dim.Update();
				  }
				  else
					  // Выдать сообщение "Ребро не является круговым"
            kompas.ksMessage( LoadString("IDS_NOTCIRCLE") );
			  }
		  }
	    return res;
    }


    //-------------------------------------------------------------------------------
    // Редактировать диаметральный размер
    // ---
    void EditDiamDimension3D( ref IDiametralDimension3D dim )
    {
	    if ( dim != null )
	    {
		    // Тип размера - с обрывом
		    dim.DimensionType = true;

		    // Получить интерфейс текстов размера
		    IDimensionText dimText = (IDimensionText)dim;

		    if ( dimText != null )
		    {
			    // Получить интерфейс текста под размерной надписью
			    IText txtUnder = dimText.TextUnder;
    			
			    // Задать текстовую строку
			    if ( txtUnder != null )
				    txtUnder.Str = LoadString( "IDS_DIMTEXT" );

			    // Подчеркнутый текст
			    dimText.Underline = true;
		    }
		    // Применить изменения
		    dim.Update();
	    }
    }


    //-------------------------------------------------------------------------------
    // Создание и редактирование диаметрального размера 3D
    // ---
    void DiametralDimension3DWork()
    {
	    // Получить контейнер обозначений 3D
	    ISymbols3DContainer symbCont = GetSymbols3DContainer();
    	
	    if ( symbCont != null )
	    {
		    // Получить коллекцию диаметральных размеров 3D
		    IDiametralDimensions3D dimsCol = symbCont.DiametralDimensions3D;
    		
		    if ( dimsCol != null )
		    {
			    // Добавить новый диаметральный размер 3D
			    IDiametralDimension3D newDim = dimsCol.Add();
    			
			    if ( newDim != null )
			    {
				    // Создать диаметральный размер 3D
				    if ( CreateDiamDimension3D(ref newDim) )
				    {	
					    string name = "";
					    // Получить интерфейс объекта дерева построения
					    IFeature7 feature = (IFeature7)newDim;
    					
					    // Получить имя объекта
					    if ( feature != null )
						    name = feature.Name;
    					
					    if ( kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
					    {
						    // Получить объект из коллекции по имени
						    IDiametralDimension3D diamDim = dimsCol.get_DiametralDimension3D( name );
						    // Редактировать размер
						    EditDiamDimension3D( ref diamDim );
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


    #region Угловой размер 3D
    //-------------------------------------------------------------------------------
    // Создать угловой размер 3D
    // ---
    bool CreateAngleDimension3D( ref IAngleDimension3D dim )
    {
	    bool res = false;

	    if ( dim != null && doc3D != null )
	    {
			  // Указать в документе 1-й объект
        ksEntity obj1 = (ksEntity)doc3D.UserSelectEntity( null, "UserFilterProc", 
                                        LoadString("IDS_OBJECT1"), 0, this );
			  // Указать в документе 2-й объект
        ksEntity obj2 = (ksEntity)doc3D.UserSelectEntity( null, "UserFilterProc", 
                                        LoadString("IDS_OBJECT2"), 0, this );

			  if ( obj1 != null && obj2 != null )
			  {
				  // Преобразовать интерфейсы объектов из API5 в API7
          IModelObject mObj1 = (IModelObject)kompas.TransferInterface( obj1, 
                               (int)Kompas6Constants.ksAPITypeEnum.ksAPI7Dual, 0 );
          IModelObject mObj2 = (IModelObject)kompas.TransferInterface( obj2, 
                               (int)Kompas6Constants.ksAPITypeEnum.ksAPI7Dual, 0 );

				  if ( mObj1 != null && mObj2 != null )
				  {
					  // 1-й объект
					  dim.Object1 = mObj1;
					  // 2-й объект
					  dim.Object2 = mObj2;
					  // Длина размерной линии
					  dim.Length = 20;
					  // Применить параметры
					  res = dim.Update();
				  }
			  }
		  }
	    return res;
    }


    //-------------------------------------------------------------------------------
    // Редактировать угловой размер 3D
    // ---
    void EditAngleDimension3D( ref IAngleDimension3D dim )
    {
	    if ( dim != null )
	    {
		    // Тип размера - на максимальный (тупой) угол
		    dim.DimensionType = Kompas6Constants.ksAngleDimTypeEnum.ksADMaxAngle;
		    // Увеличить длину размерной линии
		    dim.Length = dim.Length + 10;

		    // Получить интерфейс параметров размера
		    IDimensionParams dimPars = (IDimensionParams)dim;

		    if ( dimPars != null )
			    // Размещение размерной надписи - параллельно, в разрезе линии
			    dimPars.TextOnLine = Kompas6Constants.ksDimensionTextPosEnum.
                               ksDimTextParallelInCut;

		    // Получить интерфейс текстов размера
		    IDimensionText dimText = (IDimensionText)dim;

		    if ( dimText != null )
			    // Изменить формат отображения текста на десятичный
			    dimText.TextFormat = Kompas6Constants.ksDimTextFormatEnum.ksDimTextFormatGDD;

		    // Применить изменения
		    dim.Update();
	    }
    }


    //-------------------------------------------------------------------------------
    // Создание и редактирование углового размера 3D
    // ---
    void AngleDimension3DWork()
    {
	    // Получить контейнер обозначений 3D
	    ISymbols3DContainer symbCont = GetSymbols3DContainer();
    	
	    if ( symbCont != null )
	    {
		    // Получить коллекцию угловых размеров 3D
		    IAngleDimensions3D dimsCol = symbCont.AngleDimensions3D;
    		
		    if ( dimsCol != null )
		    {
			    // Добавить новый угловой размер 3D
			    IAngleDimension3D newDim = dimsCol.Add();
    			
			    if ( newDim != null )
			    {
				    // Создать угловой размер 3D
				    if ( CreateAngleDimension3D(ref newDim) )
				    {	
					    string name = "";
					    // Получить интерфейс объекта дерева построения
					    IFeature7 feature = (IFeature7)newDim;
    					
					    // Получить имя объекта
					    if ( feature != null )
						    name = feature.Name;
    					
					    if ( kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
					    {
						    // Получить объект из коллекции по имени
						    IAngleDimension3D angDim = dimsCol.get_AngleDimension3D( name );
						    // Редактировать размер
						    EditAngleDimension3D( ref angDim );
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
