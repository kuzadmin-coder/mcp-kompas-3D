using Kompas6API5;
using KompasAPI7;
using Kompas6Constants;
using KAPITypes;
using System;
using System.Resources;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Windows.Forms;


namespace Step2_API7_2D
{
  //-------------------------------------------------------------------------------
  // Примеры работы с размерами 2D в API7
  // ---
  [ClassInterface(ClassInterfaceType.AutoDual)]
  public class SamplesDimension
  {
    private KompasObject      kompas;
    private IApplication      appl;         // Интерфейс приложения
    private IKompasDocument2D doc;          // Интерфейс документа 2D в API7 
    private ksDocument2D      doc2D;        // Интерфейс документа 2D в API5
    private ResourceManager   resMng = new ResourceManager( typeof(SamplesDimension) );       // Менеджер ресурсов
   
    
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
          result = LoadString( "IDS_MENU7" );
          command = 7;
        break;

        case 8:
          result = LoadString( "IDS_MENU8" );
          command = 8;
        break;

        case 9:
          result = LoadString( "IDS_MENU9" );
          command = 9;
        break;

        case 10:
          result = LoadString( "IDS_MENU10" );
          command = 10;
        break;

        case 11:
          result = LoadString( "IDS_MENU11" );
          command = 11;
        break;

        case 12:
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
        case 1:   CreateLineDimension();      break;  // Создать линейный размер
        case 2:   LineDimensionNavigation();  break;  // Навигация по коллекции линейных размеров
        case 3:   EditLineDimension();        break;  // Редактирование линейного размера
        case 4:   RadialDimensionWork();      break;  // Создание и редактирование радиального размера
        case 5:   DiamrtralDimensionWork();   break;  // Создание и редактирование диаметрального размера
        case 6:   AngleDimensionWork();       break;  // Создание и редактирование углового размера
        case 7:   ArcDimensionWork();         break;  // Создание и редактирование размера дуги окружности
        case 8:   BreakLineDimensionWork();   break;  // Создание и редактирование линейного размера с обрывом
        case 9:   BreakRadialDimensionWork(); break;  // Создание и редактирование радиального размера с изломом
        case 10:  BreakAngleDimensionWork();  break;  // Создание и редактирование углового размера с обрывом
        case 11:  HeightDimensionWork();      break;  // Создание и редактирование размера высоты
      }
    }


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

    #region Линейный размер
    //-------------------------------------------------------------------------------
    // Создать линейный размер
    // ---
    void CreateLineDimension()
    {
      // Получить контейнер условных обозначений
      ISymbols2DContainer symbCont = GetSymbols2DContainer();

      if ( symbCont != null )
      {
        // Получить коллекцию линейных размеров
        ILineDimensions dimCol = symbCont.LineDimensions;

        if ( dimCol != null )
        {
          //Добавить объект в коллекцию
          ILineDimension newDim = dimCol.Add();

          if ( newDim != null )
          {
            // Координаты первой точки привязки размера
            newDim.X1 = 50;
            newDim.Y1 = 150;
            // Координаты второй точки привязки размера
            newDim.X2 = 100;
            newDim.Y2 = 150;
            // Положение размерной линии
            newDim.X3 = 75;
            newDim.Y3 = 180;
            // Тип ориентации линейного размера
            newDim.Orientation = Kompas6Constants.ksLineDimensionOrientationEnum.ksLinDHorizontal;
            // Применить параметры
            newDim.Update();    
          }
        }
      }
    }


    //-------------------------------------------------------------------------------
    // Выдать сообщение с параметрами линейного размера
    // ---
    void GetLineDimensionParam( ILineDimension dim )
    {
      if ( dim != null )
      {
        // Координаты первой точки привязки размера
        double x1 = dim.X1;
        double y1 = dim.Y1;
        // Координаты второй точки привязки размера
        double x2 = dim.X2;
        double y2 = dim.Y2;
        // Выдать сообщение с параметрами объекта
        string buf = string.Format( LoadString("IDS_COORDS"), x1, y1, x2, y2 );
        kompas.ksMessage( buf );
      }
    }


    //-------------------------------------------------------------------------------
    // Навигация по коллекции линейных размеров
    // ---
    void LineDimensionNavigation()
    {
      // Получить контейнер условных обозначений
      ISymbols2DContainer symbCont = GetSymbols2DContainer();

      if ( symbCont != null )
      {
        // Получить коллекцию линейных размеров
        ILineDimensions dimCol = symbCont.LineDimensions;

        if ( dimCol != null )
        {
          for ( long i = 0; i < dimCol.Count; i++ )
          {
            // Получить объект из коллекции по индексу
            ILineDimension lineDim = dimCol.get_LineDimension( i );

            if ( lineDim != null )
            {
              int dimRef = lineDim.Reference;

              if ( doc2D != null )
              {
                // Подсветить объект
                doc2D.ksLightObj( dimRef, 1/*включить подсветку*/ );
                // Выдать сообщение с параметрами линейного размера
                GetLineDimensionParam( lineDim );
                // Погасить объект
                doc2D.ksLightObj( dimRef, 0/*выключить подсветку*/ );
              }
            }
          }
        }
      }
    }


    //-------------------------------------------------------------------------------
    // Изменить параметры у линейного размера
    // ---
    void ChangeLineDimensionParam( ref ILineDimension dim )
    {
      if ( dim != null )
      {
        // Получить интерфейс параметров размера
        IDimensionParams dimPar = (IDimensionParams)dim;
    		
        if ( dimPar != null )
        {
          // Направление полки - влево
          dimPar.ShelfDirection = Kompas6Constants.ksShelfDirectionEnum.ksLSLeft;
          // Типы стрелок - Угол 90 град.
          dimPar.ArrowType1 = Kompas6Constants.ksArrowEnum.ksRightAngle;
          dimPar.ArrowType2 = Kompas6Constants.ksArrowEnum.ksRightAngle;
        }
      		
        // Координаты точки начала полки
        dim.ShelfX = 60;
        dim.ShelfY = 200;
      }
    }


    //-------------------------------------------------------------------------------
    // Изменить текст у линейного размера
    // ---
    void ChangeLineDimensionText( ref ILineDimension dim )
    {
      if ( dim != null )
      {
        // Получить интерфейс текста размера
        IDimensionText dimText = (IDimensionText)dim;
    		
        if ( dimText != null )
        {
          // Текст в рамке
          dimText.Rectangle = true;
          // Подчеркнуть текст
          dimText.Underline = true;
      			
          // Получить интерфейс текста до значения размера
          ITextLine prefix = dimText.Prefix;
      			
          // Изменить текст
          if ( prefix != null )
            prefix.Str = LoadString( "IDS_DIM" );
        }
      }
    }


    //-------------------------------------------------------------------------------
    // Изменить линейный размер
    // ---
    void LineDimChange( ref ILineDimension dim )
    {
      if ( dim != null )
      {
        // Изменить параметры размера
        ChangeLineDimensionParam( ref dim );
        // Изменить текст размера
        ChangeLineDimensionText( ref dim );
        // Применить изменения
        dim.Update();
      }
    }


    //-------------------------------------------------------------------------------
    // Редактирование линейного размера
    // ---
    void EditLineDimension()
    {
      // Структура параметров запроса к системе
      ksRequestInfo info = (ksRequestInfo)kompas.GetParamStruct( 
                           (int)Kompas6Constants.StructType2DEnum.ko_RequestInfo );
      info.Init();
      info.commandsString = LoadString( "IDS_COMMAND1" );
      double x = 0, y = 0;
  	
      if ( doc2D != null && info != null )
      {
        // Указать объект в документе
        if ( doc2D.ksCursorEx(info, ref x, ref y, null, null) != 0 )
        {
          // Найти объект по указанным координатам
          int pObj = doc2D.ksFindObj( x, y, doc2D.ksGetCursorLimit() );
    		
          // Если это линейный размер, редактировать
          if ( doc2D.ksGetObjParam(pObj, 0, 0) == ldefin2d.LDIMENSION_OBJ/*линейный размер*/ )
          {
            // Получить интерфейс линейного размера из референса объекта
            ILineDimension lineDim = (ILineDimension)kompas.TransferReference( pObj, 
                                      doc2D.reference );
            // Изменить линейный размер
            LineDimChange( ref lineDim );
          }
          else
            kompas.ksMessage( LoadString( "IDS_NOTDIM") );
        }
      } 
    }
    #endregion


    #region Радиальный размер
    //-------------------------------------------------------------------------------
    // Создать радиальный размер
    // ---
    bool CreateRadDimension( ref IRadialDimension dim )
    {
      bool res = false;

      if ( dim != null )
      {
        // Структура параметров запроса к системе
        ksRequestInfo info = (ksRequestInfo)kompas.GetParamStruct( 
                             (int)Kompas6Constants.StructType2DEnum.ko_RequestInfo );
        info.Init();
        info.commandsString = LoadString( "IDS_COMMAND2" );
        double x = 0, y = 0;
    		
        if ( doc2D != null && info != null )
        {
          // Указать объект в документе
          if ( doc2D.ksCursorEx( info, ref x, ref y, null, null) != 0 )
          {
            // Найти объект по указанным координатам
            int pObj = doc2D.ksFindObj( x, y, doc2D.ksGetCursorLimit() );
            // Структура параметров окружности
            CircleParam par = (CircleParam)kompas.GetParamStruct( 
                              (int)Kompas6Constants.StructType2DEnum.ko_CircleParam );
            par.Init();

            // Является ли объект окружностью
            if ( doc2D.ksGetObjParam(pObj, par, ldefin2d.ALLPARAM ) == ldefin2d.CIRCLE_OBJ )
            {
              // Получить интерфейс графического объекта из референса
              IDrawingObject circle = (IDrawingObject)kompas.TransferReference( pObj, doc2D.reference );
    				
              if ( circle != null )
              {
                // Задать базовый объект размера
                dim.BaseObject = circle;
                // Задать координаты центра размера, совпадающие с центром указанной окружности
                dim.Xc = par.xc;
                dim.Yc = par.yc;
                // Задать радиус размера, совпадающий с радиусом указанной окружности
                dim.Radius = par.rad;
                // Тип размера - от центра
                dim.DimensionType = true;
                // Угол наклона размерной линии
                dim.Angle = 30;
                // Применить параметры
                res = dim.Update();
              }
              else
                kompas.ksMessage( LoadString("IDS_NOOBJ") );
            }
            else
              kompas.ksMessage( LoadString("IDS_NOTCREATE1") );
          } 
        }
      }
      return res;
    }


    //-------------------------------------------------------------------------------
    // Изменить параметры и текст у радиального размера
    // ---
    void ChangeRadialDimensionParamText( ref IRadialDimension dim )
    {
      if ( dim != null )
      {
        // Получить интерфейс параметров размера
        IDimensionParams dimPar = (IDimensionParams)dim;

        if ( dimPar != null )
        {
          // Направление полки - вправо
          dimPar.ShelfDirection = Kompas6Constants.ksShelfDirectionEnum.ksLSRight;
          // Угол наклона полки
          dimPar.ShelfAngle = 180;
          // Длина полки
          dimPar.ShelfLength = 30;
          // Тип стрелки - засечка
          dimPar.ArrowType1 = Kompas6Constants.ksArrowEnum.ksNotch;
        }

        // Получить интерфейс текстов размера
        IDimensionText dimText = (IDimensionText)dim;

        if ( dimText != null )
        {
          // Значок перед номиналом - радиус
          dimText.Sign = 3;

          // Получить интерфейс текста единицы измерения
          ITextLine unit = dimText.Unit;

          // Задать текстовую строку
          if ( unit != null )
            unit.Str = LoadString( "IDS_UNIT" );
        }
      }
    }


    //-------------------------------------------------------------------------------
    // Редактировать радиальный размер
    // ---
    void EditRadialDimension( ref IRadialDimension dim )
    {
      if ( dim != null )
      {
        // Тип размера - не от центра
        dim.DimensionType = false;
        // Изменить параметры и текст размера
        ChangeRadialDimensionParamText( ref dim );
        // Применить изменения
        dim.Update();
      }
    }


    //-------------------------------------------------------------------------------
    // Создание и редактирование радиального размера
    // ---
    void RadialDimensionWork()
    {
      // Получить контейнер условных обозначений
      ISymbols2DContainer symbCont = GetSymbols2DContainer();

      if ( symbCont != null )
      {
        // Получить коллекцию радиальных размеров
        IRadialDimensions radCol = symbCont.RadialDimensions;

        if ( radCol != null )
        {
          // Добавить радиальный размер в коллекцию
          IRadialDimension radDim = radCol.Add();

          if ( radDim != null )
          {
            // Создать радиальный размер
            bool create = CreateRadDimension( ref radDim );
            // Получить референс размера
            int dimRef = radDim.Reference;

            if ( create && kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
            {
              // Получить объект из коллекции по референсу
              IRadialDimension radDimension = radCol.get_RadialDimension(dimRef);
              // Редактировать радиальный размер
              EditRadialDimension( ref radDimension );
            }
          }
        }
      }
    }
    #endregion


    #region Диаметральный размер
    //-------------------------------------------------------------------------------
    // Создать диаметральный размер
    // ---
    bool CreateDiamrtralDimension( ref IDiametralDimension dim )
    {
      bool res = false;

      if ( dim != null )
      {
        // Структура параметров запроса к системе
        ksRequestInfo info = (ksRequestInfo)kompas.GetParamStruct( 
                             (int)Kompas6Constants.StructType2DEnum.ko_RequestInfo );
        info.Init();
        info.commandsString = LoadString( "IDS_COMMAND2" );
        double x = 0, y = 0;
    		
        // Указать объект в документе
        if ( doc2D.ksCursorEx( info, ref x, ref y, null, null) != 0 )
        {
          // Найти объект по указанным координатам
          int pObj = doc2D.ksFindObj( x, y, doc2D.ksGetCursorLimit() );
          // Структура параметров окружности
          CircleParam par = (CircleParam)kompas.GetParamStruct( 
                            (int)Kompas6Constants.StructType2DEnum.ko_CircleParam );
          par.Init();
    			
          // Является ли объект окружностью
          if ( doc2D.ksGetObjParam(pObj, par, ldefin2d.ALLPARAM ) == ldefin2d.CIRCLE_OBJ )
          {
            // Получить интерфейс графического объекта из референса
            IDrawingObject circle = (IDrawingObject)kompas.TransferReference( pObj, 
                                     doc2D.reference );
    				
            if ( circle != null )
            {
              // Задать базовый объект размера
              dim.BaseObject = circle;
              // Задать координаты центра размера, совпадающие с центром указанной окружности
              dim.Xc = par.xc;
              dim.Yc = par.yc;
              // Задать радиус размера, совпадающий с радиусом указанной окружности
              dim.Radius = par.rad;
              // Тип размера - полная размерная линия
              dim.DimensionType = true;
              // Угол наклона размерной линии
              dim.Angle = 45;
              // Применить параметры
              res = dim.Update();
            }
            else
              kompas.ksMessage( LoadString("IDS_NOOBJ") );
          }
          else
            kompas.ksMessage( LoadString("IDS_NOTCREATE2") );
        }
      }
      return res;
    }


    //-------------------------------------------------------------------------------
    // Изменить параметры и текст у диаметрального размера
    // ---
    void ChangeDiametralDimensionParamText( ref IDiametralDimension dim )
    {
      if ( dim != null )
      {
        // Получить интерфейс параметров размера
        IDimensionParams dimPar = (IDimensionParams)dim;
    		
        if ( dimPar != null )
        {
          // Направление полки - вправо
          dimPar.ShelfDirection = Kompas6Constants.ksShelfDirectionEnum.ksLSRight;
          // Длина полки
          dimPar.ShelfLength = 20;
          // Тип стрелок - точка
          dimPar.ArrowType1 = Kompas6Constants.ksArrowEnum.ksPoint;
          dimPar.ArrowType2 = Kompas6Constants.ksArrowEnum.ksPoint;
        }
    		
        // Получить интерфейс текстов размера
        IDimensionText dimText = (IDimensionText)dim;
    		
        if ( dimText != null )
        {
          // Значок перед номиналом - диаметр
          dimText.Sign = 1;
          // Подчеркнуть
          dimText.Underline = true;
        			
          // Получить интерфейс текста единицы измерения
          ITextLine suffix = dimText.Suffix;
        			
          // Задать текстовую строку
          if ( suffix != null )
            suffix.Str = LoadString( "IDS_UNIT" );
        }
      }
    }


    //-------------------------------------------------------------------------------
    // Редактировать диаметральный размер
    // ---
    void EditDiametralDimension( ref IDiametralDimension dim )
    {
      if ( dim != null )
      {
        // Тип размера - размерная линия с обрывом
        dim.DimensionType = false;
        // Угол наклона размерной линии
        dim.Angle = 90;
        // Изменить параметры и текст размера
        ChangeDiametralDimensionParamText( ref dim );
        // Применить изменения
        dim.Update();
      }
    }

   
    //-------------------------------------------------------------------------------
    // Создание и редактирование диаметрального размера
    // ---
    void DiamrtralDimensionWork()
    {
      // Получить контейнер условных обозначений
      ISymbols2DContainer symbCont = GetSymbols2DContainer();

      if ( symbCont != null )
      {
        // Получить коллекцию диаметральных размеров
        IDiametralDimensions diamCol = symbCont.DiametralDimensions;
    		
        if ( diamCol != null )
        {
          // Добавить диаметральный размер в коллекцию
          IDiametralDimension diamDim = diamCol.Add();
    			
          if ( diamDim != null )
          {
            // Создать диаметральный размер
            bool create = CreateDiamrtralDimension( ref diamDim );
            // Получить референс размера
            int dimRef = diamDim.Reference;
    				
            if ( create && kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
            {
              // Получить объект из коллекции по референсу
              IDiametralDimension diamDimension = diamCol.get_DiametralDimension( dimRef );
              // Редактировать диаметральный размер
              EditDiametralDimension( ref diamDimension );
            }
          }
        }
      }
    }
    #endregion


    #region Угловой размер
    //-------------------------------------------------------------------------------
    // Создать угловой размер
    // ---
    bool CreateAngleDimension( ref IAngleDimension dim )
    {
      bool res = false;

      if ( dim != null && doc2D != null )
      {
        // Построение размера на объектах, указанных в чертеже
        // Структура параметров запроса к системе
        /*ksRequestInfo info = (ksRequestInfo)kompas.GetParamStruct( (int)Kompas6Constants.StructType2DEnum.ko_RequestInfo );
        info.Init();
        info.commandsString = LoadString( "IDS_COMMAND3" );
        double x = 0, y = 0;
	
        // Указать первый базовый объект в документе
        if ( doc2D.ksCursorEx(info, ref x, ref y, null, null) != 0 )
        {
          // Найти объект по указанным координатам
          int pObj1 = doc2D.ksFindObj( x, y, 1 );
          // Получить тип первого объекта
          int type1 = doc2D.ksGetObjParam( pObj1, 0, 0 );

          if ( type1 == ldefin2d.LINESEG_OBJ || type1 == ldefin2d.POLYLINE_OBJ || type1 == ldefin2d.RECTANGLE_OBJ )
          {
            // Подсветить указанный объект
            doc2D.ksLightObj( pObj1, 1 );
            info.commandsString = LoadString( "IDS_COMMAND4" );
			
            // Указать второй базовый объект в документе
            if ( doc2D.ksCursorEx(info, ref x, ref y, null, null) != 0 )
            {
              // Найти объект по указанным координатам
              int pObj2 = doc2D.ksFindObj( x, y, 1 );
              // Получить тип второго объекта
              int type2 = doc2D.ksGetObjParam( pObj2, 0, 0 );

              if ( type1 == ldefin2d.LINESEG_OBJ || type1 == ldefin2d.POLYLINE_OBJ || type1 == ldefin2d.RECTANGLE_OBJ )
              {
                // Подсветить указанный объект
                doc2D.ksLightObj( pObj2, 1 );
                // Получить интерфейс графического объекта из референса
                IDrawingObject baseObj1 = (IDrawingObject)kompas.TransferReference( pObj1, doc2D.reference );
                IDrawingObject baseObj2 = (IDrawingObject)kompas.TransferReference( pObj2, doc2D.reference );

                if ( baseObj1 != null && baseObj2 != null )
                {
                  // Задать базовые объекты размера
                  dim.BaseObject1 = baseObj1;
                  dim.BaseObject2 = baseObj2;
                  // Радиус дуги
                  dim.Radius = 20;
                  // Применить параметры
                  res = dim.Update();
                  // Убрать подсветку объекта
                  doc2D.ksLightObj( pObj2, 0 );
                }
                else
                  kompas.ksMessage( LoadString("IDS_NOOBJS") );
					
                // Убрать подсветку объекта
                doc2D.ksLightObj( pObj1, 0 );
              }
              else
                kompas.ksMessage( LoadString("IDS_NOTCREATE3") );
            }
          }
          else
            kompas.ksMessage( LoadString("IDS_NOTCREATE3") );
        }*/
        // Построение размера на двух отрезках		
        // Построить первый опорный объект
        int line1 = doc2D.ksLineSeg( 80, 120, 100, 120, 1 );
        IDrawingObject baseObj1 = (IDrawingObject)kompas.TransferReference( line1, 
                                   doc2D.reference );
        // Построить второй опорный объект
        int line2 = doc2D.ksLineSeg( 80, 120, 120, 160, 1 );
        IDrawingObject baseObj2 = (IDrawingObject)kompas.TransferReference( line2, 
                                   doc2D.reference );
		
        if ( baseObj1 != null && baseObj2 != null )
        {
          // Задать опорные объекты
          dim.BaseObject1 = baseObj1;
          dim.BaseObject2 = baseObj2;
          // Координаты центра
          dim.Xc = 90;
          dim.Yc = 150;
          // Радиус дуги
          dim.Radius = 25;
          // Начальный угол  размерной дуги
          dim.Angle1 = 45;
          // Конечный угол  размерной дуги
          dim.Angle2 = 30;
          // Координаты точки выхода первой выносной линии
          dim.X1 = 80;
          dim.Y1 = 120;
          // Координаты точки выхода второй выносной линии
          dim.X2 = 100;
          dim.Y2 = 160;
          // Тип размера - на минимальный (острый) угол
          dim.DimensionType = Kompas6Constants.ksAngleDimTypeEnum.ksADMinAngle;
			
          // Включить полку
          IDimensionParams dimPar = (IDimensionParams)dim;
			
          if ( dimPar != null )
            // Направление полки - вправо
            dimPar.ShelfDirection = Kompas6Constants.ksShelfDirectionEnum.ksLSRight;
					
          // Точка начала полки
          dim.ShelfX = 160;
          dim.ShelfY = 140;
          // Направление размерной дуги - против часовой стрелки
          dim.Direction = false;
          // Начало ножки или точка на дуге
          dim.X3 = 140;
          dim.Y3 = 120;
          // Применить параметры
          res = dim.Update();
        }
      }
      return res;
    }


    //-------------------------------------------------------------------------------
    // Редактирование углового размера
    // ---
    void EditAngleDimension( ref IAngleDimension dim )
    {
      if ( dim != null )
      {
        // Тип размера - На максимальный (тупой) угол
        dim.DimensionType = Kompas6Constants.ksAngleDimTypeEnum.ksADMaxAngle;
        // Направление размерной дуги  - по часовой стрелке
        dim.Direction = true;
        // Точка начала полки
        dim.ShelfX = 150;
        dim.ShelfY = 170;
        // Радиус дуги
        dim.Radius = 40;
        // Применить изменения
        dim.Update();
      }
    }


    //-------------------------------------------------------------------------------
    // Создание и редактирование углового размера
    // ---
    void AngleDimensionWork()
    {
	    // Получить контейнер условных обозначений
	    ISymbols2DContainer symbCont = GetSymbols2DContainer();
    	
	    if ( symbCont != null )
	    {
		    // Получить коллекцию угловых размеров
		    IAngleDimensions angCol = symbCont.AngleDimensions;
    		
		    if ( angCol != null )
		    {
			    // Добавить угловой размер в коллекцию
			    IAngleDimension angDim = angCol.Add( Kompas6Constants.DrawingObjectTypeEnum.
                                               ksDrADimension );
    			
			    if ( angDim != null )
			    {
				    // Создать угловой размер
				    bool create = CreateAngleDimension( ref angDim );
				    // Получить референс размера
				    int dimRef = angDim.Reference;
    				
            if ( create && kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
				    {
					    // Получить объект из коллекции по референсу
					    IAngleDimension angDimension = angCol.get_AngleDimension( dimRef );
					    // Редактировать угловой размер
					    EditAngleDimension( ref angDimension );
				    }
			    }
		    }
	    }
    }
    #endregion


    #region Размер дуги окружности
    //-------------------------------------------------------------------------------
    // Создание размера дуги окружности
    // ---
    bool CreateArcDimension( ref IArcDimension dim )
    {
      bool res = false;

	    if ( dim != null && doc2D != null )
	    {
		    // Структура параметров запроса к системе
        ksRequestInfo info = (ksRequestInfo)kompas.GetParamStruct( 
                             (int)Kompas6Constants.StructType2DEnum.ko_RequestInfo );
		    info.Init();
		    info.commandsString = LoadString( "IDS_COMMAND5" );
		    double x = 0, y = 0;
    		
		    // Указать объект в документе
        if ( doc2D.ksCursorEx(info, ref x, ref y, null, null) != 0 )
		    {
			    // Найти объект по указанным координатам
          int pObj = doc2D.ksFindObj( x, y, 1 );
			    // Подсветить указанный объект
			    doc2D.ksLightObj( pObj, 1 );
			    // Структура параметров дуги
			    ArcByPointParam par = (ArcByPointParam)kompas.GetParamStruct( 
                                (int)Kompas6Constants.StructType2DEnum.ko_ArcByPointParam );
    			par.Init();

			    // Является ли объект дугой
          if ( doc2D.ksGetObjParam(pObj, par, ldefin2d.POINT_ARC_PARAM ) == ldefin2d.ARC_OBJ )
			    {
				    // Получить интерфейс графического объекта из референса
				    IDrawingObject arc = (IDrawingObject)kompas.TransferReference( pObj, 
                                  doc2D.reference );
    				
				    if ( arc != null )
				    {
					    // Базовый объект
					    dim.BaseObject = arc;
					    // Координаты центра
					    dim.Xc = par.xc;
					    dim.Yc = par.yc;
					    // Координаты первой точки дуги
					    dim.X1 = par.x1;
					    dim.Y1 = par.y1;
					    // Координаты второй точки дуги
					    dim.X2 = par.x2;
					    dim.Y2 = par.y2;
    					
					    info.commandsString = LoadString( "IDS_COMMAND6" );					
					    // Указать положение размерной линии
              if ( doc2D.ksCursorEx(info, ref x, ref y, null, null) != 0 )
					    {
						    dim.X3 = x;
						    dim.Y3 = y;
					    }

					    // Направление размерной дуги
					    dim.Direction = true;
					    // Тип размера - параллельные выносные линии
					    dim.DimensionType = false;
					    // Указатель от текста к дуге
					    dim.TextPointer = true;
					    // Применить параметры
					    res = dim.Update();
				    }
				    else
              kompas.ksMessage( LoadString("IDS_NOOBJ") );
			    }
			    else
            kompas.ksMessage( LoadString("IDS_NOTCREATE4") );

			    // Убрать подсветку объекта
          doc2D.ksLightObj( pObj, 0 );
		    }
	    }
	    return res;
    }


    //-------------------------------------------------------------------------------
    // Редактирование размера дуги окружности
    // ---
    void EditArcDimension( ref IArcDimension dim )
    {
      if ( dim != null )
      {
        // Тип размера - выносные линии от центра
        dim.DimensionType = true;
        // Указатель от текста к дуге
        dim.TextPointer = false;
        // Направление размерной дуги - поменять на противоположное
        dim.Direction = !dim.Direction;

        // Получить интерфейс параметров размера
        IDimensionParams dimPar = (IDimensionParams)dim;

        if ( dimPar != null )
        {
          // Тип стрелок
          dimPar.ArrowType1 = Kompas6Constants.ksArrowEnum.ksLeaderPoint;
          dimPar.ArrowType2 = Kompas6Constants.ksArrowEnum.ksLeaderPoint;
          // Направление полки - вниз
          dimPar.ShelfDirection = Kompas6Constants.ksShelfDirectionEnum.ksLSDown;
          // Угол наклона полки
          dimPar.ShelfAngle = 180;
        }

        dim.Update();

        // Получить интерфейс текста размера
        IDimensionText dimText = (IDimensionText)dim;

        if ( dimText != null )
        {
          // Отключить автоматическое определение номинального значения
          dimText.AutoNominalValue = false;
          // Установить новое номинальное значение
          dimText.NominalValue = 50;
        			
          // Получить интерфейс текста единицы измерения
          ITextLine unit = dimText.Unit;
        			
          // Задать текстовую строку
          if ( unit != null )
            unit.Str = LoadString( "IDS_UNIT" );
        }
        // Применить изменения
        dim.Update();
      }
    }

  
    //-------------------------------------------------------------------------------
    // Создание и редактирование размера дуги окружности
    // ---
    void ArcDimensionWork()
    {
	    // Получить контейнер условных обозначений
	    ISymbols2DContainer symbCont = GetSymbols2DContainer();
    	
	    if ( symbCont != null )
	    {
		    // Получить коллекцию размеров дуги окружности
		    IArcDimensions arcCol = symbCont.ArcDimensions;
    		
		    if ( arcCol != null )
		    {
			    // Добавить размер дуги окружности в коллекцию
			    IArcDimension arcDim = arcCol.Add();
    			
			    if ( arcDim != null )
			    {
				    // Создать размер дуги окружности
				    bool create = CreateArcDimension( ref arcDim );
				    // Получить референс размера
				    int dimRef = arcDim.Reference;
    				
				    if ( create && kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
				    {
					    // Получить объект из коллекции по референсу
					    IArcDimension arcDimension = arcCol.get_ArcDimension( dimRef );
					    // Редактировать размер дуги окружности
					    EditArcDimension( ref arcDimension );
				    }
			    }
		    }
	    }	
    }
    #endregion


    #region Линейный размер с обрывом
    //-------------------------------------------------------------------------------
    // Создание линейного размера с обрывом
    // ---
    bool CreateBreakLineDim( ref IBreakLineDimension dim )
    {
      bool res = false;

	    if ( dim != null && doc2D != null )
	    {
		    // Структура параметров запроса к системе
        ksRequestInfo info = (ksRequestInfo)kompas.GetParamStruct( 
                             (int)Kompas6Constants.StructType2DEnum.ko_RequestInfo );
        info.Init();
		    info.commandsString = LoadString( "IDS_COMMAND7" );
		    double x = 0, y = 0;
    		
		    // Указать объект в документе
		    if ( doc2D.ksCursorEx(info, ref x, ref y, null, null) != 0 )
		    {
			    // Найти объект по указанным координатам
			    int pObj = doc2D.ksFindObj( x, y, doc2D.ksGetCursorLimit() );
			    // Структура параметров отрезка
			    LineSegParam par = (LineSegParam)kompas.GetParamStruct( 
                             (int)Kompas6Constants.StructType2DEnum.ko_LineSegParam );
          par.Init();
    			
			    // Является ли объект отрезком
			    if ( doc2D.ksGetObjParam(pObj, par, ldefin2d.ALLPARAM ) == ldefin2d.LINESEG_OBJ )
			    {
				    // Подсветить указанный объект
				    doc2D.ksLightObj( pObj, 1 );
				    // Получить интерфейс графического объекта из референса
				    IDrawingObject lineSeg = (IDrawingObject)kompas.TransferReference( pObj, 
                                      doc2D.reference );
    				
				    if ( lineSeg != null )
				    {
					    // Базовый объект
					    dim.BaseObject = lineSeg;
					    info.commandsString = LoadString( "IDS_COMMAND6" );					
    					
					    // Указать положение размерной линии
              if ( doc2D.ksCursorEx(info, ref x, ref y, null, null) != 0 )
					    {
						    // Координаты первой точки привязки - координаты первого конца отрезка
						    dim.X1 = par.x1;
						    dim.Y1 = par.y1;
						    // Координаты второй точки привязки - координаты второго конца отрезка
						    dim.X2 = par.x2;
						    dim.Y2 = par.y2;
						    // Положение размерной линии - указанные координаты
						    dim.X3 = x;
						    dim.Y3 = y;
					    }
					    // Применить параметры
					    res = dim.Update();
				    }
				    else
              kompas.ksMessage( LoadString("IDS_NOOBJ") );
			    }
			    else
				    kompas.ksMessage( LoadString("IDS_NOTCREATE5") );
    			
			    // Убрать подсветку объекта
          doc2D.ksLightObj( pObj, 0 );
		    }
	    }
	    return res;
    }


    //-------------------------------------------------------------------------------
    // Редактирование линейного размера с обрывом
    // ---
    void EditBreakLineDim( ref IBreakLineDimension dim )
    {
	    if ( dim != null )
	    {
		    // Получить интерфейс параметров размера
		    IDimensionParams dimPar = (IDimensionParams)dim;

		    if ( dimPar != null )
		    {
			    // Тип стрелки			
			    dimPar.ArrowType1 = Kompas6Constants.ksArrowEnum.ksLeftNotch;
		    }

		    // Получить интерфейс текста размера
		    IDimensionText dimText = (IDimensionText)dim;

		    if ( dimText != null )
		    {
			    // Получить текст перед размером
			    ITextLine prefix = dimText.Prefix;

			    if ( prefix != null )
				    // Изменить текст
				    prefix.Str = "Размер: ";

			    // Получить текст номинального значения
			    ITextLine nominal = dimText.NominalText;
    			
			    if ( nominal != null )
				    // Изменить текст
				    nominal.Str = LoadString( "IDS_NOMINAL" );
		    }
		    // Применить изменения
		    dim.Update();
	    }
    }


    //-------------------------------------------------------------------------------
    // Создание и редактирование линейного размера с обрывом
    // ---
    void BreakLineDimensionWork()
    {
	    // Получить контейнер условных обозначений
	    ISymbols2DContainer symbCont = GetSymbols2DContainer();
    	
	    if ( symbCont != null )
	    {
		    // Получить коллекцию линейных размеров с обрывом
		    IBreakLineDimensions breakCol = symbCont.BreakLineDimensions;
    		
		    if ( breakCol != null )
		    {
			    // Добавить линейный размер с обрывом в коллекцию
			    IBreakLineDimension breakDim = breakCol.Add();
    			
			    if ( breakDim != null )
			    {
				    // Создать линейный размер с обрывом
				    bool create = CreateBreakLineDim( ref breakDim );
				    // Получить референс размера
				    int dimRef = breakDim.Reference;
    				
				    if ( create && kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
				    {
					    // Получить объект из коллекции по референсу
					    IBreakLineDimension breakDimension = breakCol.get_BreakLineDimension( dimRef );
					    // Редактировать линейный размер с обрывом
					    EditBreakLineDim( ref breakDimension );
				    }
			    }
		    }
	    }	
    }
    #endregion


    #region Радиальный размер с изломом
    //-------------------------------------------------------------------------------
    // Создание радиального размера с изломом
    // ---
    bool CreateBreakRadialDim( ref IBreakRadialDimension dim )
    {
      bool res = false;

	    if ( dim != null )
	    {
		    // Структура параметров запроса к системе
        ksRequestInfo info = (ksRequestInfo)kompas.GetParamStruct( 
                             (int)Kompas6Constants.StructType2DEnum.ko_RequestInfo );
		    info.Init();
		    info.commandsString = LoadString( "IDS_COMMAND2" );
		    double x = 0, y = 0;
    		
		    // Указать объект в документе
        if ( doc2D.ksCursorEx(info, ref x, ref y, null, null) != 0 )
		    {
			    // Найти объект по указанным координатам
          int pObj = doc2D.ksFindObj( x, y, doc2D.ksGetCursorLimit() );
			    // Структура параметров окружности
			    CircleParam par = (CircleParam)kompas.GetParamStruct( 
                            (int)Kompas6Constants.StructType2DEnum.ko_CircleParam );
    			
			    // Является ли объект окружностью
			    if ( doc2D.ksGetObjParam(pObj, par, ldefin2d.ALLPARAM ) == ldefin2d.CIRCLE_OBJ )
			    {
				    // Получить интерфейс графического объекта из референса
				    IDrawingObject circle = (IDrawingObject)kompas.TransferReference( pObj, doc2D.reference );
    				
				    if ( circle != null )
				    {
					    // Задать базовый объект размера
					    dim.BaseObject = circle;
					    // Задать координаты центра размера, совпадающие с центром указанной окружности
					    dim.Xc = par.xc;
					    dim.Yc = par.yc;
					    // Задать радиус размера, совпадающий с радиусом указанной окружности
					    dim.Radius = par.rad;
					    // Угол наклона размерной линии
					    dim.Angle = 45;
					    // Длина излома
					    dim.BreakLength = 3;
					    // Положение размерной надписи
					    dim.TextOnLine = Kompas6Constants.ksDimensionTextPosEnum.ksDimTextParallelOnLine;
					    // Применить параметры
					    res = dim.Update();
				    }
				    else
					    kompas.ksMessage( LoadString("IDS_NOOBJ") );
			    }
			    else
				    kompas.ksMessage( LoadString("IDS_NOTCREATE1") );
		    } 
	    }
      return res;
    }


    //-------------------------------------------------------------------------------
    // Редактировать радиальный размер с изломом
    // ---
    void EditBreakRadialDim( ref IBreakRadialDimension dim )
    {
	    if ( dim != null )
	    {
		    // Угол наклона размерной линии
		    dim.Angle = 90;
		    // Длина излома
		    dim.BreakLength = 1;
		    // Положение размерной надписи
		    dim.TextOnLine = Kompas6Constants.ksDimensionTextPosEnum.ksDimTextParallelInCut;
    	
		    // Получить интерфейс текста размера
		    IDimensionText dimText = (IDimensionText)dim;

		    if ( dimText != null )
		    {
			    // Убрать значок радиуса
			    dimText.Sign = 0;
		    }
		    // Применить изменения
		    dim.Update();
	    }
    }


    //-------------------------------------------------------------------------------
    // Создание и редактирование радиального размера с изломом
    // ---
    void BreakRadialDimensionWork()
    {
	    // Получить контейнер условных обозначений
	    ISymbols2DContainer symbCont = GetSymbols2DContainer();
    	
	    if ( symbCont != null )
	    {
		    // Получить коллекцию радиальных размеров с изломом
		    IBreakRadialDimensions breakCol = symbCont.BreakRadialDimensions;
    		
		    if ( breakCol != null )
		    {
			    // Добавить радиальный размер с изломом в коллекцию
			    IBreakRadialDimension breakDim = breakCol.Add();
    			
			    if ( breakDim != null )
			    {
				    // Создать радиальный размер с изломом
				    bool create = CreateBreakRadialDim( ref breakDim );
				    // Получить референс размера
				    int dimRef = breakDim.Reference;
    				
            if ( create && kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
				    {
					    // Получить объект из коллекции по референсу
					    IBreakRadialDimension breakDimension = breakCol.get_BreakRadialDimension( dimRef );
					    // Редактировать радиальный размер с изломом
					    EditBreakRadialDim( ref breakDimension );
				    }
			    }
		    }
	    }		
    }
    #endregion


    #region Угловой размер с обрывом
    //-------------------------------------------------------------------------------
    // Создание углового размера с обрывом
    // ---
    bool CreateBreakAngleDim( ref IBreakAngleDimension dim )
    {
      bool res = false;

	    if ( dim != null && doc2D != null )
	    {
		    // Построить первый опорный объект
        int line1 = doc2D.ksLineSeg( 80, 120, 120, 160, 1 );
		    IDrawingObject baseObj1 = (IDrawingObject)kompas.TransferReference( line1, 
                                   doc2D.reference );
		    // Построить второй опорный объект
		    int line2 = doc2D.ksLineSeg( 80, 120, 100, 120, 1 );
		    IDrawingObject baseObj2 = (IDrawingObject)kompas.TransferReference( 
                                   line2, doc2D.reference );
    		
		    if ( baseObj1 != null && baseObj2 != null )
		    {
			    // Задать опорные объекты
			    dim.BaseObject1 = baseObj1;
			    dim.BaseObject2 = baseObj2;
			    // Координаты центра
			    dim.Xc = 80;
			    dim.Yc = 120;
			    // Радиус дуги
			    dim.Radius = 20;
			    // Начальный угол  размерной дуги
			    dim.Angle1 = 45;
			    // Конечный угол  размерной дуги
			    dim.Angle2 = 30;
			    // Координаты точки выхода первой выносной линии
			    dim.X1 = 120;
			    dim.Y1 = 160;
			    // Координаты точки выхода второй выносной линии
			    dim.X2 = 100;
			    dim.Y2 = 160;
			    // Тип размера - на минимальный (острый) угол
			    dim.DimensionType = Kompas6Constants.ksAngleDimTypeEnum.ksADMinAngle;
			    // Направление размерной дуги - против часовой стрелки
			    dim.Direction = true;
			    // Начало ножки или точка на дуге
			    dim.X3 = 165;
			    dim.Y3 = 140;
			    // Применить параметры
			    res = dim.Update();
		    }	
	    }
      return res;
    }


    //-------------------------------------------------------------------------------
    // Редактирование углового размера с обрывом
    // ---
    void EditBreakAngleDim( ref IBreakAngleDimension dim )
    {
	    if ( dim != null )
	    {
		    // Включить полку
		    IDimensionParams dimPar = (IDimensionParams)dim;
    		
		    if ( dimPar != null )
			    // Направление полки - вправо
			    dimPar.ShelfDirection = Kompas6Constants.ksShelfDirectionEnum.ksLSRight;
    		
		    // Точка начала полки
		    dim.ShelfX = 167;
		    dim.ShelfY = 177;

		    // Получить интерфейс текстов размера
		    IDimensionText dimText = (IDimensionText)dim;

		    if ( dimText != null )
		    {
			    // Выравнивание размерной надписи
			    dimText.TextAlign = Kompas6Constants.ksDimensionTextAlignEnum.ksDimALowerBoundary;
			    // Текст в круглых скобках
			    dimText.Brackets = Kompas6Constants.ksDimensionTextBracketsEnum.ksDimBrackets;
		    }
		    // Применить изменения
		    dim.Update();
	    }
    }


    //-------------------------------------------------------------------------------
    // Создание и редактирование углового размера с обрывом
    // ---
    void BreakAngleDimensionWork()
    {
	    // Получить контейнер условных обозначений
	    ISymbols2DContainer symbCont = GetSymbols2DContainer();
    	
	    if ( symbCont != null )
	    {
		    // Получить коллекцию угловых размеров с обрывом
		    IAngleDimensions breakCol = symbCont.AngleDimensions;
    		
		    if ( breakCol != null )
		    {
			    // Добавить угловой размер с обрывом в коллекцию
			    IBreakAngleDimension breakDim = (IBreakAngleDimension)breakCol.Add( 
                                          Kompas6Constants.DrawingObjectTypeEnum.
                                          ksDrABreakDimension );
    			
			    if ( breakDim != null )
			    {
				    // Создать угловой размер с обрывом
				    bool create = CreateBreakAngleDim( ref breakDim );
				    // Получить референс размера
				    int dimRef = breakDim.Reference;
    				
            if ( create && kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
				    {
					    // Получить объект из коллекции по референсу
					    IBreakAngleDimension breakDimension = (IBreakAngleDimension)breakCol.
                                                     get_AngleDimension( dimRef );
					    // Редактировать угловой размер с обрывом
					    EditBreakAngleDim( ref breakDimension );
				    }
			    }
		    }
	    }	
    }
    #endregion


    #region Размер высоты
    //-------------------------------------------------------------------------------
    // Создать размер высоты
    // ---
    bool CreateHeightDimension( ref IHeightDimension dim )
    {
      bool res = false;

	    if ( dim != null )
	    {
		    // Структура параметров запроса к системе
        ksRequestInfo info = (ksRequestInfo)kompas.GetParamStruct( (int)Kompas6Constants.StructType2DEnum.ko_RequestInfo );
		    info.Init();
		    info.commandsString = LoadString( "IDS_COMMAND8" );
		    double x = 0, y = 0;
    		
		    // Указать точку нулевого уровня
        if ( doc2D.ksCursorEx(info, ref x, ref y, null, null) != 0 )
		    {
			    info.commandsString = LoadString( "IDS_COMMAND9" );
			    double x1 = 0, y1 = 0;

			    // Указать точку измеряемого уровня
          if ( doc2D.ksCursorEx(info, ref x1, ref y1, null, null) != 0 )
			    {
				    info.commandsString = LoadString( "IDS_COMMAND10" );
				    double x2 = 0, y2 = 0;
    				
				    // Указать положение размерной надписи
            if ( doc2D.ksCursorEx(info, ref x2, ref y2, null, null) != 0 )
				    {
					    // Точка нулевого уровня
					    dim.X = x;
					    dim.Y = y;
					    // Точка измеряемого уровня
					    dim.X1 = x1;
					    dim.Y1 = y1;
					    // Положение размерной надписи
					    dim.X2 = x2;
					    dim.Y2 = y2;
					    // Применить параметры
					    res = dim.Update();
				    }
			    }
		    }
	    }
      return res;
    }


    //-------------------------------------------------------------------------------
    // Редактировать размер высоты
    // ---
    void EditHeightDimension( ref IHeightDimension dim )
    {
	    if ( dim != null )
	    {
		    // Тип размера
		    dim.DimensionType = Kompas6Constants.ksHeightDimTypeEnum.ksHDTopViewLeader;

		    // Получить интерфейс текста размера
		    IDimensionText dimText = (IDimensionText)dim;

		    if ( dimText != null )
		    {
			    // Условный значок - квадрат
			    dimText.Sign = 2;
			    // Получить текст номинального значения
			    ITextLine nominal = dimText.NominalText;
    			
			    if ( nominal != null )
				    // Изменить текст
				    nominal.Str = LoadString( "IDS_NOMINAL" );
		    }
		    // Применить изменения
		    dim.Update();
	    }
    }


    //-------------------------------------------------------------------------------
    // Создание и редактирование размера высоты
    // ---
    void HeightDimensionWork()
    {
	    // Получить контейнер условных обозначений
	    ISymbols2DContainer symbCont = GetSymbols2DContainer();
    	
	    if ( symbCont != null )
	    {
		    // Получить коллекцию размеров высоты
		    IHeightDimensions heightCol = symbCont.HeightDimensions;
    		
		    if ( heightCol != null )
		    {
			    // Добавить размер высоты в коллекцию
			    IHeightDimension heightDim = heightCol.Add();
    			
			    if ( heightDim != null )
			    {
				    // Создать размер высоты
				    bool create = CreateHeightDimension( ref heightDim );
				    // Получить референс размера
				    int dimRef = heightDim.Reference;
    				
            if ( create && kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
				    {
					    // Получить объект из коллекции по референсу
					    IHeightDimension heightDimension = heightCol.get_HeightDimension( dimRef );
					    // Редактировать размер высоты
					    EditHeightDimension( ref heightDimension );
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
