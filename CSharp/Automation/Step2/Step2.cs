using Kompas6API5;


using System;
using Microsoft.Win32;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Globalization;
using KAPITypes;
using Kompas6Constants;

namespace Steps.NET
{
	// Класс Step2 - Математика
	
	// 1. Пересечь прямые                      - Intersect2Line
	// 2. Пересечь кривые                      - Intersect2Curve
	// 3. Пересечь отрезок и дугу              - IntersectLineSegArc
	// 4. Касательная из точки                 - TanLinePointCircle
	// 5. Касательная под углом                - TanLineAngCircle
	// 6. Поворот точки                        - RotatePoint
	// 7. Симметрия точки                      - SymmetryPoint
	// 8. Сопрягающие окружности к двум прямым - Couplin2Lines
	// 9. Перепендикуляр                       - Perpendicular

	[ClassInterface(ClassInterfaceType.AutoDual)]
	public class Step2
	{
		private KompasObject kompas;
		private ksDocument2D doc;
		private ksMathematic2D mat;

		// Отрисовка точек пересечения в документе doc по
		// присланному массиву и выдача пользователю их координат
		private void DrawPointByArray(ksDynamicArray arr)
		{
			if (arr != null)
			{
				// Создать интерфейс параметров математической точки
				ksMathPointParam par = (ksMathPointParam) kompas.GetParamStruct((short) StructType2DEnum.ko_MathPointParam);

				if (par != null)
				{
					// Интерфейс создан
					for (int i = 0; i < arr.ksGetArrayCount(); i ++)
					{
						arr.ksGetArrayItem(i, par);
						doc.ksPoint(par.x, par.y, 5);
						string buf = string.Format("x = {0:.##} y = {1:.##}", par.x, par.y);
						kompas.ksMessage(buf);
					}
				}
			}
		}


		// Пересечь прямые
		private void IntersectLines()
		{
			ksDynamicArray arr = (ksDynamicArray) kompas.GetDynamicArray(ldefin2d.POINT_ARR);
			if (arr != null)
			{
				doc.ksLine(10, 10, 0);
				doc.ksLine(15, 5, 90);
				mat.ksIntersectLinLin(10, 10, 0, 15, 5, 90, arr);
				DrawPointByArray(arr);
				arr.ksDeleteArray();
			}
		}


		// Пересечь кривые
		private void IntersectCurves()
		{
			ksDynamicArray arr = (ksDynamicArray) kompas.GetDynamicArray(ldefin2d.POINT_ARR);
			if (arr != null)
			{
				doc.ksBezier(0, 0);
					doc.ksPoint(20, 0, 0);
					doc.ksPoint(10, 20, 0);
					doc.ksPoint(20, 40, 0);
					doc.ksPoint(30, 20, 0);
					doc.ksPoint(20, 0, 0);
				int pp1 = doc.ksEndObj();

				doc.ksBezier(0, 0);
					doc.ksPoint(0, 20, 0);
					doc.ksPoint(20, 10, 0);
					doc.ksPoint(40, 20, 0);
					doc.ksPoint(20, 30, 0);
					doc.ksPoint(0, 20, 0);
				int pp2 = doc.ksEndObj();

				mat.ksIntersectCurvCurv(pp1, pp2, arr);
				DrawPointByArray(arr);
				arr.ksDeleteArray();
			}
		}


		// Пересечь отрезок и дугу
		private void IntercectLineSegAndArc()
		{
			ksDynamicArray arr = (ksDynamicArray) kompas.GetDynamicArray(ldefin2d.POINT_ARR);

			if (arr != null)
			{
				doc.ksLineSeg(0, 40, 100, 40, 1);					// Отрисовка отрезка
				doc.ksArcByPoint(50, 40, 20, 30, 40, 70, 40, 1, 1);	// Отрисовка дуги по центру и конечным точкам
				double a1 = mat.ksAngle(50, 40, 30, 40);			// Начальный угол дуги
				double a2 = mat.ksAngle(50, 40, 70, 40);			// Конечный угол дуги
    
				// Получить координаты точек пересечения отрезка и дуги
				// Первая точка отрезка (0, 40), Вторая точка отрезка (100, 40),
				// Центр дуги (50, 40), Радиус дуги 20

				mat.ksIntersectLinSArc (0, 40, 100, 40, 50, 40, 20, a1, a2, 1, arr);
			    DrawPointByArray(arr);	// Отрисовка точек пересечения
    
				arr.ksDeleteArray();
			}
		}


		// Касательная из точки
		private void TanFromPoint()
		{
			ksDynamicArray arr = (ksDynamicArray) kompas.GetDynamicArray(ldefin2d.POINT_ARR);
			ksMathPointParam par = (ksMathPointParam) kompas.GetParamStruct((short) StructType2DEnum.ko_MathPointParam);

			if ((arr != null) & (par != null))
			{
				doc.ksPoint(10, 50, 3);			// Отрисовка точки
				doc.ksCircle(50, 10, 40, 1);	// Отрисовка окружности
    
				// Получить точки касания окружности и прямой, проходящей через заданную точку
				// Координаты внешней точки (10, 50), Координаты центра (50, 10),
				// радиус окружности 40

				mat.ksTanLinePointCircle(10, 50, 50, 10, 40, arr);
				DrawPointByArray(arr);		// Отрисовка точек пересечения
         
				// Отрисовка касательных
				for (int i = 0; i < arr.ksGetArrayCount(); i ++)
				{
					arr.ksGetArrayItem(i, par);	// Параметры текущей точки
					doc.ksLine(10, 50, mat.ksAngle(10, 50, par.x, par.y));
				}
      
				arr.ksDeleteArray();
			}
		}


		// Касательная под углом
		private void TanToAngle()
		{
			ksDynamicArray arr = (ksDynamicArray) kompas.GetDynamicArray(ldefin2d.POINT_ARR);
			ksMathPointParam par = (ksMathPointParam) kompas.GetParamStruct((short) StructType2DEnum.ko_MathPointParam);

			if ((arr != null) & (par != null))
			{
				doc.ksLineSeg(0, 40, 100, 40, 1);	// Отрисовка отрезка
				doc.ksCircle(50, 10, 40, 1);		// Отрисовка окружности
    
				// Получить точки касания окружности и прямой, проходящей под заданным углом
				// Координаты центра (50, 10), радиус окружности 40, Угол касательной прямой 45

				mat.ksTanLineAngCircle(50, 10, 40, 45, arr);
				DrawPointByArray(arr);			// Отрисовка точек пересечения
         
				// Отрисовка касательных
				for (int i = 0; i < arr.ksGetArrayCount(); i ++)
				{
					arr.ksGetArrayItem(i, par);	// Параметры текущей точки
					doc.ksLine(par.x, par.y, 45);
				}
       
			    arr.ksDeleteArray();
			}
		}


		// Поворот точки
		private void RotatePoint()
		{
			double x = 0;
			double y = 0;	//Результат поворота

			doc.ksPoint(60, 50, 3);	//Отрисовка точки
			doc.ksPoint(50, 50, 2);	//Отрисовка точки

			mat.ksRotate(60, 50, 50, 50, 180,  out x, out y);	//Поворот точки

			doc.ksPoint(x, y, 5);	// Отрисовка результирующей точки
			// Результат поворота
			kompas.ksMessage(string.Format("x = {0:.##} y = {1:.##}", x, y));
		}


		//Симметрия точки
		private void SymmetryPoint()
		{
			double x = 0;	// Результат симметрии точки
			double y = 0;

			doc.ksPoint(30, 60, 3);				// Отрисовка точки
			doc.ksLineSeg(0, 50, 60, 50, 3);	// Отрисовка отрезка
  
			// Получить координаты точки, симметричной относительно заданной оси
			mat.ksSymmetry(30, 60, 0, 50, 60, 50, out x, out y);
  
			doc.ksPoint(x, y, 5);	// Отрисовка результирующей точки
			// Результат симметрии
			kompas.ksMessage(string.Format("x = {0:.##} y = {1:.##}", x, y));
		}


		// Сопрягающие окружности к двум прямым
		private void CouplingTwoLines()
		{
			// Создать интерфейс массива координат точек сопряжения
			ksCON iCON = (ksCON) kompas.GetParamStruct((short) StructType2DEnum.ko_CON);

			if (iCON != null)	// Интерфейс создан
			{
				doc.ksLine(100, 100, 45);	// Отрисовка прямых - Первая прямая
				doc.ksLine(100, 100, -45);	// Вторая прямая
    
				// Получить параметры окружностей, касательной к двум прямым
				// Радиус сопряжения 20

				mat.ksCouplingLineLine(100, 100, 45, 100, 100, -45, 20, iCON);

				// Отрисовка сопрягающихся окружностей и точек касания
				for (int i = 0; i < iCON.GetCount(); i ++)
				{
					doc.ksCircle(iCON.GetXc(i), iCON.GetYc(i), 20, 2);
					doc.ksPoint(iCON.GetX1(i), iCON.GetY1(i), i);
					doc.ksPoint(iCON.GetX2(i), iCON.GetY2(i), i);
				}
  
				// Результат сопряжения
				string msg = string.Format("count = {0:.##} con[0].x1 = {1:.##} con[0].y1 = {2:.##} con[0].x2 = {3:.##} con[0].y2 = {4:.##} ...",
					iCON.GetCount(), 
					iCON.GetX1(0), 
					iCON.GetY1(0), 
					iCON.GetX2(0), 
					iCON.GetY2(0));
				kompas.ksMessage(msg);
			}
		}


		// Перепендикуляр
		private void BuildPerpend()
		{
			doc.ksPoint(50, 50, 2);				// Отрисовка точки
			doc.ksLineSeg(60, 10, 100, 10, 1);	// Отрисовка отрезка
  
			double x = 0;		// Точка пересечения отрезка и перпендикуляра
			double y = 0;

			// Координаты точки пересечения отрезка и перпендикуляра
			// Координаты произвольной внешней точки (50, 50)
			// Координаты первой точки отрезка (60, 10), Координаты второй точки отрезка (100, 10)

			mat.ksPerpendicular(50, 50, 60, 10, 100, 10, out x, out y);
			// Отрисовка перпендикуляра
			doc.ksLine(50, 50, mat.ksAngle(50, 50, x, y));
			// Отрисовка точки пересечения отрезка
			doc.ksPoint(x, y, 5);
  
			// Результат расчета перпендикуляра
			kompas.ksMessage(string.Format("x = {0:.##} y = {1:.##}", x, y));
		}


		[return: MarshalAs(UnmanagedType.BStr)] public string GetLibraryName()
		{
			return "Step2 - Использованиe математических функций";
		}


		[return: MarshalAs(UnmanagedType.BStr)] public string ExternalMenuItem(short number, ref short itemType, ref short command)
		{
			string result = string.Empty;
			itemType = 1; // "MENUITEM"
			switch (number)
			{
				case 1:
					result = "Пересечь прямые";
					command = 1;
					break;
				case 2:
					result = "Пересечь кривые";
					command = 2;
					break;
				case 3:
					result = "Пересечь отрезок и дугу";
					command = 3;
					break;
				case 4:
					result = "Касательная из точки";
					command = 4;
					break;
				case 5:
					result = "Касательная под углом";
					command = 5;
					break;
				case 6:
					result = "Поворот точки";
					command = 6;
					break;
				case 7:
					result = "Симметрия точки";
					command = 7;
					break;
				case 8:
					result = "Сопрягающие окружности к двум прямым";
					command = 8;
					break;
				case 9:
					result = "Перпендикуляр";
					command = 9;
					break;
				case 10:
					command = -1;
					itemType = 3; // "ENDMENU"
					break;
			}
            return result;
		}

	
		public void ExternalRunCommand([In] short command, [In] short mode, [In, MarshalAs(UnmanagedType.IDispatch)] object kompas_)
		{
			kompas = (KompasObject) kompas_;

			if (kompas == null)
				return;

			doc = (ksDocument2D) kompas.ActiveDocument2D();

			if (doc == null)
				return;

			mat = (ksMathematic2D) kompas.GetMathematic2D();

			if (mat == null)
				return;

			switch (command)
			{
				case 1:	IntersectLines();			break; // пересечь прямые
				case 2: IntersectCurves();			break; // пересечь кривые
				case 3: IntercectLineSegAndArc();	break; // пересечь отрезок и дугу
				case 4:	TanFromPoint();				break; // касательная из точки
				case 5:	TanToAngle();				break; // касательная под углом
				case 6:	RotatePoint();				break; // поворот точки
				case 7:	SymmetryPoint();			break; // симметрия точки
				case 8:	CouplingTwoLines();			break; // сопрягающие окружности к двум прямым
				case 9:	BuildPerpend();				break; // перепендикуляр
			}

			kompas.ksMessageBoxResult();
		}


		public object ExternalGetResourceModule()
		{
			return Assembly.GetExecutingAssembly().Location;
		}


		public int ExternalGetToolBarId(short barType, short index)
		{
			int result = 0;

			if (barType == 0)
			{
				result = -1;
			}
			else
			{
				switch (index)
				{
					case 1:
						result = 3001;
						break;
					case 2:
						result = -1;
						break;
				}
			}

			return result;
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
