using Kompas6API5;

using System;
using Microsoft.Win32;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Kompas6Constants;
using Kompas6Constants3D;
using KAPITypes;

using reference = System.Int32;

namespace Steps.NET
{
	// Класс Step3d3 - Операции, оси и плоскости
  [ClassInterface(ClassInterfaceType.AutoDual)]
	public class Step3d3
	{
		private KompasObject kompas;
		private ksDocument3D doc;

		// Имя библиотеки
		[return: MarshalAs(UnmanagedType.BStr)] public string GetLibraryName()
		{
			return "Step3d3 - Операции, оси и плоскости";
		}
		

		// Головная функция библиотеки
		public void ExternalRunCommand([In] short command, [In] short mode, [In, MarshalAs(UnmanagedType.IDispatch)] object kompas_)
		{
			kompas = (KompasObject) kompas_;
			if (kompas != null)
			{
				doc = (ksDocument3D)kompas.ActiveDocument3D();
				if (doc != null && doc.reference != 0)
				{
					switch (command)
					{
						case 1 : ConstrAxisOperations();	break; // Конструктивная ось операции
						case 2 : ConstrAxis2Point();		break; // Конструктивная ось по двум точкам
						case 3 : ConstrAxisEdge();			break; // Конструктивная ось, проходящая через ребро
						case 4 : CreateConstrElem();		break; // Создание смещенной плоскости, оси по двум плоскостям и плоскости под углом к заданной
						case 5 : ConstrPlane3Point();		break; // Плоскость через три вершины
					}
				}
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
					result = "Конструктивная ось операции";
					command = 1;
					break;
				case 2:
					result = "Конструктивная ось по двум точкам";
					command = 2;
					break;
				case 3:
					result = "Конструктивная ось, проходящая через ребро";
					command = 3;
					break;
				case 4:
					result = "Смещенная плоскость, ось по двум плоскостям, плоскость под углом к другой пл-ти";
					command = 4;
					break;
				case 5:
					result = "Плоскость через три вершины";
					command = 5;
					break;
				case 6:
					itemType = 3;			//ENDMENU
					command = -1;
					break;
			}

			return result;
		}


		// Конструктивная ось операции
		void ConstrAxisOperations()
		{
      ksPart part = (ksPart)doc.GetPart((short)Part_Type.pTop_Part);	// новый компонент
			if (part != null)
			{
				ksEntity entitySketch = (ksEntity)part.NewEntity((short)Obj3dType.o3d_sketch);
				if (entitySketch != null)
				{
					// интерфейс свойств эскиза
					ksSketchDefinition sketchDef = (ksSketchDefinition)entitySketch.GetDefinition();
					if (sketchDef != null)
					{
						// получим интерфейс базовой плоскости XOY
						ksEntity basePlane = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);
						sketchDef.SetPlane(basePlane);	// установим плоскость XOY базовой для эскиза
						entitySketch.Create();			// создадим эскиз

						// интерфейс редактора эскиза
						ksDocument2D sketchEdit = (ksDocument2D)sketchDef.BeginEdit();
						sketchEdit.ksCircle(20, 0, 10, 1);
						sketchEdit.ksLineSeg(0, 0, 0, 5, 3);
						sketchDef.EndEdit();			// завершение редактирования эскиза

						ksEntity entityRotate = (ksEntity)part.NewEntity((short)Obj3dType.o3d_bossRotated);
						bool res = false;
						if (entityRotate != null)
						{
							ksBossRotatedDefinition rotateDef = (ksBossRotatedDefinition)entityRotate.GetDefinition(); // интерфейс базовой операции вращения
							if (rotateDef != null)
							{
								ksRotatedParam rotproperty = (ksRotatedParam)rotateDef.RotatedParam();
								if (rotproperty != null)
								{
									rotproperty.direction = (short)Direction_Type.dtBoth;
									rotproperty.toroidShape = false;
								}

								rotateDef.SetThinParam(true, (short)Direction_Type.dtBoth, 1, 1);	// тонкая стенка в два направления
								rotateDef.SetSideParam(true, 180);
								rotateDef.SetSideParam(false, 180);
								rotateDef.SetSketch(entitySketch);	// эскиз операции вращения
								res = entityRotate.Create();		// создать операцию
							}
						}
						if (res)
						{
							// создадим операцию вырезания по сечениям
							ksEntity entityAxisOperation = (ksEntity)part.NewEntity((short)Obj3dType.o3d_axisOperation);
							if (entityAxisOperation != null)
							{
								ksAxisOperationsDefinition axisOperation = (ksAxisOperationsDefinition)entityAxisOperation.GetDefinition();
								if (axisOperation != null)
								{
									axisOperation.SetOperation(entityRotate);
									entityAxisOperation.Create();	// создать операцию
								}
							}
							kompas.ksMessage("Ось операции");
						}
						else
							kompas.ksMessage("Ошибка создания операции");
					}
				}
			}
		}


		// Конструктивная ось по двум точкам
		void ConstrAxis2Point()
		{
      ksPart part = (ksPart)doc.GetPart((short)Part_Type.pTop_Part);	// новый компонент
			if (part != null)
			{
				ksEntity entitySketch = (ksEntity)part.NewEntity((short)Obj3dType.o3d_sketch);
				if (entitySketch != null)
				{
					// интерфейс свойств эскиза
					ksSketchDefinition sketchDef = (ksSketchDefinition)entitySketch.GetDefinition();
					if (sketchDef != null)
					{
						// получим интерфейс базовой плоскости XOY
						ksEntity basePlane = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);
						sketchDef.SetPlane(basePlane);	// установим плоскость XOY базовой для эскиза
						entitySketch.Create();			// создадим эскиз

						// интерфейс редактора эскиза
						ksDocument2D sketchEdit = (ksDocument2D)sketchDef.BeginEdit(); 
						// введем новый эскиз - квадрат
						sketchEdit.ksLineSeg(50,  50, -50,  50, 1);
						sketchEdit.ksLineSeg(50, -50, -50, -50, 1); 
						sketchEdit.ksLineSeg(50, -50,  50,  50, 1);
						sketchEdit.ksLineSeg(-50, -50, -50,  50, 1);
						sketchDef.EndEdit();	// завершение редактирования эскиза                
				
						ksEntity entityExtr = (ksEntity)part.NewEntity((short)Obj3dType.o3d_bossExtrusion);
						if (entityExtr != null)
						{
							// интерфейс свойств базовой операции выдавливания
							ksBossExtrusionDefinition extrusionDef = (ksBossExtrusionDefinition)entityExtr.GetDefinition();	// интерфейс базовой операции выдавливания
							if (extrusionDef != null)
							{
								extrusionDef.directionType = (short)Direction_Type.dtNormal;			// направление выдавливания
								extrusionDef.SetSideParam(true /*прямое направление*/, (short)End_Type.etBlind /*строго на глубину*/, 20, 0, false); 
								extrusionDef.SetThinParam(true, (short)Direction_Type.dtBoth, 20, 20);	// тонкая стенка в два направления
								extrusionDef.SetSketch(entitySketch);									// эскиз операции выдавливания
								entityExtr.Create();													// создать операцию
							}
						}
					}
				}

				ksEntityCollection entityCollection = (ksEntityCollection)part.EntityCollection((short)Obj3dType.o3d_vertex);

				if (entityCollection != null && entityCollection.GetCount() != 0)
				{
					// создадим ось по двум точкам
					ksEntity entityAxis2Point = (ksEntity)part.NewEntity((short)Obj3dType.o3d_axis2Points); 
					if (entityAxis2Point != null)
					{
						ksAxis2PointsDefinition axis2Point = (ksAxis2PointsDefinition)entityAxis2Point.GetDefinition();
						if (axis2Point != null)
						{
							axis2Point.SetPoint(1, entityCollection.GetByIndex(0));
							axis2Point.SetPoint(2, entityCollection.GetByIndex(entityCollection.GetCount() - 1));
							entityAxis2Point.Create();
						}
					}
					kompas.ksMessage("Ось через две точки");
				}
			}
		}


		// Конструктивная ось, проходящая через ребро
		void ConstrAxisEdge()
		{
      ksPart part = (ksPart)doc.GetPart((short)Part_Type.pTop_Part);	// новый компонент
			if (part != null)
			{
				ksEntity entitySketch = (ksEntity)part.NewEntity((short)Obj3dType.o3d_sketch);
				if (entitySketch != null)
				{
					// интерфейс свойств эскиза
					ksSketchDefinition sketchDef = (ksSketchDefinition)entitySketch.GetDefinition();
					if (sketchDef != null)
					{
						// получим интерфейс базовой плоскости XOY
						ksEntity basePlane = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);
						sketchDef.SetPlane(basePlane);	// установим плоскость XOY базовой для эскиза
						entitySketch.Create();			// создадим эскиз

						// интерфейс редактора эскиза
						ksDocument2D sketchEdit = (ksDocument2D)sketchDef.BeginEdit(); 
						// введем новый эскиз - квадрат
						sketchEdit.ksLineSeg(50,  50, -50,  50, 1);
						sketchEdit.ksLineSeg(50, -50, -50, -50, 1); 

						sketchEdit.ksLineSeg(50, -50,  50,  50, 1);
						sketchEdit.ksLineSeg(-50, -50, -50,  50, 1);
						sketchDef.EndEdit();	// завершение редактирования эскиза                
				
						ksEntity entityExtr = (ksEntity)part.NewEntity((short)Obj3dType.o3d_bossExtrusion);
						if (entityExtr != null)
						{
							// интерфейс свойств базовой операции выдавливания
							ksBossExtrusionDefinition extrusionDef = (ksBossExtrusionDefinition)entityExtr.GetDefinition(); // интерфейс базовой операции выдавливания
							if (extrusionDef != null)
							{
								extrusionDef.directionType = (short)Direction_Type.dtNormal;			// направление выдавливания
								extrusionDef.SetSideParam(true /*прямое направление*/, (short)End_Type.etBlind /*строго на глубину*/, 20, 0, false); 
								extrusionDef.SetThinParam(true, (short)Direction_Type.dtBoth, 20, 20);	// тонкая стенка в два направления
								extrusionDef.SetSketch(entitySketch);									// эскиз операции выдавливания
								entityExtr.Create();													// создать операцию
							}
						}
					}
				}

				ksEntityCollection entityCollection = (ksEntityCollection)part.EntityCollection((short)Obj3dType.o3d_edge);

				if (entityCollection != null && entityCollection.GetCount() > 1)
				{
					// создадим ось через грань
					ksEntity entityAxisEdge = (ksEntity)part.NewEntity((short)Obj3dType.o3d_axisEdge); 
					if (entityAxisEdge != null)
					{
						ksAxisEdgeDefinition axisEdge = (ksAxisEdgeDefinition)entityAxisEdge.GetDefinition();
						if (axisEdge != null)
						{
							axisEdge.SetEdge(entityCollection.GetByIndex(0));
							entityAxisEdge.Create();
						}
					}
					kompas.ksMessage("Ось через грань");

					// создадим еще ось через грань
					ksEntity entityAxisEdge2 = (ksEntity)part.NewEntity((short)Obj3dType.o3d_axisEdge); 
					if (entityAxisEdge2 != null)
					{
						ksAxisEdgeDefinition axisEdge = (ksAxisEdgeDefinition)entityAxisEdge2.GetDefinition();
						if (axisEdge != null)
						{
							axisEdge.SetEdge(entityCollection.GetByIndex(1));
							entityAxisEdge2.Create();
						}
					}
					kompas.ksMessage("Другая ось через грань");
				}
			}
		}


		// Плоскость через три вершины
		void ConstrPlane3Point()
		{
      ksPart part = (ksPart)doc.GetPart((short)Part_Type.pTop_Part);	// новый компонент
			if (part != null)
			{
				ksEntity entitySketch = (ksEntity)part.NewEntity((short)Obj3dType.o3d_sketch);
				if (entitySketch != null)
				{
					// интерфейс свойств эскиза
					ksSketchDefinition sketchDef = (ksSketchDefinition)entitySketch.GetDefinition();
					if (sketchDef != null)
					{
						// получим интерфейс базовой плоскости XOY
						ksEntity basePlane = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);
						sketchDef.SetPlane(basePlane);	// установим плоскость XOY базовой для эскиза
						entitySketch.Create();			// создадим эскиз

						// интерфейс редактора эскиза
						ksDocument2D sketchEdit = (ksDocument2D)sketchDef.BeginEdit(); 
						// введем новый эскиз - квадрат
						sketchEdit.ksLineSeg(50,  50, -50,  50, 1);
						sketchEdit.ksLineSeg(50, -50, -50, -50, 1); 

						sketchEdit.ksLineSeg(50, -50,  50,  50, 1);
						sketchEdit.ksLineSeg(-50, -50, -50,  50, 1);
						sketchDef.EndEdit();	// завершение редактирования эскиза                
				
						ksEntity entityExtr = (ksEntity)part.NewEntity((short)Obj3dType.o3d_bossExtrusion);
						if (entityExtr != null)
						{
							// интерфейс свойств базовой операции выдавливания
							ksBossExtrusionDefinition extrusionDef = (ksBossExtrusionDefinition)entityExtr.GetDefinition(); // интерфейс базовой операции выдавливания
							if (extrusionDef != null)
							{
								extrusionDef.directionType = (short)Direction_Type.dtNormal;			// направление выдавливания
								extrusionDef.SetSideParam(true /*прямое направление*/, (short)End_Type.etBlind /*строго на глубину*/, 20, 30, false); 
								extrusionDef.SetThinParam(true, (short)Direction_Type.dtBoth, 10, 10);	// тонкая стенка в два направления
								extrusionDef.SetSketch(entitySketch);									// эскиз операции выдавливания
								entityExtr.Create();													// создать операцию
							}
						}
					}
				}

				ksEntityCollection entityCollection = (ksEntityCollection)part.EntityCollection((short)Obj3dType.o3d_vertex);
				if (entityCollection.GetCount() > 2)
				{
					// Плоскость через три вершины
					ksEntity entityConstrPlane3Point = (ksEntity)part.NewEntity((short)Obj3dType.o3d_plane3Points); 
					if (entityConstrPlane3Point != null)
					{
						ksPlane3PointsDefinition constrPlane3Point = (ksPlane3PointsDefinition)entityConstrPlane3Point.GetDefinition();
						if (constrPlane3Point != null)
						{
							constrPlane3Point.SetPoint(1, entityCollection.GetByIndex(0));
							constrPlane3Point.SetPoint(2, entityCollection.GetByIndex(1));
							constrPlane3Point.SetPoint(3, entityCollection.GetByIndex(2));
							entityConstrPlane3Point.Create();
						}
					}
					kompas.ksMessage("Плоскость через три вершины");
				}
			}
		}


		// Создание смещенной плоскости, оси по двум плоскостям и плоскости под углом к заданной
		void CreateConstrElem()
		{
      ksPart part = (ksPart)doc.GetPart((short)Part_Type.pTop_Part);	// новый компонент
			if (part != null)
			{
				ksEntity entity = (ksEntity)part.NewEntity((short)Obj3dType.o3d_planeOffset);
				if (entity != null)
				{
					// интерфейс свойств смещенной плоскости
					ksPlaneOffsetDefinition offsetDef = (ksPlaneOffsetDefinition)entity.GetDefinition();
					if (offsetDef != null)
					{
						offsetDef.offset = 150;		// расстояние от базовой плоскости
						ksEntity basePlane = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);
						basePlane.name = "XOY";		// название для плоскости
						basePlane.Update();			// обновить параметры
      
						offsetDef.SetPlane(basePlane);			// базовая плоскость
						offsetDef.direction = false;			// направление смещения от базовой плоскости
						entity.name = "Смещенная плоскость";	// имя для смещенной плоскости
						entity.Create();						// создать смещенную плоскость 
				
						kompas.ksMessage("Изменим параметры смещенной плоскости");
     
						offsetDef.offset = 50;				// изменим расстояние до базовой плоскости

						// возьмем другую базовую плоскость
						basePlane = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeYOZ);
						basePlane.name = "YOZ";
						basePlane.Update();					// обновить параметры

						offsetDef.direction = true;			// изменим направления смещения относительно базовой плоскости
						offsetDef.SetPlane(basePlane); 
						entity.Update();					// обновить параметры

						// возьмем другую базовую плоскость
						basePlane = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);
						basePlane.name = "XOY";
				
						kompas.ksMessage("На пересечении плоскостей построим ось");

						// Ось на пересечении двух плоскостей
						ksEntity entityAxis = (ksEntity)part.NewEntity((short)Obj3dType.o3d_axis2Planes);
						if (entityAxis != null)
						{
							ksAxis2PlanesDefinition axis2PlanesDef = (ksAxis2PlanesDefinition)entityAxis.GetDefinition();
							if (axis2PlanesDef != null)
							{
								axis2PlanesDef.SetPlane(1, entity);			// Базовая плоскость 1
								axis2PlanesDef.SetPlane(2, basePlane);		// Базовая плоскость 2
								entityAxis.name = "Ось по двум плоскостям";	// имя для оси
								entityAxis.Create();						// создаем ось  

								kompas.ksMessage("Поменяем одну из базовых плоскостей для построения оси");
						
								// возьмем другую базовую плоскость
								basePlane = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOZ);
								basePlane.name = "XOZ";
						
								axis2PlanesDef.SetPlane(2, basePlane);	// Базовая плоскость 2
								entityAxis.Update();

								kompas.ksMessage("Через смещенную плоскость и построенную ось \n проведем плоскость под углом 45");

								ksEntity entityAnglePlane = (ksEntity)part.NewEntity((short)Obj3dType.o3d_planeAngle); 
								if (entityAnglePlane != null)
								{
									// интерфейс свойств плоскости под углом к другой плоскости
									ksPlaneAngleDefinition planeAngleDef = (ksPlaneAngleDefinition)entityAnglePlane.GetDefinition();
									if (planeAngleDef != null)
									{
										planeAngleDef.angle = 45;			// угол наклона к базовой плоскости
										planeAngleDef.SetPlane(entity);		// базовая плоскость
										planeAngleDef.SetAxis(entityAxis);	// базовая ось
										entityAnglePlane.name = "Плоскость под углом к другой плоскости";
										entityAnglePlane.Create();			// создать плоскость под углом 

										kompas.ksMessage("Изменим одну из базовых плоскостей");

										planeAngleDef.SetPlane(basePlane);	// базовая плоскость
										entityAnglePlane.Update();			// обновить параметры плоскости
									}
								}
							}
						}
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
