
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
	// Класс Step3d1 - Объекты 3D
  [ClassInterface(ClassInterfaceType.AutoDual)]
	public class Step3d1
	{
		private KompasObject kompas;
		private ksDocument3D doc;

		// Имя библиотеки
		[return: MarshalAs(UnmanagedType.BStr)] public string GetLibraryName()
		{
			return "Step3d1 - Объекты 3D";
		}
		

		// Головная функция библиотеки
		public void ExternalRunCommand([In] short command, [In] short mode, [In, MarshalAs(UnmanagedType.IDispatch)] object kompas_)
		{
			kompas = (KompasObject) kompas_;
			if (kompas != null)
			{
				doc = (ksDocument3D)kompas.ActiveDocument3D();
				if (doc == null || doc.reference == 0)
				{
					doc = (ksDocument3D)kompas.Document3D();
					doc.Create(true, true);

					doc.author = "Ethereal";
					doc.comment = "3D Steps - Step3d1";
					doc.UpdateDocumentParam();
				}

				switch (command)
				{
					case 1 : CreateExtrusion();			break; // базовая операция выдавливания
					case 2 : OperationRotated();		break; // Операции вращения
					case 3 : OperationLoft();			break; // Операции по сечениям
					case 4 : CreateFilletAndChamfer();	break; // создание фаски и скругления
					case 5 : CreateNextOper();			break; // Операции : оболочка, уклон, сечение плоскостью, сечение эскизом
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
					result = "Операции выдавливания";
					command = 1;
					break;
				case 2:
					result = "Операции вращения";
					command = 2;
					break;
				case 3:
					result = "Оперции по сечениям";
					command = 3;
					break;
				case 4:
					result = "Создание фаски и скругления";
					command = 4;
					break;
				case 5:
					result = "Операции : оболочка, уклон, сечение плоскостью, сечение эскизом";
					command = 5;
					break;
				case 6:
					itemType = 3;			//ENDMENU
					command = -1;
					break;
			}

			return result;
		}


		// Удалить все объекты из текущего эскиза
		void ClearCurrentSketch(ksDocument2D sketchEdit)
		{
			// создаим итератор и удалим все существующие объекты в эскизе          
			ksIterator iter = (ksIterator)kompas.GetIterator();
			if (iter != null)
			{
				if (iter.ksCreateIterator(ldefin2d.ALL_OBJ, 0))
				{
					reference rf;
					if ((rf = iter.ksMoveIterator("F")) != 0) 
					{
						// сместить указатель на первый элемент в списке
						// в цикле сместить указатель на следующий элемент в списке пока не дойдем до последнего
						do
						{
							if (sketchEdit.ksExistObj(rf) == 1)
								sketchEdit.ksDeleteObj(rf);	// если объект существует удалить его 
						}
						while ((rf = iter.ksMoveIterator("N")) != 0); 
					}
					iter.ksDeleteIterator();	// удалим итератор
				}
			}
		}


		// Оперция выдавливания, работа с экизом
		void CreateExtrusion()
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
						sketchDef.angle = 45;			// угол поворота эскиза
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
								extrusionDef.directionType = (short)Direction_Type.dtNormal;         // направление выдавливания
								extrusionDef.SetSideParam(true,	// прямое направление
									(short)End_Type.etBlind,	// строго на глубину
									200, 0, false);
								extrusionDef.SetThinParam(true, (short)Direction_Type.dtBoth, 10, 10); // тонкая стенка в два направления
								extrusionDef.SetSketch(entitySketch);	// эскиз операции выдавливания
								entityExtr.Create();					// создать операцию

								kompas.ksMessage("Изменим параметры операции выдавливания");

								extrusionDef.SetSideParam(false,	// обратное направление
									(short)End_Type.etBlind,		// строго на глубину
									150, 0, false);
								extrusionDef.directionType = (short)Direction_Type.dtBoth; // направление выдавливания dtBoth - в оба направления
								entityExtr.Update();	// обновить параметры

								kompas.ksMessage("Поменяем эскиз");

								// режим редактирования эскиза
								sketchEdit = (ksDocument2D)sketchDef.BeginEdit();

								// создаим итератор и удалим все существующие объекты в эскизе
								ClearCurrentSketch(sketchEdit);
								// введем в эскиз окружность
								sketchEdit.ksCircle(0, 0, 100, 1);

								sketchDef.EndEdit();	// завершение редактирования эскиза
								entitySketch.Update();	// обновить параметры эскиза
								entityExtr.Update();	// обновить параметры операции выдавливания

								kompas.ksMessage("Приклеем выдавливанием");

								// создадим новый эскиз
								ksEntity entitySketch2 = (ksEntity)part.NewEntity((short)Obj3dType.o3d_sketch);
								if (entitySketch2 != null)
								{
									// интерфейс свойств эскиза
									ksSketchDefinition sketchDef2 = (ksSketchDefinition)entitySketch2.GetDefinition();
									if (sketchDef2 != null)
									{
										sketchDef2.SetPlane(basePlane);	// установим плоскость
										sketchDef2.angle = 45;			// повернем эскиз на 45 град.
										entitySketch2.Create();			// создадим эскиз

										// интерфейс редактора эскиза
										ksDocument2D sketchEdit2 = (ksDocument2D)sketchDef2.BeginEdit();
										sketchEdit2.ksCircle(0, 0, 150, 1);
										sketchDef2.EndEdit();	// завершение редактирования эскиза

										// приклеим выдавливанием
										ksEntity entityBossExtr = (ksEntity)part.NewEntity((short)Obj3dType.o3d_bossExtrusion);
										if (entityBossExtr != null)
										{
											ksBossExtrusionDefinition bossExtrDef = (ksBossExtrusionDefinition)entityBossExtr.GetDefinition();
											if (bossExtrDef != null)
											{
												ksExtrusionParam extrProp = (ksExtrusionParam)bossExtrDef.ExtrusionParam(); // интерфейс структуры параметров выдавливания
												ksThinParam thinProp = (ksThinParam)bossExtrDef.ThinParam();      // интерфейс структуры параметров тонкой стенки
												if (extrProp != null  && thinProp != null)
												{
													bossExtrDef.SetSketch(entitySketch2); // эскиз операции выдавливания

													extrProp.direction = (short)Direction_Type.dtNormal;      // направление выдавливания (прямое)
													extrProp.typeNormal = (short)End_Type.etBlind;      // тип выдавливания (строго на глубину)
													extrProp.depthNormal = 100;         // глубина выдавливания

													thinProp.thin = false;              // без тонкой стенки

													entityBossExtr.Create();                // создадим операцию
												}
											}
										}

										// создадим новый эскиз
										ksEntity entitySketch3 = (ksEntity)part.NewEntity((short)Obj3dType.o3d_sketch);
										if (entitySketch3 != null)
										{
											// интерфейс свойств эскиза
											ksSketchDefinition sketchDef3 = (ksSketchDefinition)entitySketch3.GetDefinition();
											if (sketchDef3 != null)
											{
												sketchDef3.SetPlane(basePlane);	// установим плоскость
												sketchDef3.angle = 45;			// повернем эскиз на 45 град.
												entitySketch3.Create();			// создадим эскиз

												// интерфейс редактора эскиза
												ksDocument2D sketchEdit3 = (ksDocument2D)sketchDef3.BeginEdit();
												// введем новый эскиз - квадрат
												sketchEdit3.ksLineSeg(50,  50, -50,  50, 1);
												sketchEdit3.ksLineSeg(50, -50, -50, -50, 1);
												sketchEdit3.ksLineSeg(50, -50,  50,  50, 1);
												sketchEdit3.ksLineSeg(-50, -50, -50,  50, 1);
												sketchDef3.EndEdit();	// завершение редактирования эскиза

												// вырежим выдавливанием
												ksEntity entityCutExtr = (ksEntity)part.NewEntity((short)Obj3dType.o3d_cutExtrusion);
												if (entityCutExtr != null)
												{
													ksCutExtrusionDefinition cutExtrDef = (ksCutExtrusionDefinition)entityCutExtr.GetDefinition();
													if (cutExtrDef != null)
													{
														cutExtrDef.SetSketch(entitySketch3);	// установим эскиз операции
														cutExtrDef.directionType = (short)Direction_Type.dtNormal; //прямое направление
														cutExtrDef.SetSideParam(true, (short)End_Type.etBlind, 50, 0, false);
														cutExtrDef.SetThinParam(false, 0, 0, 0);
													}

													entityCutExtr.Create();	// создадим операцию вырезание выдавливанием
												}
											}
										}
									}
								}
							}
						}
					}
				}
			}
		}


		// создание фаски и скругления
		void CreateFilletAndChamfer()
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
						sketchDef.angle = 45;			// угол поворота эскиза
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
								extrusionDef.directionType = (short)Direction_Type.dtNormal;	// направление выдавливания
								extrusionDef.SetSideParam(true /*прямое направление*/,
									(short)End_Type.etBlind /*строго на глубину*/, 100, 0, false);
								extrusionDef.SetThinParam(false, 0, 0, 0);	// без тонкой стенки
								extrusionDef.SetSketch(entitySketch);		// эскиз операции выдавливания
								entityExtr.Create();						// создать операцию

								ksEntityCollection collect = (ksEntityCollection)part.EntityCollection((short)Obj3dType.o3d_face);
								if (collect != null  && collect.SelectByPoint(0, 0, 0) && collect.GetCount()  != 0)
								{

									kompas.ksMessage("Создание скругления");

									ksEntity entityFillet = (ksEntity)part.NewEntity((short)Obj3dType.o3d_fillet);
									if (entityFillet != null)
									{
										ksFilletDefinition filletDef = (ksFilletDefinition)entityFillet.GetDefinition();
										if (filletDef != null)
										{
											filletDef.radius = 10;		// радиус скругления
											filletDef.tangent = false;	// продолжить по касательной
											ksEntityCollection arr = (ksEntityCollection)filletDef.array();	// динамический массив объектов
											if (arr != null)
											{
												for (int i = 0, count = collect.GetCount(); i < count; i++)
													arr.Add(collect.GetByIndex(i));

												entityFillet.Create();
											}
										}
									}
								}

								ksEntityCollection collect2 = (ksEntityCollection)part.EntityCollection((short)Obj3dType.o3d_face);
								if (collect2 != null  && collect2.SelectByPoint(0, 0, 100) && collect2.GetCount() != 0) 
								{

									kompas.ksMessage("Создание фаски");
									ksEntity entityChamfer = (ksEntity)part.NewEntity((short)Obj3dType.o3d_chamfer);
									if (entityChamfer != null)
									{
										ksChamferDefinition ChamferDef = (ksChamferDefinition)entityChamfer.GetDefinition();
										if (ChamferDef != null)
										{
											ChamferDef.tangent = false;
											ChamferDef.SetChamferParam(true, 10, 10);
											ksEntityCollection arr = (ksEntityCollection)ChamferDef.array();	// динамический массив объектов
											if (arr != null)
											{
												for (int i = 0, count = collect2.GetCount(); i < count; i++)
													arr.Add(collect2.GetByIndex(i));

												entityChamfer.Create();
											}
										}
									}
								}
							}
						}
					}
				}
			}
		}


		// Операции вращения
		void OperationRotated()
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
						if (basePlane != null)
						{
							sketchDef.SetPlane(basePlane);	// установим плоскость XOY базовой для эскиза
							entitySketch.Create();			// создадим эскиз

							// интерфейс редактора эскиза
							ksDocument2D sketchEdit = (ksDocument2D)sketchDef.BeginEdit();
							if (sketchEdit != null)
							{
								sketchEdit.ksArcByAngle(0, 0, 20, -90, 90, 1, 1);
								sketchEdit.ksLineSeg(0, -20, 0, 20, 3);
								sketchDef.EndEdit();	// завершение редактирования эскиза
							}

							ksEntity entityRotate = (ksEntity)part.NewEntity((short)Obj3dType.o3d_bossRotated);
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
									entityRotate.Create();				// создать операцию
								}
							}
						}
					}
				}
				kompas.ksMessage("Базовая операция вращения");

				ksEntity entitySketch2 = (ksEntity)part.NewEntity((short)Obj3dType.o3d_sketch);
				if (entitySketch2 != null)
				{
					ksSketchDefinition sketchDef2 = (ksSketchDefinition)entitySketch2.GetDefinition();
					if (sketchDef2 != null)
					{
						ksEntity basePlane = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);
						if (basePlane != null)
						{
							sketchDef2.SetPlane(basePlane);	// установим плоскость XOY базовой для эскиза
							entitySketch2.Create();			// создадим эскиз
							// интерфейс редактора эскиза
							ksDocument2D sketchEdit2 = (ksDocument2D)sketchDef2.BeginEdit();
							if (sketchEdit2 != null)
							{
								sketchEdit2.ksArcByAngle(15, 0, 10, -90, 90, 1, 1);
								sketchEdit2.ksLineSeg(15, -10, 15, 10, 3);
								sketchDef2.EndEdit();	// завершение редактирования эскиза
							}

							ksEntity entityBossRotate = (ksEntity)part.NewEntity((short)Obj3dType.o3d_bossRotated);
							if (entityBossRotate != null)
							{
								ksBossRotatedDefinition bossRotateDef = (ksBossRotatedDefinition)entityBossRotate.GetDefinition();
								if (bossRotateDef != null)
								{
									bossRotateDef.directionType = (short)Direction_Type.dtNormal;
									bossRotateDef.SetSideParam(true, 360);
									bossRotateDef.SetSketch(entitySketch2);		// эскиз операции вращения
									entityBossRotate.Create();					// создать операцию
								}
							}
						}
					}
				}
				kompas.ksMessage("Операция приклеивания вращением");

				ksEntity entitySketch3 = (ksEntity)part.NewEntity((short)Obj3dType.o3d_sketch);
				if (entitySketch3 != null)
				{
					ksSketchDefinition sketchDef = (ksSketchDefinition)entitySketch3.GetDefinition();
					if (sketchDef != null)
					{
						ksEntity basePlane = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);
						if (basePlane != null)
						{
							sketchDef.SetPlane(basePlane);	// установим плоскость XOY базовой для эскиза
							entitySketch3.Create();			// создадим эскиз
							// интерфейс редактора эскиза
							ksDocument2D sketchEdit3 = (ksDocument2D)sketchDef.BeginEdit();
							if (sketchEdit3 != null)
							{
								sketchEdit3.ksArcByAngle(20, 0, 20, 90, 270, 1, 1);
								sketchEdit3.ksLineSeg(20, -20, 20, 20, 3);
								sketchDef.EndEdit();	// завершение редактирования эскиза
							}

							ksEntity entityCutRotate = (ksEntity)part.NewEntity((short)Obj3dType.o3d_cutRotated);
							if (entityCutRotate != null)
							{
								ksCutRotatedDefinition cutRotateDef = (ksCutRotatedDefinition)entityCutRotate.GetDefinition();
								if (cutRotateDef != null)
								{
									cutRotateDef.directionType = (short)Direction_Type.dtNormal;
									cutRotateDef.SetSideParam(true, 90);
									cutRotateDef.SetThinParam(true, (short)Direction_Type.dtBoth, 5, 7);	// тонкая стенка в два направления
									cutRotateDef.SetSketch(entitySketch3);	// эскиз операции вращения
									entityCutRotate.Create();				// создать операцию
								}
							}
						}
					}
				}
				kompas.ksMessage("Операция вырезания вращением");
			}
		}


		// операции по сечениям
		void OperationLoft()
		{
      ksPart part = (ksPart)doc.GetPart((short)Part_Type.pTop_Part);	// новый компонент
			if (part != null)
			{
				// создадим эскиз
				ksEntity entitySketch = (ksEntity)part.NewEntity((short)Obj3dType.o3d_sketch);
				if (entitySketch != null)
				{
					ksSketchDefinition sketchDef = (ksSketchDefinition)entitySketch.GetDefinition();
					if (sketchDef != null)
					{
						ksEntity basePlane = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);	// получим интерфейс базовой плоскости XOY
						sketchDef.SetPlane(basePlane);	// установим плоскость XOY базовой для эскиза
						entitySketch.Create();			// создадим эскиз
						entitySketch.hidden = false;

						// интерфейс редактора эскиза
						ksDocument2D sketchEdit = (ksDocument2D)sketchDef.BeginEdit();
						if (sketchEdit != null)
							sketchEdit.ksCircle(0, 0, 4.5, 1);
						sketchDef.EndEdit();	// завершение редактирования эскиза
					}
				}

				// создадим смещенную плоскость, а в ней эскиз
				ksEntity entityOffsetPlane = (ksEntity)part.NewEntity((short)Obj3dType.o3d_planeOffset);
				ksEntity entitySketch2 = (ksEntity)part.NewEntity((short)Obj3dType.o3d_sketch);
				if (entityOffsetPlane != null)
				{
					// интерфейс свойств смещенной плоскости
					ksPlaneOffsetDefinition offsetDef = (ksPlaneOffsetDefinition)entityOffsetPlane.GetDefinition();
					if (offsetDef != null)
					{
						offsetDef.offset = 30;							// расстояние от базовой плоскости
						ksEntity basePlane = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);
						basePlane.name = "Смещенная плоскость";			// название для плоскости

						offsetDef.SetPlane(basePlane);					// базовая плоскость
						entityOffsetPlane.name = "Смещенная плоскость";	// имя для смещенной плоскости
						entityOffsetPlane.hidden = true;
						entityOffsetPlane.Create();						// создать смещенную плоскость

						if (entitySketch2 != null)
						{
							ksSketchDefinition sketchDef = (ksSketchDefinition)entitySketch2.GetDefinition();
							if (sketchDef != null)
							{
								sketchDef.SetPlane(entityOffsetPlane);	// установим плоскость XOY базовой для эскиза
								entitySketch2.Create();					// создадим эскиз

								// интерфейс редактора эскиза
								ksDocument2D sketchEdit = (ksDocument2D)sketchDef.BeginEdit();
								sketchEdit.ksCircle(0, 0, 8, 1);
								sketchDef.EndEdit();					// завершение редактирования эскиза
							}
						}
					}
				}

				// создадим смещенную плоскость, а в ней эскиз
				ksEntity entityOffsetPlane2 = (ksEntity)part.NewEntity((short)Obj3dType.o3d_planeOffset);
				ksEntity entitySketch3 = (ksEntity)part.NewEntity((short)Obj3dType.o3d_sketch);
				if (entityOffsetPlane2 != null  && entitySketch3 != null)
				{
					// интерфейс свойств смещенной плоскости
					ksPlaneOffsetDefinition offsetDef = (ksPlaneOffsetDefinition)entityOffsetPlane2.GetDefinition();
					if (offsetDef != null)
					{
						offsetDef.offset = 60;	// расстояние от базовой плоскости
						ksEntity basePlane = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);
						basePlane.name = "Смещенная плоскость";	// название для плоскости
        
						offsetDef.SetPlane(basePlane);						// базовая плоскость
						entityOffsetPlane2.name = "Смещенная плоскость2";	// имя для смещенной плоскости
						entityOffsetPlane2.hidden = true;
						entityOffsetPlane2.Create();						// создать смещенную плоскость 

						ksSketchDefinition sketchDef = (ksSketchDefinition)entitySketch3.GetDefinition();
						if (sketchDef != null)
						{
							sketchDef.SetPlane(entityOffsetPlane2);			// установим плоскость XOY базовой для эскиза
							entitySketch3.Create();							// создадим эскиз

							// интерфейс редактора эскиза
							ksDocument2D sketchEdit = (ksDocument2D)sketchDef.BeginEdit(); 
							sketchEdit.ksCircle(0, 0, 1.5, 1);
							sketchDef.EndEdit();							// завершение редактирования эскиза                
						}
					}
				}

				// создадим базовую операцию по сечениям
				ksEntity entityBaseLoft = (ksEntity)part.NewEntity((short)Obj3dType.o3d_bossLoft); 
				if (entityBaseLoft != null)
				{
					ksBossLoftDefinition baseLoft = (ksBossLoftDefinition)entityBaseLoft.GetDefinition();
					if (baseLoft != null)
					{
						ksEntityCollection entCol = (ksEntityCollection)baseLoft.Sketchs();
						if (entCol != null)
						{
							entCol.Add(entitySketch);
							entCol.Add(entitySketch2);
							entCol.Add(entitySketch3);
						}
						entityBaseLoft.name = "Ручка";
						entityBaseLoft.SetAdvancedColor(12345678, .8, .8, .8, .8, 1, .8);
						entityBaseLoft.Create();	// создать операцию
					}
				}
				kompas.ksMessage("Базовая операция по сечениям");

				// создадим смещенную плоскость, а в ней эскиз
				ksEntity entitySketch7 = (ksEntity)part.NewEntity((short)Obj3dType.o3d_sketch);
				if (entitySketch7 != null)
				{
					// интерфейс свойств смещенной плоскости
					ksSketchDefinition sketchDef = (ksSketchDefinition)entitySketch7.GetDefinition();
					if (sketchDef != null)
					{
						sketchDef.SetPlane(entityOffsetPlane2);	// установим плоскость XOY базовой для эскиза
						entitySketch7.Create();					// создадим эскиз

						// интерфейс редактора эскиза
						ksDocument2D sketchEdit = (ksDocument2D)sketchDef.BeginEdit();
						sketchEdit.ksCircle(0, 0, 1.5, 1);
						sketchDef.EndEdit();					// завершение редактирования эскиза                
					}
				}

				// создадим смещенную плоскость, а в ней эскиз
				ksEntity entityOffsetPlane3 = (ksEntity)part.NewEntity((short)Obj3dType.o3d_planeOffset);
				ksEntity entitySketch4 = (ksEntity)part.NewEntity((short)Obj3dType.o3d_sketch);
				if (entityOffsetPlane3 != null  && entitySketch4 != null)
				{
					// интерфейс свойств смещенной плоскости
					ksPlaneOffsetDefinition offsetDef = (ksPlaneOffsetDefinition)entityOffsetPlane3.GetDefinition();
					if (offsetDef != null)
					{
						offsetDef.offset = 120;					// расстояние от базовой плоскости
						ksEntity basePlane = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);
						basePlane.name = "Смещенная плоскость";	// название для плоскости
        
						offsetDef.SetPlane(basePlane);			// базовая плоскость
						entityOffsetPlane3.name = "Смещенная плоскость";	// имя для смещенной плоскости
						entityOffsetPlane3.hidden = true;
						entityOffsetPlane3.Create();			// создать смещенную плоскость 

						ksSketchDefinition sketchDef = (ksSketchDefinition)entitySketch4.GetDefinition();
						if (sketchDef != null)
						{
							sketchDef.SetPlane(entityOffsetPlane3);	// установим плоскость XOY базовой для эскиза
							entitySketch4.Create();					// создадим эскиз

							// интерфейс редактора эскиза
							ksDocument2D sketchEdit = (ksDocument2D)sketchDef.BeginEdit();
							sketchEdit.ksCircle(0, 0, 1.8, 1);
							sketchDef.EndEdit();					// завершение редактирования эскиза                
						}
					}
				}


				// создадим операцию приклеивания по сечениям
				ksEntity entityBossLoft = (ksEntity)part.NewEntity((short)Obj3dType.o3d_bossLoft);
				if (entityBossLoft != null)
				{
					ksBossLoftDefinition bossLoft = (ksBossLoftDefinition)entityBossLoft.GetDefinition();
					if (bossLoft != null)
					{
						ksEntityCollection entCol = (ksEntityCollection)bossLoft.Sketchs();
						if (entCol != null)
						{
							entCol.Add(entitySketch7);
							entCol.Add(entitySketch4);
						}
						entityBossLoft.name = "Цевьё";
						entityBossLoft.SetAdvancedColor(1234567890, .8, .8, .8, .8, 1, .8);

						entityBossLoft.Create();	// создать операцию
					}
				}
				kompas.ksMessage("Операция приклеивание по сечениям");

				// создадим эскиз в уже созданной смещенной плоскости
				ksEntity entitySketch5 = (ksEntity)part.NewEntity((short)Obj3dType.o3d_sketch);
				if (entitySketch5 != null)
				{
					ksSketchDefinition sketchDef = (ksSketchDefinition)entitySketch5.GetDefinition();
					if (sketchDef != null)
					{
						sketchDef.SetPlane(entityOffsetPlane3);	// установим плоскость XOY базовой для эскиза
						entitySketch5.Create();					// создадим эскиз

						// интерфейс редактора эскиза
						ksDocument2D sketchEdit = (ksDocument2D)sketchDef.BeginEdit();
						ksRectangleParam recPar = (ksRectangleParam)kompas.GetParamStruct((short)StructType2DEnum.ko_RectangleParam);
						recPar.Init();
						if (recPar != null)
						{
							recPar.x = -1.8;
							recPar.y = -.4;
							recPar.height = .8;
							recPar.width = 3.6;
							recPar.style = 1;
						}
						sketchEdit.ksRectangle(recPar, 0); 
						sketchDef.EndEdit();	// завершение редактирования эскиза                
					}
				}

				// создадим смещенную плоскость, а в ней эскиз
				ksEntity entityOffsetPlane4 = (ksEntity)part.NewEntity((short)Obj3dType.o3d_planeOffset);
				ksEntity entitySketch6 = (ksEntity)part.NewEntity((short)Obj3dType.o3d_sketch);
				if (entityOffsetPlane4 != null  && entitySketch6 != null)
				{
					// интерфейс свойств смещенной плоскости
					ksPlaneOffsetDefinition offsetDef = (ksPlaneOffsetDefinition)entityOffsetPlane4.GetDefinition();
					if (offsetDef != null)
					{
						offsetDef.offset = 110;	// расстояние от базовой плоскости
						ksEntity basePlane = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);
						basePlane.name = "Смещенная плоскость";	// название для плоскости
        
						offsetDef.SetPlane(basePlane);	// базовая плоскость
						entityOffsetPlane4.name = "Смещенная плоскость";	// имя для смещенной плоскости
						entityOffsetPlane4.hidden = true;
						entityOffsetPlane4.Create();	// создать смещенную плоскость 

						ksSketchDefinition sketchDef = (ksSketchDefinition)entitySketch6.GetDefinition();
						if (sketchDef != null)
						{
							sketchDef.SetPlane(entityOffsetPlane4);	// установим плоскость XOY базовой для эскиза
							entitySketch6.Create();					// создадим эскиз

							// интерфейс редактора эскиза
							ksDocument2D sketchEdit = (ksDocument2D)sketchDef.BeginEdit(); 
							ksRectangleParam recPar = (ksRectangleParam)kompas.GetParamStruct((short)StructType2DEnum.ko_RectangleParam);
							recPar.Init();
							if (recPar != null)
							{
								recPar.x = -1.8;
								recPar.y = -1.8;
								recPar.height = 3.6;
								recPar.width = 3.6;
								recPar.style = 1;
							}
							sketchEdit.ksRectangle(recPar, 0); 
							sketchDef.EndEdit();	// завершение редактирования эскиза                
						}
					}
				}


				// создадим операцию вырезания по сечениям
				ksEntity entityCutLoft = (ksEntity)part.NewEntity((short)Obj3dType.o3d_cutLoft); 
				if (entityCutLoft != null)
				{
					ksCutLoftDefinition cutLoft = (ksCutLoftDefinition)entityCutLoft.GetDefinition();
					if (cutLoft != null)
					{
						ksEntityCollection entCol = (ksEntityCollection)cutLoft.Sketchs();
						if (entCol != null)
						{
							entCol.Add(entitySketch5);
							entCol.Add(entitySketch6);
						}

						cutLoft.SetThinParam(true, (short)Direction_Type.dtNormal, 3, 0);
						entityCutLoft.name = "Рабочая поверхность";
						entityCutLoft.SetAdvancedColor(1234, .8, .8, .8, .8, 1, .8);

						entityCutLoft.Create();	// создать операцию
					}
				}

				kompas.ksMessage("Операция вырезания по сечениям");		
			}
		}


		// Операции : оболочка, уклон, сечение плоскостью, сечение эскизом
		void CreateNextOper()
		{
			// создадим новую деталь
      ksPart part = (ksPart)doc.GetPart((short)Part_Type.pTop_Part);
			if (part != null)
			{
				// создадим эскиз для базовой операции
				ksEntity entitySketch = (ksEntity)part.NewEntity((short)Obj3dType.o3d_sketch);
				if (entitySketch != null)
				{
					// получим интерфейс свойств эскиза
					ksSketchDefinition sketchDef = (ksSketchDefinition)entitySketch.GetDefinition();
					if (sketchDef != null)
					{
						// получим интерфейс базовой плоскости XOY
						ksEntity basePlane = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);
						sketchDef.SetPlane(basePlane);	// установим плоскость XOY базовой для эскиза
						sketchDef.angle = 0;			// угол поворота эскиза
						entitySketch.Create();			// создадим эскиз

						// интерфейс редактора эскиза
						ksDocument2D sketchEdit = (ksDocument2D)sketchDef.BeginEdit(); 
						// введем новый эскиз - квадрат
						sketchEdit.ksLineSeg(50,  50, -50,  50, 1);
						sketchEdit.ksLineSeg(50, -50, -50, -50, 1); 
						sketchEdit.ksLineSeg(50, -50,  50,  50, 1);
						sketchEdit.ksLineSeg(-50, -50, -50,  50, 1);
						sketchDef.EndEdit();	// завершение редактирования эскиза                

						kompas.ksMessage("Создаем операцию выдавливания");
						// Создаем операцию выдавливания
						ksEntity entityExtr = (ksEntity)part.NewEntity((short)Obj3dType.o3d_bossExtrusion);
						if (entityExtr != null)
						{
							// интерфейс свойств базовой операции выдавливания
							ksBossExtrusionDefinition extrusionDef = (ksBossExtrusionDefinition)entityExtr.GetDefinition(); 
							if (extrusionDef != null)
							{
								extrusionDef.directionType = (short)Direction_Type.dtNormal;	// направление выдавливания
								extrusionDef.SetSideParam(true /*прямое направление*/,
									(short)End_Type.etBlind /*строго на глубину*/, 200, 0, false); 
								extrusionDef.SetThinParam(true, (short)Direction_Type.dtBoth, 10, 10);	// тонкая стенка в два направления
								extrusionDef.SetSketch(entitySketch);	// эскиз операции выдавливания
								entityExtr.Create();					// создать операцию
							}

							bool update = false;	// если update = true, то параметры операции изменены 
							if (MessageBox.Show("Создать операцию \"Оболочка\"?", "Сообщение библиотеки", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
							{
								// Создаем операцию оболочка
								ksEntity entShell = (ksEntity)part.NewEntity((short)Obj3dType.o3d_shellOperation);
								if (entShell != null)
								{
									// интерфейс свойств базовой операции выдавливания
									ksShellDefinition incDef = (ksShellDefinition)entShell.GetDefinition();
									if (incDef != null)
									{
										ksEntityCollection entCol = (ksEntityCollection)incDef.FaceArray();               // заведем массив граней для операции оболочка
										ksEntityCollection collect = (ksEntityCollection)part.EntityCollection((short)Obj3dType.o3d_face); // массив всех граней
										if (entCol != null  && collect != null)
										{
											incDef.thickness = 8;				// толщиина оболочки
											incDef.thinType = true;				// направление оболочки внутрь
											collect.SelectByPoint(50, 0, 0);	// выделим точку на детале для заполнения массива всех граней, рпоходяших через эту точку
											entCol.Add(collect.GetByIndex(0));	// добавим в массив граней для операции грань с индексом = 0 
											collect.refresh();					// очистим массив
											entShell.Create();					// создать операцию

											if (MessageBox.Show("Изменить параметры операции \"Оболочка\"?", "Сообщение библиотеки", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
											{

												incDef.thickness = 25;				// толщиина оболочки
												incDef.thinType = false;			// направление оболочки наружу
												collect.SelectByPoint(60, 0, 10);	// выделим точку на детале для заполнения массива всех граней, рпоходяших через эту точку
												entCol.Add(collect.GetByIndex(0));	// добавим в массив граней для операции грань с индексом = 0 
												collect.refresh();					// очистим массив
												entShell.Update();					// перестроим операцию
												update = true;
											}
										}
									}
								}
							}

							if (MessageBox.Show("Создать операцию \"Уклон\"?", "Сообщение библиотеки", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
							{
								// создадим операцию уклон
								ksEntity entInc = (ksEntity)part.NewEntity((short)Obj3dType.o3d_incline);
								if (entInc != null)
								{
									// интерфейс свойств базовой операции уклона
									ksInclineDefinition incDef = (ksInclineDefinition)entInc.GetDefinition();
									ksEntityCollection collect = (ksEntityCollection)part.EntityCollection((short)Obj3dType.o3d_face); // массив всех граней
									if (incDef != null  && collect != null) 
									{
										incDef.direction = true;		// Направление уклона - наружу  
										incDef.angle = 3;				// Угол уклона  
										incDef.SetPlane(basePlane);		// базовая плоскость  
										ksEntityCollection entColInc = (ksEntityCollection)incDef.FaceArray();  // заведем массив граней для операции уклон
										if (entColInc != null)
										{
											collect.SelectByPoint(0, update ? 85 : 60, 10);	// выделим точку на детале для заполнения массива всех граней, проходяших через эту точку
											entColInc.Add(collect.GetByIndex(0));			// добавим в массив граней для операции грань с индексом = 0 
											collect.refresh();								// очистим массив
											entInc.Create();								// создать операцию

											if (MessageBox.Show("Изменить параметры операции \"Уклон\"?", "Сообщение библиотеки", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
											{
												incDef.direction = false;							// Направление уклона - внутрь
												incDef.angle = 25;									// Угол уклона  
												collect.SelectByPoint(0, update ? -85 : -60, 10);	// выделим точку на детале для заполнения массива всех граней, проходяших через эту точку
												entColInc.Add(collect.GetByIndex(0));				// добавим в массив граней для операции грань с индексом = 0 
												collect.refresh();									// очистим массив
												entInc.Update();
											}
										}
									}
								}
							}

							if (MessageBox.Show("Создать операцию \"Сечение плоскостью\"?", "Сообщение библиотеки", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
							{
								// создадим операцию сечение плоскостью
								ksEntity entityOpr = (ksEntity)part.NewEntity((short)Obj3dType.o3d_cutByPlane);
								if (entityOpr != null)
								{
									// интерфейс свойств операции сечение плоскостью
									ksCutByPlaneDefinition incOpr = (ksCutByPlaneDefinition)entityOpr.GetDefinition(); // интерфейс операции
									if (incOpr != null)
									{
										// получим интерфейс базовой плоскости XOZ, которая будет являться плоскостью сечения
										ksEntity basePlaneXOZ = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOZ);
										incOpr.SetPlane(basePlaneXOZ);	// задаем плоскость сечения
										incOpr.direction = false;		// направление сечения - обратное
										entityOpr.Create();				// создать операцию

										if (MessageBox.Show("Изменить параметры операции \"Сечение плоскостью\"?", "Сообщение библиотеки", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
										{
											incOpr.direction = true;	// направление сечения - прямое
											entityOpr.Update();			// перестроить операцию
										}
									}
								}
							}

							if (MessageBox.Show("Создать операцию \"Сечение эскизом\"?", "Сообщение библиотеки", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
							{
								// создадим эскиз для операции сечение эскизом
								ksEntity entitySketch2 = (ksEntity)part.NewEntity((short)Obj3dType.o3d_sketch);  
								if (entitySketch2 != null)
								{
									// интерфейс свойств эскиза
									ksSketchDefinition sketchDef2 = (ksSketchDefinition)entitySketch2.GetDefinition();
									if (sketchDef2 != null)
									{
										// получим интерфейс базовой плоскости YOZ
										ksEntity basePlaneYOZ = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeYOZ);
										sketchDef2.SetPlane(basePlaneYOZ);	// установим плоскость yoz базовой для эскиза
										sketchDef2.angle = 45;				// угол поворота эскиза
										entitySketch2.Create();				// создадим эскиз

										// интерфейс редактора эскиза
										ksDocument2D sketchEdit2 = (ksDocument2D)sketchDef2.BeginEdit(); 
										// введем новый эскиз
										sketchEdit2.ksArcBy3Points(-200, 45, -150, 10, -50, 10, 1);
										sketchEdit2.ksLineSeg     (-200, 45, -300, 20, 1);
										sketchEdit2.ksLineSeg     (-50, 10,   60, 10, 1);
										sketchDef2.EndEdit();	// завершение редактирования эскиза
              
										// создадаим Операцию сечение эскизом
										ksEntity entityOpr = (ksEntity)part.NewEntity((short)Obj3dType.o3d_cutBySketch);
										if (entityOpr != null)
										{
											// интерфейс операции сечение эскизом
											ksCutBySketchDefinition incOpr = (ksCutBySketchDefinition)entityOpr.GetDefinition(); 
											if (incOpr != null)
											{
												incOpr.SetSketch(entitySketch2);	// задаем эскиз
												incOpr.direction = true;			// задаем направление сечения
												entityOpr.Create();					// создать операцию
											}
										}
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
