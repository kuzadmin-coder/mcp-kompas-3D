
using Kompas6API5;

using System;
using Microsoft.Win32;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Resources;
using Kompas6Constants;
using Kompas6Constants3D;
using KAPITypes;

using reference = System.Int32;

namespace Steps.NET
{
	// Класс EventsAuto - Тест событий на C#
  [ClassInterface(ClassInterfaceType.AutoDual)]
	public class EventsAuto
	{
		#region Private fields
		KompasObject kompas;
		#endregion

		#region Constants
		private const int SPC_BASE_OBJECT = 1;	// базовый объект
		private const int SPC_COMMENT = 2;		// комментарий
		private const int IDM_REQUEST_OBJECT_2D = 10001;
		#endregion

		// Имя библиотеки
		[return: MarshalAs(UnmanagedType.BStr)] public string GetLibraryName()
		{
			ResourceManager rm = new ResourceManager(typeof (EventsAuto));
			return rm.GetString("LibraryName");
		}
		

		// Головная функция библиотеки
		public void ExternalRunCommand([In] short command, [In] short mode, [In, MarshalAs(UnmanagedType.IDispatch)] object kompas_)
		{
			kompas = (KompasObject) kompas_;
			Global.Kompas = kompas;

			switch ((int)command)
			{
					#region События приложения
				case 1 : // Подписаться
					if (!BaseEvent.FindEvent(typeof(ApplicationEvent), null, -1, null)) 
					{
						ApplicationEvent aplEvent = new ApplicationEvent(kompas, true);	// Обработчик событий приложения КОМПАС
						aplEvent.Advise();
					}
					else 
						kompas.ksError("На события приложения Компас уже подписались");
					break;

				case 2 : // Отписаться
					BaseEvent.TerminateEvents(typeof(ApplicationEvent), null, -1, null);
					break;
					#endregion

					#region События документов
				case 3 : // Подписаться
					object doc = kompas.ksGetDocumentByReference(0);
					int docType = kompas.ksGetDocumentType(0);
					if (doc != null) 
					{
						if (!BaseEvent.FindEvent(typeof(DocumentEvent), doc, -1, null)) 
							AdviseDoc((ksDocumentFileNotify_Event)doc, docType, false, false,
								false, false, false, false, -1);
					}
					else
						kompas.ksError("Нет активного документа"); 
					break;

				case 4 : // Отписаться от всех событий документа
					if (kompas != null) 
					{
						doc = kompas.ksGetDocumentByReference(0);
						if (doc != null)
							BaseEvent.TerminateEvents(typeof(DocumentEvent), doc, -1, null);
						else
							kompas.ksError("Нет активного документа");
					} 
					break;
					#endregion

					#region События объекта 2D документа
				case 5 : // Подписаться
					ksDocument2D doc2D = (ksDocument2D)kompas.ActiveDocument2D();
					docType = kompas.ksGetDocumentType(0);
					if (doc2D != null)
					{
						int objType = (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrDrawText;
//						if (Request2DObject(ref objType))
							AdviseDoc((ksDocumentFileNotify_Event)doc2D, docType, false, 
								true, 
								false, 
								false, 
								false, 
								false, objType);
					}
					else
						kompas.ksError("Нет активного 2D документа");
					break;

				case 6  : // Отписаться
					doc2D = (ksDocument2D)kompas.ActiveDocument2D();
					docType = kompas.ksGetDocumentType(0);
					if (doc2D != null) 
					{
						int objType = 0;
						if (Request2DObject(ref objType))
							BaseEvent.TerminateEvents(typeof(Object2DEvent), doc2D, objType, null); 
					}
					else
						kompas.ksError("Нет активного 2D документа");
					break;

				case 7  : // Отписаться от всех событий объектов 2D
					doc2D = (ksDocument2D)kompas.ActiveDocument2D();
					if (doc2D != null) 
						BaseEvent.TerminateEvents(typeof(Object2DEvent), doc2D, -1, null); 
					else
						kompas.ksError("Нет активного 2D документа");
					break;

				case 8  : // Подписаться на события вида по номеру
					doc2D = (ksDocument2D)kompas.ActiveDocument2D();
					if (doc2D != null) 
					{
						int view = 0;
						if (kompas.ksReadInt("Введите номер вида", 0, 0, 254, ref view) != 0) 
						{
							reference refView = doc2D.ksGetViewReference(view);
							if (refView != 0)       
							{
								if (!BaseEvent.FindEvent(typeof(Object2DEvent), doc2D, refView, null)) 
								{          
									object objNotify = doc2D != null ? doc2D.GetObject2DNotify(refView) : null;
									if (objNotify != null)
									{  
										Object2DEvent objEvent = new Object2DEvent(objNotify, doc2D, refView, doc2D.GetObject2DNotifyResult(), true); 
										objEvent.Advise();
									}
								}
								else
									kompas.ksError("На события объекта 2D документа уже подписались");
							}
							else
								kompas.ksError("Вид с таким номером не существует");
						}
					}
					else
						kompas.ksError("Нет активного 2D документа");
					break;

				case 9 : // Отписаться от событий для вида по номеру
					doc2D = (ksDocument2D)kompas.ActiveDocument2D();
					if (doc2D != null) 
					{
						int view = 0;
						if (kompas.ksReadInt("Введите номер вида", 0, 0, 254, ref view) != 0) 
						{
							reference refView = doc2D.ksGetViewReference(view);
							if (refView != 0) 
								BaseEvent.TerminateEvents(typeof(Object2DEvent), doc2D, refView, null);
							else
								kompas.ksError("Вид с таким номером не существует");
						}
					}
					else
						kompas.ksError("Нет активного 2D документа");
					break;

				case 10 : // Подписаться на события слоя по номеру
					doc2D  = (ksDocument2D)kompas.ActiveDocument2D();
					if (doc2D != null)
					{
						int layer = 0;
						if (kompas.ksReadInt("Введите номер слоя", 0, 0, 254, ref layer) != 0)
						{
							reference refLayer = doc2D.ksGetLayerReference(layer);
							if (refLayer != 0) 
							{
								if (!BaseEvent.FindEvent(typeof(Object2DEvent), doc2D, refLayer, null)) 
								{          
									object objNotify = doc2D != null ? doc2D.GetObject2DNotify(refLayer) : null;
									if (objNotify != null) 
									{  
										Object2DEvent objEvent = new Object2DEvent(objNotify, doc2D, refLayer, doc2D.GetObject2DNotifyResult(), true); 
										objEvent.Advise();
									}
								}
								else
									kompas.ksError("На события объекта 2D документа уже подписались");
							}
							else
								kompas.ksError("Слой с таким номером не существует");
						}
					}
					else
						kompas.ksError("Нет активного 2D документа");
					break;

				case 11 : // Отписаться от событий для слоя по номеру
					doc2D = (ksDocument2D)kompas.ActiveDocument2D();
					if (doc2D != null) 
					{
						int layer = 0;
						if (kompas.ksReadInt("Введите номер слоя", 0, 0, 254, ref layer) != 0) 
						{
							reference refLayer = doc2D.ksGetLayerReference(layer);
							if (refLayer != 0) 
								BaseEvent.TerminateEvents(typeof(Object2DEvent), doc2D, refLayer, null);
							else
								kompas.ksError("Слой с таким номером не существует");
						}
					}
					else
						kompas.ksError("Нет активного 2D документа");
					break;
					#endregion

					#region События селектирования  
				case 12 : // Подписаться 
					doc = kompas.ksGetDocumentByReference(0);
					docType = kompas.ksGetDocumentType(0);
					if (doc != null &&  
						(docType == (int)DocType.lt_DocPart3D || 
						docType == (int)DocType.lt_DocAssemble3D ||
						docType == (int)DocType.lt_DocSheetStandart ||
						docType == (int)DocType.lt_DocSheetUser ||  
						docType == (int)DocType.lt_DocFragment))
						AdviseDoc(doc, docType, true, 
							false, 
							false, 
							false, 
							false, 
							false,
							-1);
					else
						kompas.ksError("Нет активного 2D или 3D документа");
					break;

				case 13 : // Отписаться
					doc = kompas.ksGetDocumentByReference(0);
					docType = kompas.ksGetDocumentType(0);
					if (doc != null &&  
						(docType == (int)DocType.lt_DocPart3D || 
						docType == (int)DocType.lt_DocAssemble3D ||
						docType == (int)DocType.lt_DocSheetStandart ||
						docType == (int)DocType.lt_DocSheetUser ||  
						docType == (int)DocType.lt_DocFragment))
						BaseEvent.TerminateEvents(typeof(SelectMngEvent), doc, -1, null); 
					else
						kompas.ksError("Нет активного 2D или 3D документа");
					break;
					#endregion

					#region События редактирования штампа
				case 14 : // Подписаться 
					doc = kompas.ksGetDocumentByReference(0);
					docType = kompas.ksGetDocumentType(0);
					if (doc != null && 
						docType != (int)DocType.lt_DocPart3D && 
						docType != (int)DocType.lt_DocAssemble3D) 
						AdviseDoc((ksDocumentFileNotify_Event)doc, docType, false, 
							false, 
							true, 
							false, 
							false, 
							false,
							-1);
					else
						kompas.ksError("Нет активного 2D, текстового документа или документа спецификации");
					break;

				case 15 : // Отписаться
					doc = kompas.ksGetDocumentByReference(0);
					docType = kompas.ksGetDocumentType(0);
					if (doc != null && 
						docType != (int)DocType.lt_DocPart3D && 
						docType != (int)DocType.lt_DocAssemble3D) 
						BaseEvent.TerminateEvents(typeof(StampEvent), doc, -1, null); 
					else
						kompas.ksError("Нет активного 2D, текстового документа или документа спецификации");
					break;
					#endregion
			
					#region События объектов 3D документа
      				// Подписаться
				case 16: // Указать объект в модели 
					ksDocument3D doc3D = (ksDocument3D)kompas.ActiveDocument3D();
					docType = kompas.ksGetDocumentType(0);
					if (doc3D != null) 
						AdviseDoc((ksDocumentFileNotify_Event)doc3D, docType, false, 
							true, 
							false, 
							false, 
							false, 
							false,
							-1);
					else
						kompas.ksError("Нет активного 3D документа");
					break;

					// Отписаться
				case 17: // Указать объект в модели 
					doc3D = (ksDocument3D)kompas.ActiveDocument3D();
					docType = kompas.ksGetDocumentType(0);
					if (doc3D != null) 
					{
						ksFeature obj3D = null;
						int objType = 0;
						if (Request3DObject(ref obj3D, ref objType))
							BaseEvent.TerminateEvents(typeof(Object3DEvent), doc3D, objType, obj3D); 
					}
					else
						kompas.ksError("Нет активного 3D документа");
					break;

				case 18: // Отписаться от всех объектов 3D документа
					doc3D = (ksDocument3D)kompas.ActiveDocument3D();
					if (doc3D != null) 
						BaseEvent.TerminateEvents(typeof(Object3DEvent), doc3D, -1, null); 
					else
						kompas.ksError("Нет активного 3D документа");
					break;
					#endregion

					#region Cобытия спецификации 
				case 19 : // Подписаться 
					doc = kompas.ksGetDocumentByReference(0);
					docType = kompas.ksGetDocumentType(0);
					if (doc != null && 
						docType != (int)DocType.lt_DocTxtStandart && 
						docType != (int)DocType.lt_DocTxtUser) 
						AdviseDoc((ksDocumentFileNotify_Event)doc, docType, false, 
							false, 
							false, 
							false, 
							true, 
							false,
							-1);
					else
						kompas.ksError("Нет активного 2D, 3D документа или документа спецификации");
					break;

				case 20 : // Отписаться
					doc = kompas.ksGetDocumentByReference(0);
					docType = kompas.ksGetDocumentType(0);
					if (doc != null && 
						docType != (int)DocType.lt_DocTxtStandart && 
						docType != (int)DocType.lt_DocTxtUser) 
						BaseEvent.TerminateEvents(typeof(SpecificationEvent), doc, -1, null); 
					else
						kompas.ksError("Нет активного 2D, 3D документа или документа спецификации");
					break;
					#endregion

					#region События объекта спецификации
    				// Подписаться
				case 21 : // Текущий объект
					doc = kompas.ksGetDocumentByReference(0);
					docType = kompas.ksGetDocumentType(0);
					// Документ не должен быть текстовым (2D, 3D документ или документ спецификации)
					ksSpecification specification = GetSpecification();
					reference refSpcObj = specification.ksGetCurrentSpcObject();
					if (refSpcObj == 0)
						kompas.ksError("Нет текущего объекта спецификации");
					if (doc != null && 
						docType != (int)DocType.lt_DocTxtStandart && 
						docType != (int)DocType.lt_DocTxtUser) 
						AdviseDoc(doc, docType, false, 
							false, 
							false, 
							false, 
							false, 
							true, 
							refSpcObj);
					else
						kompas.ksError("Нет активного 2D, 3D документа или документа спецификации");
					break;

				case 22 : // Все объекты
				case SPC_BASE_OBJECT + 350 : // Базовый объект
				case SPC_COMMENT     + 350 : // Комментарий
					doc = kompas.ksGetDocumentByReference(0);
					docType = kompas.ksGetDocumentType(0);
					// Документ не должен быть текстовым (2D, 3D документ или документ спецификации)
					if (doc != null && 
						docType != (int)DocType.lt_DocTxtStandart && 
						docType != (int)DocType.lt_DocTxtUser) 
					{
						int objType = command == 22 ? 0 : command - 350; 
						AdviseDoc(doc, docType, false, 
							false, 
							false, 
							false, 
							false, 
							true, objType);
					}
					else
						kompas.ksError("Нет активного 2D, 3D документа или документа спецификации");
					break;

					// Отписаться    
				case 23 : // Текущий объект
					doc = kompas.ksGetDocumentByReference(0);
					docType = kompas.ksGetDocumentType(0);
					// Документ не должен быть текстовым (2D, 3D документ или документ спецификации)
					if (doc != null && 
						docType != (int)DocType.lt_DocTxtStandart && 
						docType != (int)DocType.lt_DocTxtUser) 
					{
						specification = GetSpecification();
						if (specification != null)
						{
							refSpcObj = specification.ksGetCurrentSpcObject();
							if (refSpcObj == 0)
								kompas.ksError("Нет текущего объекта спецификации");
							else
								BaseEvent.TerminateEvents(typeof(SpcObjectEvent), doc, refSpcObj, null);        
						}
					}
					else
						kompas.ksError("Нет активного 2D, 3D документа или документа спецификации");
					break;

				case 24 : // Все объекты
				case SPC_BASE_OBJECT + 400 : // Базовый объект
				case SPC_COMMENT     + 400 : // Комментарий
					doc = kompas.ksGetDocumentByReference(0);
					docType = kompas.ksGetDocumentType(0);
					if (doc != null && 
						docType != (int)DocType.lt_DocTxtStandart && 
						docType != (int)DocType.lt_DocTxtUser) 
					{
						int objType = command == 24 ? 0 : command - 400; 
						BaseEvent.TerminateEvents(typeof(SpcObjectEvent), doc, objType, null); 
					}
					else
						kompas.ksError("Нет активного 2D, 3D документа или документа спецификации");
					break;

				case 25 : // Отписаться от всех объектов спецификации
					doc = kompas.ksGetDocumentByReference(0);
					docType = kompas.ksGetDocumentType(0);
					if (doc != null && 
						docType != (int)DocType.lt_DocTxtStandart && 
						docType != (int)DocType.lt_DocTxtUser) 
						BaseEvent.TerminateEvents(typeof(SpcObjectEvent), doc, -1, null); 
					else
						kompas.ksError("Нет активного 2D, 3D документа или документа спецификации");
					break;
					#endregion

					#region События 3D документа
				case 26 : // Подписаться
					doc3D = (ksDocument3D)kompas.ActiveDocument3D();
					docType = kompas.ksGetDocumentType(0);
					if (doc3D != null) 
						AdviseDoc((ksDocumentFileNotify_Event)doc3D, docType, false, 
							false, 
							false, 
							true, 
							false, 
							false,
							-1);
					else
						kompas.ksError("Нет активного 3D документа");
					break;

				case 27 : // Отписаться
					doc3D = (ksDocument3D)kompas.ActiveDocument3D();
					if (doc3D != null) 
						BaseEvent.TerminateEvents(typeof(Document3DEvent), doc3D, -1, null); 
					else
						kompas.ksError("Нет активного 3D документа");
					break;
					#endregion

					#region События документа спецификации
				case 28 : // Подписаться
					doc = kompas.ksGetDocumentByReference(0);
					docType = kompas.ksGetDocumentType(0);
					if (doc != null && 
						(docType == (int)DocType.lt_DocSpcUser ||
						docType == (int)DocType.lt_DocSpc)) 
						AdviseDoc((ksDocumentFileNotify_Event)doc, docType, false, 
							false, 
							false, 
							true, 
							false, 
							false,
							-1);
					else
						kompas.ksError("Нет активного документа спецификации");
					break;

				case 29 : // Отписаться
					doc = kompas.ksGetDocumentByReference(0);
					docType = kompas.ksGetDocumentType(0);
					if (doc != null && 
						(docType == (int)DocType.lt_DocSpcUser ||
						docType == (int)DocType.lt_DocSpc)) 
						BaseEvent.TerminateEvents(typeof(SpcDocumentEvent), doc, -1, null); 
					else
						kompas.ksError("Нет активного документа спецификации");
					break;
					#endregion

					#region События 2D документа
				case 30 : // Подписаться
					doc2D = (ksDocument2D)kompas.ActiveDocument2D();
					docType = kompas.ksGetDocumentType(0);
					if (doc2D != null) 
						AdviseDoc(doc2D, docType, false, 
							false, 
							false, 
							true, 
							false, 
							false,
							-1);
					else
						kompas.ksError("Нет активного 2D документа");
					break;

				case 31 : // Отписаться
					doc2D = (ksDocument2D)kompas.ActiveDocument2D();
					if (doc2D != null) 
						BaseEvent.TerminateEvents(typeof(Document2DEvent), doc2D, -1, null); 
					else
						kompas.ksError("Нет активного 2D документа");
					break;
					#endregion

					#region Вспомогательные
				case 32 : // Отписаться от всех событий
					BaseEvent.TerminateEvents();
					break;

				case 33 : // Конфигурация
					FrmConfig.Instance.ShowDialog();
					break;

				case 34 : // Подписанные события
					BaseEvent.ListEvents();
					break; 
					#endregion
      
					#region Общее
				default :
					// События объекта 3D документа
					// Подписаться
					if (command >= (int)Obj3dType.o3d_unknown + 50 && command <= (int)Obj3dType.o3d_feature + 50)
					{
						doc3D = (ksDocument3D)kompas.ActiveDocument3D();
						docType = kompas.ksGetDocumentType(0);
						if (doc3D != null) 
						{
							int objType = command - 50;
							AdviseDoc(doc3D, docType, false, 
								true,
								false, 
								false, 
								false, 
								false, objType);
						}
					}
					else
						kompas.ksError("Нет активного 3D документа");
      
					// Отписаться
					if (command >= (int)Obj3dType.o3d_unknown + 200 && command <= (int)Obj3dType.o3d_feature + 200)
					{
						if (command >= (int)Obj3dType.o3d_unknown + 50 && command <= (int)Obj3dType.o3d_feature + 50)
						{
							doc3D = (ksDocument3D)kompas.ActiveDocument3D();
							docType = kompas.ksGetDocumentType(0);
							if (doc3D != null) 
							{
								int objType = command - 50;
								BaseEvent.TerminateEvents(typeof(Object3DEvent), doc3D, objType, null); 
							}
						}
						else
							kompas.ksError("Нет активного 3D документа");
					}
					break;
					#endregion
			}
		}


		// Формирование меню библиотеки
		[return: MarshalAs(UnmanagedType.BStr)]	public string ExternalMenuItem(short number, ref short itemType, ref short command)
		{
			string result = string.Empty;	//По уполчанию - пустая строка
			itemType = 1;					//MENUITEM
			command = -1;

			switch (number)
			{
				case 1:
					itemType = 2;			//POPUP
					result = "События Компаса";
					break;
				case 2:
					result = "Подписаться";
					command = 1;
					break;
				case 3:
					result = "Отписаться";
					command = 2;
					break;
				case 4:
					itemType = 3;			//ENDMENU
					break;
				case 5:
					itemType = 2;			//POPUP
					result = "События документов";
					break;
				case 6:
					result = "Подписаться";
					command = 3;
					break;
				case 7:
					result = "Отписаться от всех событий документа";
					command = 4;
					break;
				case 8:
					itemType = 3;			//ENDMENU
					break;
				case 9:
					itemType = 2;			//POPUP
					result = "События объекта 2D документа";
					break;
				case 10:
					result = "Подписаться";
					command = 5;
					break;
				case 11:
					result = "Отписаться";
					command = 6;
					break;
				case 12:
					result = "Отписаться от всех событий объектов 2D";
					command = 7;
					break;
				case 13:
					itemType = 0;			//SEPARATOR
					break;
				case 14:
					result = "Подписаться на события вида по номеру";
					command = 8;
					break;
				case 15:
					result = "Отписаться от событий для вида по номеру";
					command = 9;
					break;
				case 16:
					itemType = 0;			//SEPARATOR
					break;
				case 17:
					result = "Подписаться на события слоя по номеру";
					command = 10;
					break;
				case 18:
					result = "Отписаться от событий для слоя по номеру";
					command = 11;
					break;
				case 19:
					itemType = 3;			//ENDMENU
					break;
				case 20:
					itemType = 2;			//POPUP
					result = "События селектирования";
					break;
				case 21:
					result = "Подписаться";
					command = 12;
					break;
				case 22:
					result = "Отписаться";
					command = 13;
					break;
				case 23:
					itemType = 3;			//ENDMENU
					break;
				case 24:
					itemType = 2;			//POPUP
					result = "События штампа";
					break;
				case 25:
					result = "Подписаться";
					command = 14;
					break;
				case 26:
					result = "Отписаться";
					command = 15;
					break;
				case 27:
					itemType = 3;			//ENDMENU
					break;
				case 28:
					itemType = 2;			//POPUP
					result = "События объекта 3D документа";
					break;
				case 29:
					itemType = 2;			//POPUP
					result = "Подписаться";
					break;
				case 30:
					result = "Указать объект в модели";
					command = 16;
					break;
				case 31:
					itemType = 0;			//SEPARATOR
					break;
				case 32:
					result = "Все объекты";
					command = (short)Obj3dType.o3d_unknown + 50;
					break;
				case 33:
					result = "Плоскость XOY";
					command = (short)Obj3dType.o3d_planeXOY + 50;
					break;
				case 34:
					result = "Плоскость XOZ";
					command = (short)Obj3dType.o3d_planeXOZ + 50;
					break;
				case 35:
					result = "Плоскость YOZ";
					command = (short)Obj3dType.o3d_planeYOZ + 50;
					break;
				case 36:
					result = "Точка начала системы координат";
					command = (short)Obj3dType.o3d_pointCS + 50;
					break;
				case 37:
					result = "Эскиз";
					command = (short)Obj3dType.o3d_sketch + 50;
					break;
				case 38:
					result = "Ось по двум плоскостям";
					command = (short)Obj3dType.o3d_axis2Planes + 50;
					break;
				case 39:
					result = "Ось по двум точкам";
					command = (short)Obj3dType.o3d_axis2Points + 50;
					break;
				case 40:
					result = "Ось конической грани";
					command = (short)Obj3dType.o3d_axisConeFace + 50;
					break;
				case 41:
					result = "Ось проходящая через ребро";
					command = (short)Obj3dType.o3d_axisEdge + 50;
					break;
				case 42:
					result = "Ось операции";
					command = (short)Obj3dType.o3d_axisOperation + 50;
					break;
				case 43:
					result = "Смещённая плоскость";
					command = (short)Obj3dType.o3d_planeOffset + 50;
					break;
				case 44:
					result = "Плоскость под углом";
					command = (short)Obj3dType.o3d_planeAngle + 50;
					break;
				case 45:
					result = "Плоскость по 3-м точкам";
					command = (short)Obj3dType.o3d_plane3Points + 50;
					break;
				case 46:
					result = "Нормальная плоскость";
					command = (short)Obj3dType.o3d_planeNormal + 50;
					break;
				case 47:
					result = "Касательная плоскость";
					command = (short)Obj3dType.o3d_planeTangent + 50;
					break;
				case 48:
					result = "Плоскость через ребро и вершину";
					command = (short)Obj3dType.o3d_planeEdgePoint + 50;
					break;
				case 49:
					result = "Плоскость через вершину параллельно другой плоскости";
					command = (short)Obj3dType.o3d_planeParallel + 50;
					break;
				case 50:
					result = "Плоскость через вершину перпендикулярно ребру";
					command = (short)Obj3dType.o3d_planePerpendicular + 50;
					break;
				case 51:
					result = "Плоскость через ребро пар-но/пер-но другому ребру";
					command = (short)Obj3dType.o3d_planeLineToEdge + 50;
					break;
				case 52:
					result = "Плоскость через ребро пар-но/пер-но грани";
					command = (short)Obj3dType.o3d_planeLineToPlane + 50;
					break;
				case 53:
					result = "Базовая операция выдавливания";
					command = (short)Obj3dType.o3d_baseExtrusion + 50;
					break;
				case 54:
					result = "Приклеивание выдавливанием";
					command = (short)Obj3dType.o3d_bossExtrusion + 50;
					break;
				case 55:
					result = "Вырезать выдавливанием";
					command = (short)Obj3dType.o3d_cutExtrusion + 50;
					break;
				case 56:
					result = "Базовая операция вращения";
					command = (short)Obj3dType.o3d_baseRotated + 50;
					break;
				case 57:
					result = "Приклеивание вращением";
					command = (short)Obj3dType.o3d_bossRotated + 50;
					break;
				case 58:
					result = "Вырезать вращением";
					command = (short)Obj3dType.o3d_cutRotated + 50;
					break;
				case 59:
					result = "Базовая операция по сечениям";
					command = (short)Obj3dType.o3d_baseLoft + 50;
					break;
				case 60:
					result = "Приклеивание по сечениям";
					command = (short)Obj3dType.o3d_bossLoft + 50;
					break;
				case 61:
					result = "Вырезать по сечениям";
					command = (short)Obj3dType.o3d_cutLoft + 50;
					break;
				case 62:
					result = "Операция \"Фаска\"";
					command = (short)Obj3dType.o3d_chamfer + 50;
					break;
				case 63:
					result = "Операция \"Скругления\"";
					command = (short)Obj3dType.o3d_fillet + 50;
					break;
				case 64:
					result = "Операция копирования по сетке";
					command = (short)Obj3dType.o3d_meshCopy + 50;
					break;
				case 65:
					result = "Операция копирования по концентрической сетке";
					command = (short)Obj3dType.o3d_circularCopy + 50;
					break;
				case 66:
					result = "Операция копирования по кривой";
					command = (short)Obj3dType.o3d_curveCopy + 50;
					break;
				case 67:
					result = "Операция массив по концентрической сетке для сборки";
					command = (short)Obj3dType.o3d_circPartArray + 50;
					break;
				case 68:
					result = "Операция массив по сетке для сборки";
					command = (short)Obj3dType.o3d_meshPartArray + 50;
					break;
				case 69:
					result = "Операция массив по кривой для сборки";
					command = (short)Obj3dType.o3d_curvePartArray + 50;
					break;
				case 70:
					result = "Операция массив по образцу для сборки";
					command = (short)Obj3dType.o3d_derivPartArray + 50;
					break;
				case 71:
					result = "Операция \"Уклон\"";
					command = (short)Obj3dType.o3d_incline + 50;
					break;
				case 72:
					result = "Операция \"Оболочка\"";
					command = (short)Obj3dType.o3d_shellOperation + 50;
					break;
				case 73:
					result = "Операция \"Ребро жесткости\"";
					command = (short)Obj3dType.o3d_ribOperation + 50;
					break;
				case 74:
					result = "Кинематическая операция";
					command = (short)Obj3dType.o3d_baseEvolution + 50;
					break;
				case 75:
					result = "Приклеить кинематически";
					command = (short)Obj3dType.o3d_bossEvolution + 50;
					break;
				case 76:
					result = "Вырезать кинематически";
					command = (short)Obj3dType.o3d_cutEvolution + 50;
					break;
				case 77:
					result = "Операция \"Зеркальный массив\"";
					command = (short)Obj3dType.o3d_mirrorOperation + 50;
					break;
				case 78:
					result = "Операция \"Зеркально отразить все\"";
					command = (short)Obj3dType.o3d_mirrorAllOperation + 50;
					break;
				case 79:
					result = "Операция \"Сечение поверхностью\"";
					command = (short)Obj3dType.o3d_cutByPlane + 50;
					break;
				case 80:
					result = "Операция \"Сечение эскизом\"";
					command = (short)Obj3dType.o3d_cutBySketch + 50;
					break;
				case 81:
					result = "Отверстие";
					command = (short)Obj3dType.o3d_holeOperation + 50;
					break;
				case 82:
					result = "Ломаная";
					command = (short)Obj3dType.o3d_polyline + 50;
					break;
				case 83:
					result = "Коническая спираль";
					command = (short)Obj3dType.o3d_conicSpiral + 50;
					break;
				case 84:
					result = "Сплайн";
					command = (short)Obj3dType.o3d_spline + 50;
					break;
				case 85:
					result = "Цилиндрическая спираль";
					command = (short)Obj3dType.o3d_cylindricSpiral + 50;
					break;
				case 86:
					result = "Импортирванная поверхность";
					command = (short)Obj3dType.o3d_importedSurface + 50;
					break;
				case 87:
					result = "Условное изображение резьбы";
					command = (short)Obj3dType.o3d_thread + 50;
					break;
				case 88:
					result = "Компонента";
					command = (short)Obj3dType.o3d_part + 50;
					break;
				case 89:
					result = "Объект дерева";
					command = (short)Obj3dType.o3d_feature + 50;
					break;
				case 90:
					itemType = 3;			//ENDMENU
					break;
				case 91:
					itemType = 2;			//POPUP
					result = "Отписаться";
					break;
				case 92:
					result = "Указать объект в модели";
					command = 16;
					break;
				case 93:
					itemType = 0;			//SEPARATOR
					break;
				case 94:
					result = "Все объекты";
					command = (short)Obj3dType.o3d_unknown + 50;
					break;
				case 95:
					result = "Плоскость XOY";
					command = (short)Obj3dType.o3d_planeXOY + 50;
					break;
				case 96:
					result = "Плоскость XOZ";
					command = (short)Obj3dType.o3d_planeXOZ + 50;
					break;
				case 97:
					result = "Плоскость YOZ";
					command = (short)Obj3dType.o3d_planeYOZ + 50;
					break;
				case 98:
					result = "Точка начала системы координат";
					command = (short)Obj3dType.o3d_pointCS + 50;
					break;
				case 99:
					result = "Эскиз";
					command = (short)Obj3dType.o3d_sketch + 50;
					break;
				case 100:
					result = "Ось по двум плоскостям";
					command = (short)Obj3dType.o3d_axis2Planes + 50;
					break;
				case 101:
					result = "Ось по двум точкам";
					command = (short)Obj3dType.o3d_axis2Points + 50;
					break;
				case 102:
					result = "Ось конической грани";
					command = (short)Obj3dType.o3d_axisConeFace + 50;
					break;
				case 103:
					result = "Ось проходящая через ребро";
					command = (short)Obj3dType.o3d_axisEdge + 50;
					break;
				case 104:
					result = "Ось операции";
					command = (short)Obj3dType.o3d_axisOperation + 50;
					break;
				case 105:
					result = "Смещённая плоскость";
					command = (short)Obj3dType.o3d_planeOffset + 50;
					break;
				case 106:
					result = "Плоскость под углом";
					command = (short)Obj3dType.o3d_planeAngle + 50;
					break;
				case 107:
					result = "Плоскость по 3-м точкам";
					command = (short)Obj3dType.o3d_plane3Points + 50;
					break;
				case 108:
					result = "Нормальная плоскость";
					command = (short)Obj3dType.o3d_planeNormal + 50;
					break;
				case 109:
					result = "Касательная плоскость";
					command = (short)Obj3dType.o3d_planeTangent + 50;
					break;
				case 110:
					result = "Плоскость через ребро и вершину";
					command = (short)Obj3dType.o3d_planeEdgePoint + 50;
					break;
				case 111:
					result = "Плоскость через вершину параллельно другой плоскости";
					command = (short)Obj3dType.o3d_planeParallel + 50;
					break;
				case 112:
					result = "Плоскость через вершину перпендикулярно ребру";
					command = (short)Obj3dType.o3d_planePerpendicular + 50;
					break;
				case 113:
					result = "Плоскость через ребро пар-но/пер-но другому ребру";
					command = (short)Obj3dType.o3d_planeLineToEdge + 50;
					break;
				case 114:
					result = "Плоскость через ребро пар-но/пер-но грани";
					command = (short)Obj3dType.o3d_planeLineToPlane + 50;
					break;
				case 115:
					result = "Базовая операция выдавливания";
					command = (short)Obj3dType.o3d_baseExtrusion + 50;
					break;
				case 116:
					result = "Приклеивание выдавливанием";
					command = (short)Obj3dType.o3d_bossExtrusion + 50;
					break;
				case 117:
					result = "Вырезать выдавливанием";
					command = (short)Obj3dType.o3d_cutExtrusion + 50;
					break;
				case 118:
					result = "Базовая операция вращения";
					command = (short)Obj3dType.o3d_baseRotated + 50;
					break;
				case 119:
					result = "Приклеивание вращением";
					command = (short)Obj3dType.o3d_bossRotated + 50;
					break;
				case 120:
					result = "Вырезать вращением";
					command = (short)Obj3dType.o3d_cutRotated + 50;
					break;
				case 121:
					result = "Базовая операция по сечениям";
					command = (short)Obj3dType.o3d_baseLoft + 50;
					break;
				case 122:
					result = "Приклеивание по сечениям";
					command = (short)Obj3dType.o3d_bossLoft + 50;
					break;
				case 123:
					result = "Вырезать по сечениям";
					command = (short)Obj3dType.o3d_cutLoft + 50;
					break;
				case 124:
					result = "Операция \"Фаска\"";
					command = (short)Obj3dType.o3d_chamfer + 50;
					break;
				case 125:
					result = "Операция \"Скругления\"";
					command = (short)Obj3dType.o3d_fillet + 50;
					break;
				case 126:
					result = "Операция копирования по сетке";
					command = (short)Obj3dType.o3d_meshCopy + 50;
					break;
				case 127:
					result = "Операция копирования по концентрической сетке";
					command = (short)Obj3dType.o3d_circularCopy + 50;
					break;
				case 128:
					result = "Операция копирования по кривой";
					command = (short)Obj3dType.o3d_curveCopy + 50;
					break;
				case 129:
					result = "Операция массив по концентрической сетке для сборки";
					command = (short)Obj3dType.o3d_circPartArray + 50;
					break;
				case 130:
					result = "Операция массив по сетке для сборки";
					command = (short)Obj3dType.o3d_meshPartArray + 50;
					break;
				case 131:
					result = "Операция массив по кривой для сборки";
					command = (short)Obj3dType.o3d_curvePartArray + 50;
					break;
				case 132:
					result = "Операция массив по образцу для сборки";
					command = (short)Obj3dType.o3d_derivPartArray + 50;
					break;
				case 133:
					result = "Операция \"Уклон\"";
					command = (short)Obj3dType.o3d_incline + 50;
					break;
				case 134:
					result = "Операция \"Оболочка\"";
					command = (short)Obj3dType.o3d_shellOperation + 50;
					break;
				case 135:
					result = "Операция \"Ребро жесткости\"";
					command = (short)Obj3dType.o3d_ribOperation + 50;
					break;
				case 136:
					result = "Кинематическая операция";
					command = (short)Obj3dType.o3d_baseEvolution + 50;
					break;
				case 137:
					result = "Приклеить кинематически";
					command = (short)Obj3dType.o3d_bossEvolution + 50;
					break;
				case 138:
					result = "Вырезать кинематически";
					command = (short)Obj3dType.o3d_cutEvolution + 50;
					break;
				case 139:
					result = "Операция \"Зеркальный массив\"";
					command = (short)Obj3dType.o3d_mirrorOperation + 50;
					break;
				case 140:
					result = "Операция \"Зеркально отразить все\"";
					command = (short)Obj3dType.o3d_mirrorAllOperation + 50;
					break;
				case 141:
					result = "Операция \"Сечение поверхностью\"";
					command = (short)Obj3dType.o3d_cutByPlane + 50;
					break;
				case 142:
					result = "Операция \"Сечение эскизом\"";
					command = (short)Obj3dType.o3d_cutBySketch + 50;
					break;
				case 143:
					result = "Отверстие";
					command = (short)Obj3dType.o3d_holeOperation + 50;
					break;
				case 144:
					result = "Ломаная";
					command = (short)Obj3dType.o3d_polyline + 50;
					break;
				case 145:
					result = "Коническая спираль";
					command = (short)Obj3dType.o3d_conicSpiral + 50;
					break;
				case 146:
					result = "Сплайн";
					command = (short)Obj3dType.o3d_spline + 50;
					break;
				case 147:
					result = "Цилиндрическая спираль";
					command = (short)Obj3dType.o3d_cylindricSpiral + 50;
					break;
				case 148:
					result = "Импортирванная поверхность";
					command = (short)Obj3dType.o3d_importedSurface + 50;
					break;
				case 149:
					result = "Условное изображение резьбы";
					command = (short)Obj3dType.o3d_thread + 50;
					break;
				case 150:
					result = "Компонента";
					command = (short)Obj3dType.o3d_part + 50;
					break;
				case 151:
					result = "Объект дерева";
					command = (short)Obj3dType.o3d_feature + 50;
					break;
				case 152:
					itemType = 3;			//ENDMENU
					break;
				case 153:
					result = "Отписаться от всех объектов 3D документа";
					command = 18;
					break;
				case 154:
					itemType = 3;			//ENDMENU
					break;
				case 155:
					itemType = 2;			//POPUP
					result = "События спецификации";
					break;
				case 156:
					result = "Подписаться";
					command = 19;
					break;
				case 157:
					result = "Отписаться";
					command = 20;
					break;
				case 158:
					itemType = 3;			//ENDMENU
					break;
				case 159:
					itemType = 2;			//POPUP
					result = "События объекта спецификации";
					break;
				case 160:
					itemType = 2;			//POPUP
					result = "Подписаться";
					command = 0;
					break;
				case 161:
					result = "Текущий объект";
					command = 21;
					break;
				case 162:
					itemType = 0;			//SEPARATOR
					break;
				case 163:
					result = "Все объекты";
					command = 22;
					break;
				case 164:
					result = "Базовый объект";
					command = ldefin2d.SPC_BASE_OBJECT + 350;
					break;
				case 165:
					result = "Комментарий";
					command = ldefin2d.SPC_COMMENT + 350;
					break;
				case 166:
					itemType = 3;			//ENDMENU
					break;
				case 167:
					itemType = 2;			//POPUP
					result = "Отписаться";
					break;
				case 168:
					result = "Текущий объект";
					command = 23;
					break;
				case 169:
					itemType = 0;			//SEPARATOR
					break;
				case 170:
					result = "Все объекты";
					command = 24;
					break;
				case 171:
					result = "Базовый объект";
					command = ldefin2d.SPC_BASE_OBJECT + 400;
					break;
				case 172:
					result = "Комментарий";
					command = ldefin2d.SPC_COMMENT + 400;
					break;
				case 173:
					itemType = 3;			//ENDMENU
					break;
				case 174:
					result = "Отписаться от всех объектов спецификации";
					command = 25;
					break;
				case 175:
					itemType = 3;			//ENDMENU
					break;
				case 176:
					itemType = 2;			//POPUP
					result = "События 3D документа";
					break;
				case 177:
					result = "Подписаться";
					command = 26;
					break;
				case 178:
					result = "Отписаться";
					command = 27;
					break;
				case 179:
					itemType = 3;			//ENDMENU
					break;
				case 180:
					itemType = 2;			//POPUP
					result = "События документа спецификации";
					break;
				case 181:
					result = "Подписаться";
					command = 28;
					break;
				case 182:
					result = "Отписаться";
					command = 29;
					break;
				case 183:
					itemType = 3;			//ENDMENU
					break;
				case 184:
					itemType = 2;			//POPUP
					result = "События 2D документа";
					break;
				case 185:
					result = "Подписаться";
					command = 30;
					break;
				case 186:
					result = "Отписаться";
					command = 31;
					break;
				case 187:
					itemType = 3;			//ENDMENU
					break;
				case 188:
					result = "Отписаться от всех";
					command = 32;
					break;
				case 189:
					result = "Конфигурация";
					command = 33;
					break;
				case 190:
					result = "Подписанные события";
					command = 34;
					break;
				case 191:
					itemType = 3;			//ENDMENU
					break;
			}

			return result;
		}


		// Подписка на события
		public bool LibInterfaceNotifyEntry(object application)
		{
			bool result = true;

			if (FrmConfig.Instance.chbAutoAdvise.Checked)
			{
				// Захват интерфейса приложения КОМПАС
				if (kompas == null && application != null)
				{
					kompas = (KompasObject)application;
					Global.Kompas = kompas;
				}

				if (kompas != null) 
				{
					// Обработчик событий приложения КОМПАС
					ksKompasObjectNotify_Event kompasNotify = (ksKompasObjectNotify_Event)application;
					ApplicationEvent aplEvent = new ApplicationEvent(application, true);        

					// Подписка на события приложения КОМПАС
					aplEvent.Advise();
    				AdviseDocuments();
				}
			}

			return result;
		}


		// Выбор 2D объекта для подписки
		bool Request2DObject(ref int objType)
		{
			objType = -1;
			ksDocument2D doc2D = (ksDocument2D)kompas.ActiveDocument2D();
			ksRequestInfo info = (ksRequestInfo)kompas.GetParamStruct((short)StructType2DEnum.ko_RequestInfo);
			if (info != null && doc2D != null) 
			{
				info.Init();
	
				info.menuId = IDM_REQUEST_OBJECT_2D;
				info.title = "Укажите объект";

				double x = 0, y = 0;
				int command = doc2D.ksCursor(info, ref x, ref y, null);
				if (command != 0)
				{
					if (command == -1)
					{
						reference refObject = doc2D.ksFindObj(x, y, 1000);
						if (doc2D.ksExistObj(refObject) == 1) 
						{
							objType = refObject; // Объект
							return true;
						}
					}
					else
					{
						objType = --command; // Номер объекта
						return true;
					}
				}
			}
			return false; // Отмена
		}


		// Выбор объекта для подписки
		bool Request3DObject(ref ksFeature obj3D, ref int objType)
		{
			obj3D = null;
			objType = (int)Obj3dType.o3d_unknown;
			bool res = true;
			ksDocument3D doc3D = (ksDocument3D)kompas.ActiveDocument3D();
			ksEntity entity = doc3D != null ? (ksEntity)doc3D.UserSelectEntity(null, "", "Укажите объект", 0, null) : null;
			if (entity != null)
			{
				objType = entity.type;
				if (objType == (int)Obj3dType.o3d_face || objType == (int)Obj3dType.o3d_edge || objType == (int)Obj3dType.o3d_vertex) 
				{
					int resYesNo = kompas.ksYesNo("Данный тип не поддерживается\nДа - Подписаться на Part\nНет - Подписаться на Feature\nОтмена - не подписываться");
					switch (resYesNo) 
					{
						case 1 :    
							objType = (int)Obj3dType.o3d_part;  
							obj3D = (ksFeature)entity.GetParent();
							res = true;
							break;
						case 0: 
							objType = (int)Obj3dType.o3d_feature;           
							obj3D = (ksFeature)entity.GetFeature();
							res = true;
							break;
					}
				}
			}
			return res;
		}

		private ksSpecification GetSpecification()
		{
			ksSpecification specification = null;

			object doc = kompas.ksGetDocumentByReference(0);
			int docType = kompas.ksGetDocumentType(0);

			switch (docType) 
			{
				case (int)DocType.lt_DocSheetStandart : // Чертеж стандартный
				case (int)DocType.lt_DocSheetUser :     // Чертеж нестандартный
				case (int)DocType.lt_DocFragment :      // Фрагмент
				{
					ksDocument2D doc2D = (ksDocument2D)doc; // Интерфейс документа
					specification = (ksSpecification)doc2D.GetSpecification();
					break;
				}
				case (int)DocType.lt_DocPart3D :     // 3d-документ модель
				case (int)DocType.lt_DocAssemble3D : // 3d-документ сборка
				{
					ksDocument3D doc3D = (ksDocument3D)doc; // Интерфейс документа
					specification = (ksSpecification)doc3D.GetSpecification(); 
					break;
				}
				case (int)DocType.lt_DocSpc :		// Спецификация
				case (int)DocType.lt_DocSpcUser :	// Спецификация нестандартный формат
				{
					ksSpcDocument spcDoc = (ksSpcDocument)doc;	// Интерфейс документа
					specification = (ksSpecification)spcDoc.GetSpecification();
					break;
				}
			}

			return specification;
		}

		void AdviseDocuments() 
		{
			if (kompas != null)	
			{
				// Интерфейс итератора
				ksIterator iter	= (ksIterator)kompas.GetIterator();		  

				// Создаем итератор	для	хождения по	документам
				if (iter !=	null &&	iter.ksCreateIterator(134, 0))
				{
					reference refDoc = iter.ksMoveIterator("F"); //	Первый документ
					while (refDoc != 0)
					{
						// Подписка	на события документа
						AdviseDoc((ksDocumentFileNotify_Event)kompas.ksGetDocumentByReference(refDoc), 
							kompas.ksGetDocumentType(refDoc),
							true, true,	true,
							true, true,	true, -1);

						// Cледующий документ
						refDoc = iter.ksMoveIterator("N");			 
					}
		
					// Удалить итератор
					iter.ksDeleteIterator();					   
				}
			}
		}


		void AdviseDoc(object doc, int docType, 
			bool fSelectMng		/*true*/, 
			bool fObject		/*true*/, 
			bool fStamp			/*true*/,
			bool fDocument		/*true*/,
			bool fSpecification	/*true*/,
			bool fSpcObject		/*true*/,
			int objType			/*-1*/) 
		{ 
			if (doc == null)
				return;

			ksEntity objEntity = null;

			// События документа, необходимы для своевременной отписки
			if (!BaseEvent.FindEvent(typeof(DocumentEvent), doc, -1, null)) 
			{
				bool fFileDoc = !fSelectMng && !fObject && !fStamp && !fDocument && !fSpecification && !fSpcObject;

				// Обработчик событий от документа
				DocumentEvent docEvent = new DocumentEvent((ksDocumentFileNotify_Event)doc, fFileDoc); 
				// Подписка на события документа
				int advise = docEvent.Advise();

				if (!BaseEvent.FindEvent(typeof(DocumentFrameEvent), (object)doc, 0, null))
				{
					if (doc != null)
					{
						KompasAPI7.IApplication appl = (KompasAPI7.IApplication)kompas.ksGetApplication7();
						KompasAPI7.IKompasDocument kDoc = appl.ActiveDocument;
						KompasAPI7.DocumentFrames frames = kDoc.DocumentFrames;
						KompasAPI7.DocumentFrame frame = frames[0];
						DocumentFrameEvent frmEvent = new DocumentFrameEvent((KompasAPI7.ksDocumentFrameNotify_Event)frame, (object)doc, true);
						frmEvent.Advise();
					}
				}


				// Неудачная подписка на события документа
				if (advise == 0)
					return;
			}
			else
				kompas.ksError("На события документа уже подписались");

			switch (docType) 
			{
				case (int)DocType.lt_DocSheetStandart :		// 1 - чертеж стандартный
				case (int)DocType.lt_DocSheetUser :			// 2 - чертеж нестандартный
				case (int)DocType.lt_DocFragment :			// 3 - фрагмент
					ksDocument2D doc2D = (ksDocument2D)doc;	// Интерфейс документа

					// Документ 2D
					if (fDocument) 
					{
						if (!BaseEvent.FindEvent(typeof(Document2DEvent), doc2D, -1, null))
						{ 
							object doc2DNotify = doc2D.GetDocument2DNotify();
							if (doc2DNotify != null)
							{
								Document2DEvent document2DEvent = new Document2DEvent(doc2DNotify, doc2D, true); 
								document2DEvent.Advise();                                               
							}
						}
						else
							kompas.ksError("На события 3D документа уже подписались");
					}

					// Спецификация в 2D документе
					if (fSpecification) 
					{
						if (!BaseEvent.FindEvent(typeof(SpecificationEvent), doc2D, -1, null))
						{ 
							ksSpecificationNotify specification = (ksSpecificationNotify)doc2D.GetSpecification();
							if (specification != null)
							{
								SpecificationEvent specificationEvent = new SpecificationEvent(specification, doc2D, true); 
								specificationEvent.Advise();                                               
							}
						}
						else
							kompas.ksError("На события спецификации уже подписались");
					}

					// Объект спецификации
					if (fSpcObject) 
					{ 
						ksSpecification specification = doc2D != null ? (ksSpecification)doc2D.GetSpecification() : null;
						if (specification != null)
						{
							reference refSpcObj = specification.ksGetCurrentSpcObject();
							if (refSpcObj == 0)
								kompas.ksError("Нет текущего объекта спецификации");
							else
							{
								if (!BaseEvent.FindEvent(typeof(SpcObjectEvent), doc2D, refSpcObj, null)) 
								{       
									ksSpcObjectNotify objNotify = specification != null ? (ksSpcObjectNotify)specification.GetSpcObjectNotify(refSpcObj) : null; 
									if (objNotify != null)
									{  
										SpcObjectEvent objEvent = new SpcObjectEvent(objNotify, doc2D, refSpcObj, true); 
										objEvent.Advise();
									}
								}
								else
									kompas.ksError("На события объекта спецификации (2D) уже подписались");
							}
						}
					}   

					// Селектирование
					if (fSelectMng) 
					{ 
						if (!BaseEvent.FindEvent(typeof(SelectMngEvent), doc2D, -1, null))
						{ 
							object selMsg = doc2D != null ? doc2D.GetSelectionMngNotify() : null;
							if (selMsg != null)
							{
								SelectMngEvent selEvent = new SelectMngEvent(selMsg, doc2D, true);
								selEvent.Advise();                                               
							}
						}
						else
							kompas.ksError("На события селектирования уже подписались");
					}

					// Штамп
					if (fStamp && docType != (int)DocType.lt_DocFragment)
					{
						if (!BaseEvent.FindEvent(typeof(StampEvent), doc2D, -1, null)) 
						{
							object stamp = doc2D.GetStamp(); 
							if (stamp != null)
							{
								StampEvent stampEvent = new StampEvent(stamp, doc2D, true);
								stampEvent.Advise();
							}
						}
						else
							kompas.ksError("На события редактирования штампа уже подписались");
					}

					// Объект 2D документа
					if (fObject && objType >= 0) // Тип приходит всегда
					{ 
						if (!BaseEvent.FindEvent(typeof(Object2DEvent), doc2D, objType, null)) 
						{          
							object objNotify = doc2D != null ? doc2D.GetObject2DNotify(objType) : null;
							if (objNotify != null) 
							{  
								Object2DEvent objEvent = new Object2DEvent(objNotify, doc2D, objType, doc2D.GetObject2DNotifyResult(), true);
								objEvent.Advise();
							}
						}
						else
							kompas.ksError("На события объекта 2D документа уже подписались");
					}  
					break;      

				case (short)DocType.lt_DocPart3D :			// 5 - 3d-документ модель
				case (short)DocType.lt_DocAssemble3D :		// 6 - 3d-документ сборка
					ksDocument3D doc3D = (ksDocument3D)doc;	// Интерфейс документа

					// Документ 3D
					if (fDocument) 
					{
						if (!BaseEvent.FindEvent(typeof(Document3DEvent), doc3D, -1, null))
						{ 
							object doc3DNotify = doc3D.GetDocument3DNotify();
							if (doc3DNotify != null)
							{
								Document3DEvent document3DEvent = new Document3DEvent(doc3DNotify, doc3D, true); 
								document3DEvent.Advise();                                               
							}
						}
						else
							kompas.ksError("На события 3D документа уже подписались");
					}

					// Спецификация в 3D документе
					if (fSpecification) 
					{
						if (!BaseEvent.FindEvent(typeof(SpecificationEvent), doc3D, -1, null))
						{ 
							ksSpecificationNotify specification = (ksSpecificationNotify)doc3D.GetSpecification();
							if (specification != null) 
							{
								SpecificationEvent specificationEvent = new SpecificationEvent(specification, doc3D, true); 
								specificationEvent.Advise();                                               
							}
						}
						else
							kompas.ksError("На события спецификации уже подписались");
					}

					// Объект спецификации
					if (fSpcObject) 
					{ 
						ksSpecification specification = doc3D != null ? (ksSpecification)doc3D.GetSpecification() : null;
						if (specification != null) 
						{
							reference refSpcObj = specification.ksGetCurrentSpcObject();
							if (refSpcObj == 0)
								kompas.ksError("Нет текущего объекта спецификации");
							else
							{
								if (!BaseEvent.FindEvent(typeof(SpcObjectEvent), doc3D, refSpcObj, null))
								{
									ksSpcObjectNotify objNotify = specification != null ? (ksSpcObjectNotify)specification.GetSpcObjectNotify(refSpcObj) : null;
									if (objNotify != null) 
									{  
										SpcObjectEvent objEvent = new SpcObjectEvent(objNotify, doc3D, refSpcObj, true); 
										objEvent.Advise();
									}
								}
								else
									kompas.ksError("На события объекта спецификации (3D) уже подписались");
							}
						}
					}  

					// Селектирование
					if (fSelectMng)
					{
						if (!BaseEvent.FindEvent(typeof(SelectMngEvent), doc3D, -1, null)) 
						{
							ksSelectionMngNotify selMng = doc3D != null ? (ksSelectionMngNotify)doc3D.GetSelectionMng() : null;
							if (selMng != null)
							{
								SelectMngEvent selEvent = new SelectMngEvent(selMng, doc3D, true); 
								selEvent.Advise();                                              
							}
						}
						else
							kompas.ksError("На события селектирования уже подписались");
					}
      
					// Объект 3D документа
					if (fObject) 
					{ 
						if (objType >= 0)
						{
							if (!BaseEvent.FindEvent(typeof(Object3DEvent), doc3D, objType, null)) 
							{  
								ksPart containerPart = (ksPart)doc3D.GetPart((int)Part_Type.pTop_Part);
								object objNotify = containerPart != null ? containerPart.GetObject3DNotify(objType, null) : null;
								if (objNotify != null)
								{ 
									Object3DEvent objEvent = new Object3DEvent(objNotify, doc3D, objType, null, containerPart.GetObject3DNotifyResult(), true);
									objEvent.Advise();
								}
							}
							else
								kompas.ksError("На события объекта 3D документа уже подписались");
						}
						else
						{
							ksFeature obj3D = null;
							if (Request3DObject(ref obj3D, ref objType))
							{  
								if (!BaseEvent.FindEvent(typeof(Object3DEvent), doc3D, objType, obj3D)) 
								{  
									ksPart containerPart = null;
									switch (objType) 
									{          
										case (int)Obj3dType.o3d_part: 
											ksPart objPart = (ksPart)obj3D;
											if (objPart != null)
												containerPart = (ksPart)objPart.GetPart((int)Part_Type.pTop_Part);
											break;
                
										case (int)Obj3dType.o3d_feature: 
											ksFeature objFeature = (ksFeature)obj3D;
											if (objFeature != null)
											{ 
												object obj = objFeature.GetObject();
												if (obj != null)
												{
													objEntity = (ksEntity)obj;
													if (objEntity != null)
														containerPart = (ksPart)objEntity.GetParent();
													else 
													{
														objPart = (ksPart)obj;
														if (objPart != null)
															containerPart = (ksPart)objPart.GetPart((int)Part_Type.pTop_Part);
													}
												}
											}
											break;

										default: 
											objEntity = (ksEntity)obj3D;
											if (objEntity != null)
												containerPart = (ksPart)objEntity.GetParent();
											break;
									}       

									if (containerPart == null) 
										containerPart = (ksPart)doc3D.GetPart((int)Part_Type.pTop_Part);
            
									object objNotify = containerPart != null ? containerPart.GetObject3DNotify(objType, obj3D) : null;
									if (objNotify != null) 
									{ 
										Object3DEvent objEvent = new Object3DEvent(objNotify, doc3D, objType, obj3D, containerPart.GetObject3DNotifyResult(), true); 
										objEvent.Advise();
									}
								}
								else
									kompas.ksError("На события объекта 3D документа уже подписались");
							}
						}
					}
					break; 

				case (int)DocType.lt_DocTxtStandart :	// 7 - текстовый документ стандартный
				case (int)DocType.lt_DocTxtUser :		// 8 - текстовый документ нестандартный
					ksDocumentTxt docTxt = (ksDocumentTxt)doc; // Интерфейс документа

					// Штамп
					if (fStamp)
					{
						if (!BaseEvent.FindEvent(typeof(StampEvent), docTxt, -1, null)) 
						{
							ksStampNotify stamp = (ksStampNotify)docTxt.GetStamp(); 
							if (stamp != null)
							{
								StampEvent stampEvent = new StampEvent(stamp, docTxt, true);
								stampEvent.Advise();
							}
						}
						else
							kompas.ksError("На события редактирования штампа уже подписались");
					}
					break;

				case (int)DocType.lt_DocSpcUser :	// 9 - спецификация нестандартный формат
				case (int)DocType.lt_DocSpc :		// 4 - спецификация
				{
					ksSpcDocument spcDoc = (ksSpcDocument)doc;	// Интерфейс документа

					// Документ спецификации
					if (fDocument) 
					{
						if (!BaseEvent.FindEvent(typeof(SpcDocumentEvent), spcDoc, -1, null))
						{
							object spcDocNotify = spcDoc.GetSpcDocumentNotify();
							if (spcDocNotify != null)
							{
								SpcDocumentEvent spcDocumentEvent = new SpcDocumentEvent(spcDocNotify, spcDoc, true);
								spcDocumentEvent.Advise();
							}
						}
						else
							kompas.ksError("На события документа спецификации уже подписались");
					}

					// Спецификация в документе спецификации
					if (fSpecification) 
					{
						if (!BaseEvent.FindEvent(typeof(SpecificationEvent), spcDoc, -1, null))
						{ 
							object specification = spcDoc.GetSpecification();
							if (specification != null) 
							{
								SpecificationEvent specificationEvent = new SpecificationEvent(specification, spcDoc, true);
								specificationEvent.Advise();                                               
							}
						}
						else
							kompas.ksError("На события спецификации уже подписались");
					}

					// Объект документа спецификации
					if (fSpcObject) 
					{ 
						ksSpecification specification = spcDoc != null ? (ksSpecification)spcDoc.GetSpecification() : null;
						if (specification != null)
						{
							if (!BaseEvent.FindEvent(typeof(SpcObjectEvent), spcDoc, objType, null)) 
							{       
								object objNotify = specification != null ? specification.GetSpcObjectNotify(objType) : null;
								if (objNotify != null)
								{  
									SpcObjectEvent objEvent = new SpcObjectEvent(objNotify, spcDoc, objType, true); 
									objEvent.Advise();
								}
							}
							else
								kompas.ksError("На события объекта спецификации (Spc) уже подписались");
						}
					}  

					// Штамп
					if (fStamp)
					{
						if (!BaseEvent.FindEvent(typeof(StampEvent), spcDoc, -1, null))
						{
							ksStampNotify stamp = (ksStampNotify)spcDoc.GetStamp();
							if (stamp != null)
							{
								StampEvent stampEvent = new StampEvent(stamp, spcDoc, true);
								stampEvent.Advise();
							}
						}
						else
							kompas.ksError("На события редактирования штампа уже подписались");
					}

					break;
				}
			}
		}

    #region Реализаця интерфейса IDisposable
    public void Dispose()
    {
      if (kompas != null)
      {
        Marshal.ReleaseComObject(Global.Kompas);
        GC.SuppressFinalize(Global.Kompas);
        kompas = null;
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
