
using Kompas6API5;
using KompasAPI7;


using System;
using System.Drawing;
using System.Reflection;
using System.Collections;
using KAPITypes;
using Kompas6Constants;
using Kompas6Constants3D;

using reference = System.Int32;

namespace Steps.NET
{
	public class Global
	{
		private static PropertyManagerLayout layout = PropertyManagerLayout.pmAlignRight;
		public static PropertyManagerLayout Layout
		{
			get
			{
				return layout;
			}
			set
			{
				layout = value;
			}
		}


		private static Rectangle rectangle = new Rectangle(20, 20, 250, 300);
		public static Rectangle Rectangle
		{
			get
			{
				return rectangle;
			}
			set
			{
				rectangle = value;
			}
		}


		private static ArrayList eventList = new ArrayList();
		public static ArrayList EventList
		{
			get
			{
				return eventList;
			}
			set
			{
				eventList = value;
			}
		}


		private static KompasObject kompas;
		public static KompasObject Kompas
		{
			get
			{
				return kompas;
			}
			set
			{
				kompas = value;
			}
		}


		// Компас Application 7
		private static IApplication newKompasAPI;
		public static IApplication NewKompasAPI
		{
			get
			{
				return newKompasAPI;
			}
			set
			{
				newKompasAPI = value;
			}
		}


		// Менежер панели свойств
		private static IPropertyManager propMng;
		public static IPropertyManager PropMng
		{
			get
			{
				return propMng;
			}
			set
			{
				propMng = value;
			}
		}


		// Окно в котором будет отрисовываться текущий документ
		private static IPropertySlideBox slideBox;
		public static IPropertySlideBox SlideBox
		{
			get
			{
				return slideBox;
			}
			set
			{
				slideBox = value;
			}
		}


		// Обновление скна
		/*
		public static void UpdateSlideBox()
		{
			if (Global.SlideBox != null)
				Global.SlideBox.UpdateParam();
		}
		*/


		// Обновление скна
		/*
		public static void UpdateSlideBox(reference docRef)
		{
			if (Global.SlideBox != null)
			{
				object doc = Global.Kompas.ksGetDocumentByReference(docRef);
				Global.SlideBox.DrawingSlide = GetDocReference(doc);
				Global.SlideBox.Value = docRef != 0;
				Global.SlideBox.CheckBoxVisibility = (docRef != 0);
				Global.SlideBox.UpdateParam();
			}
		}*/


		// Обновление скна
		public static void UpdateSlideBox(object doc)
		{
			if (Global.SlideBox != null)
			{
				if (doc != null)
				{
					reference docRef = GetDocReference(doc);
					Global.SlideBox.DrawingSlide = docRef;
					Global.SlideBox.Value = docRef != 0;
					Global.SlideBox.CheckBoxVisibility = (docRef != 0);
				}
				Global.SlideBox.UpdateParam();
			}
		}


		// Подписка на события документа
		public static void AdviseDoc(object doc, int docType)
		{
			reference docRef = GetDocReference(doc);
			if (docRef != 0)
			{
				if (DocumentEvent.NewDocumentEvent(doc) != null) 
				{
					switch(docType) 
					{
						case (int)DocType.lt_DocSheetStandart :
						case (int)DocType.lt_DocSheetUser :
						case (int)DocType.lt_DocFragment :
							StampEvent.NewStampEvent(doc, docType);
							DocumentEvent.NewDocument2DEvent(docRef);
							Object2DEvent.NewObj2DEvent(doc, docType, ldefin2d.ALL_OBJ);
							break;
						case (int)DocType.lt_DocPart3D :
						case (int)DocType.lt_DocAssemble3D : 
						{
							DocumentEvent.NewDocument3DEvent(docRef);
							Object3DEvent.NewObj3DEvent(doc, docType, (short)Obj3dType.o3d_unknown, null); // На все
							break;
						}
						case (int)DocType.lt_DocSpc :
						case (int)DocType.lt_DocSpcUser : 
						{
							SpcDocumentEvent.NewSpcDocEvent(doc, docType);
							StampEvent.NewStampEvent(doc, docType); 
							break;
						}	
						case (int)DocType.lt_DocTxtStandart :
						case (int)DocType.lt_DocTxtUser : 
							StampEvent.NewStampEvent(doc, docType);
							break;
					}    

					switch (docType) 
					{
						case (int)DocType.lt_DocSheetStandart:
						case (int)DocType.lt_DocSheetUser:
						case (int)DocType.lt_DocSpc:
						case (int)DocType.lt_DocSpcUser:
						case (int)DocType.lt_DocAssemble3D: 
						{
							SpecificationEvent.NewSpecificationEvent(doc);
							SpcObjectEvent.NewSpcObjectEvent(doc, ldefin2d.SPC_BASE_OBJECT /*На базовые объекты*/);
						}
						break;
					}
				}
			}
		}


		// Получение reference по IDispatch документа
		public static reference GetDocReference(object doc)
		{
			reference result = 0;

			object refObj = doc.GetType().InvokeMember("reference", BindingFlags.GetProperty, null, doc, null);
			if (refObj != null)
				result = Convert.ToInt32(refObj);

			return result;
		}

		
		// Создать закладку и подписаться
		public static void CreateAndSubscriptionPropertyManager(bool mes)
		{
			if (Global.PropMng == null)
			{
				reference docRef = 0;
				object doc = Global.Kompas.ksGetDocumentByReference(0);
				int docType = Global.Kompas.ksGetDocumentType(0);

				ApplicationEvent.NewApplicationEvent();   
				Global.PropMng = Global.NewKompasAPI.CreatePropertyManager(true);
				Global.PropMng.Layout = Global.Layout;
				Global.PropMng.Caption = "Панель из библиотеки";
				Global.PropMng.SetGabaritRect(Global.Rectangle.X, Global.Rectangle.Y, Global.Rectangle.Right, Global.Rectangle.Bottom);
				Global.PropMng.SpecToolbar = SpecPropertyToolBarEnum.pnEnterEscCreateSaveSearchHelp;
				new PropertyManagerEvent(Global.PropMng);	// Подписываемся на события процесса 
				IPropertyTab tab = Global.PropMng.PropertyTabs.Add("Закладка из библиотеки");
    
				// Получаем коллекцию контролов
				IPropertyControls collection = tab.PropertyControls;
				AdviseDoc(doc, docType);
				// Добавсляем окно просмотра
				slideBox = (IPropertySlideBox)collection.Add(ControlTypeEnum.ksControlSlideBox);
				slideBox.SlideType = SlideTypeEnum.ksKompasDocument;	// Отрисовываем группу
				slideBox.DrawingSlide = docRef;							// Передаем reference группы 
				slideBox.Value = docRef != 0;
				slideBox.CheckBoxVisibility = (docRef != 0);
				slideBox.Hint = "Hint для слайда";
				slideBox.Tips = "Tips для слайда";
				slideBox.Id = 10000;
				slideBox.Name = "Окно просмотра";
				slideBox.NameVisibility = PropertyControlNameVisibility.ksNameHorizontalVisible;
    
				Global.PropMng.ShowTabs();
				Global.PropMng.Caption = "Панель из библиотеки";
				//propMng.ShowTabs();
				//propMng.UpdateTabs();

			}
			else 
			{
				Global.PropMng.Visible = true; // Если панель была скрыта пользователем по крестику востановим ее
				if (mes)
				{
					Global.Kompas.ksMessage("Панель уже загружена");
				}
			}
		}
 

		// Остановить работу с панелью свойствт  и отписаться
		public static void ClosePropertyManager(bool mes)
		{
			if (Global.PropMng != null)
			{
  				Global.Layout = Global.PropMng.Layout;

				int left;
				int top;
				int right;
				int bottom;
				Global.PropMng.GetGabaritRect(out left, out top, out right, out bottom);
				Global.Rectangle = new Rectangle(left, top, right - left, bottom - top);

				BaseEvent.TerminateEvents(typeof(PropertyManagerEvent), 0);
				Global.PropMng.HideTabs();
				Global.SlideBox = null; 
				Global.PropMng = null;
			}
			else 
			{
				if (mes)
					Global.Kompas.ksMessage("Панель уже выгружена");
			}
		}
	}
}
