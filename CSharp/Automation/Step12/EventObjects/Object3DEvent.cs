////////////////////////////////////////////////////////////////////////////////
//
// Object3DEvent - обработчик событий объектов 3D документа
//
////////////////////////////////////////////////////////////////////////////////

using Kompas6API5;

using System;
using Kompas6Constants;
using Kompas6Constants3D;
using KAPITypes;

using reference = System.Int32;

namespace Steps.NET
{
	public class Object3DEvent : BaseEvent, ksObject3DNotify
	{
		public Object3DEvent(object obj, object doc, int objType, object obj3d)
			: base(obj, typeof(ksObject3DNotify).GUID, doc, objType, obj3d) {}

		// o3BeginDelete - Начало удаления объектов
		public bool BeginDelete(object obj)
		{
			return true;
		}


		// o3Delete - Oбъекты удалены
		public bool Delete(object obj)
		{
			Global.UpdateSlideBox(null);
			return true;     
		}


		// o3Excluded - Oбъект исключен/включен в расчет
		public bool excluded(object obj, bool excluded)
		{
			Global.UpdateSlideBox(null);
			return true;
		}


		// o3Hidden - Oбъект скрыт/показан
		public bool hidden(object obj, bool _hidden)
		{
			Global.UpdateSlideBox(null);
			return true;
		}


		// o3BeginPropertyChanged - Начало изменения свойств объета
		public bool BeginPropertyChanged(object obj)
		{
			return true;
		}


		// o3PropertyChanged - Изменены свойства объета
		public bool PropertyChanged(object obj)
		{
			Global.UpdateSlideBox(null);
			return true;
		}


		// o3BeginPlacementChanged - Начало изменения положения объета
		public bool BeginPlacementChanged(object obj)
		{
			return true;
		}


		// o3PlacementChanged - Изменено положения объета
		public bool PlacementChanged(object obj)
		{
			Global.UpdateSlideBox(null);
			return true;
		}


		// o3BeginProcess - Начало редактирования\создания объекта
		public bool BeginProcess(int pType, object obj)
		{
			return true;
		}


		// o3EndProcess - Конец редактирования\создания объекта
		public bool EndProcess(int pType)
		{
			Global.UpdateSlideBox(null);
			return true;
		}


		// o3CreateObject - Создание объекта
		public bool CreateObject(object obj)
		{
			Global.UpdateSlideBox(null);
			return true;
		}


		// o3UpdateObject - Редактирование объекта
		public bool UpdateObject(object obj)
		{
			Global.UpdateSlideBox(null);
			return true;
		}

    public bool BeginLoadStateChange(object obj, int loadState)
    {
      return true;
    }
    public bool LoadStateChange(object obj, int loadState)
    {
      Global.UpdateSlideBox(null);
      return true;
    }

    // Найти контейнер для 3D объекта
    private static object GetContainer(reference doc, int objType, object iObj)
		{
			reference docRef = Global.GetDocReference(doc);
			object container = null;
			if (iObj != null) 
			{
				switch (objType)
				{
					case 0: 
					{
						// тип не известен
						ksPart iPart = (ksPart)iObj;
						if (iPart != null)
							container = iPart.GetPart((short)Part_Type.pTop_Part);
						else 
						{
							ksEntity iEntity = (ksEntity)iObj;
							if (iEntity != null)
								container = iEntity.GetParent();
						}
						break;
					}

					case (short)Obj3dType.o3d_part: 
					{
						ksPart iPart = (ksPart)iObj;
						if (iPart != null)
							container = iPart.GetPart((short)Part_Type.pTop_Part);
						break;
					}

					case (short)Obj3dType.o3d_feature: 
					{
						ksFeature iFeature = (ksFeature)iObj;
						if (iFeature != null)
						{ 
							object obj = iFeature.GetObject();
							if (obj != null)
								container = GetContainer(0, 0, obj);
						}
						break;
					} 

					default: 
					{
						ksEntity iEntity = (ksEntity)iObj;
						if (iEntity != null)
							container = iEntity.GetParent();
						break;
					}
				}
			}
			if (container == null)
			{
				if (docRef != 0) 
				{
					ksDocument3D doc3D = (ksDocument3D)Global.Kompas.ksGet3dDocumentFromRef(docRef);
					if (doc3D != null)
						container = doc3D.GetPart((short)Part_Type.pTop_Part);
				}
			}
			return container;	 
		}


		// Создать обработчик событий 3D документа
		public static BaseEvent NewObj3DEvent(object doc, int docType, int objType, object iObj)
		{
			Object3DEvent res = null;
			if (doc != null)
			{
				int typeObj = iObj != null ? 0 : objType; 
				if (!BaseEvent.FindEvent(typeof(Object3DEvent), doc, typeObj, iObj))
				{
					object container = GetContainer(Global.GetDocReference(doc), objType, iObj);
					if (container != null)
					{
						DocumentEvent.NewDocumentEvent(doc);	// чтобы при закрытии документа отписаться
						res = new Object3DEvent(container, doc, typeObj, iObj);
						res.Advise();
					}
				}
			}
			return res;	
		}
	}
}