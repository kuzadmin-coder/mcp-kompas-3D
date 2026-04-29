////////////////////////////////////////////////////////////////////////////////
//
// SpcObjectEvent - обработчик событий объектов документа спецификации
//
////////////////////////////////////////////////////////////////////////////////

using Kompas6API5;

using System;
using System.Reflection;
using Kompas6Constants;
using KAPITypes;

namespace Steps.NET
{
	public class SpcObjectEvent : BaseEvent, ksSpcObjectNotify
	{
		public SpcObjectEvent(object obj, object doc, int objType)
			: base(obj, typeof(ksSpcObjectNotify).GUID, doc, objType, 0) {}

		// koBeginDelete - Попытка удаления объекта
		public bool BeginDelete(int objRef)
		{
			return true;
		}


		// koDelete - Объект удален
		public bool Delete(int objRef)
		{
			Global.UpdateSlideBox(null);
			return true;
		}
       

		// soCellDblClick - Двойной клик в ячейке 
		public bool CellDblClick(int objRef, int number)
		{ 
			return true;  
		}


		// soCellBeginEdit - Начало редактирования в ячейке   
		public bool CellBeginEdit(int objRef, int number) 
		{ 
			return true;
		}


		// soChangeCurrent - Изменился текущий объект  
		public bool ChangeCurrent(int objRef)
		{ 
			return true;  
		}


		// soDocumentBeginAdd - Начало добавления документа
		public bool DocumentBeginAdd(int objRef)
		{ 
			return true;
		}


		// soDocumentAdd - Добавление документа в объекте СП   
		public bool DocumentAdd(int objRef, string docName) 
		{ 
			return true; 
		}


		// soDocumentRemove - Удаление документа из объекта СП  
		public bool DocumentRemove(int objRef, string docName)
		{ 
			return true;
		}


		// soBeginGeomChange - Начало измения геометрии объекта СП
		public bool BeginGeomChange(int objRef)
		{ 
			return true;
		}


		// soGeomChange - Геометрия объекта СП изменилась    
		public bool GeomChange(int objRef) 
		{
			return true;  
		}


		// koBeginProcess - Начало редактирования\создания объекта
		public bool BeginProcess(int pType, int objRef)
		{
			return true;  
		}
       

		// koEndProcess - Конец редактирования\создания объекта
		public bool EndProcess(int pType)
		{
			Global.UpdateSlideBox(null);
			return true; 
		}


		// koCreate - Создание объектов
		public bool CreateObject(int objRef)
		{
			Global.UpdateSlideBox(null);
			return true; 
		}

    
		// koUpdateObject - Редактирование объекта
		public bool UpdateObject(int objRef)
		{
			Global.UpdateSlideBox(null);
			return true;
		} 


		// Создать обработчик событий для объекта спецификации
		public static BaseEvent NewSpcObjectEvent(object doc, int objType)
		{
			SpcObjectEvent res = null;
			if (doc != null) 
			{
				if (!BaseEvent.FindEvent(typeof(SpcObjectEvent), doc, objType)) 
				{
					object spcObj = doc.GetType().InvokeMember("GetSpecification", BindingFlags.InvokeMethod, null, doc, null);
					ksSpecification spc = (ksSpecification)spcObj;
					if (spc != null)
					{
						DocumentEvent.NewDocumentEvent(doc); // чтобы при закрытии документа отписаться
						res = new SpcObjectEvent(spc.GetSpcObjectNotify(objType), doc, objType);
						res.Advise();
					}
				}
			}
			return res;
		}
    
    // koBeginCopy - Копирование объекта
    public bool BeginCopy( int objRef )
    {
      return true;
    }

    // koCopy - Копирование объекта
    public bool copy( int objRef )
    {
      return true;
    }


	}
}