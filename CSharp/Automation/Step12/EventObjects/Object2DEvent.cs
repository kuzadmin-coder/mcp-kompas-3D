////////////////////////////////////////////////////////////////////////////////
//
// Object2DEvent - обработчик событий объектов 2D документа
//
////////////////////////////////////////////////////////////////////////////////

using Kompas6API5;


using System;
using Kompas6Constants;
using KAPITypes;

namespace Steps.NET
{
	public class Object2DEvent : BaseEvent, ksObject2DNotify
	{
		public Object2DEvent(object obj, object doc, int objType)
			: base(obj, typeof(ksObject2DNotify).GUID, doc, objType, 0) {}

		// kdChangeActive - Переключение вида/слоя в текущий
		public bool ChangeActive(int viewRef)
		{
			Global.UpdateSlideBox(null);
			return true;
		}


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
       

		// koBeginMove - Начало смещения объекта
		public bool BeginMove(int objRef)
		{
			return true;
		}


		// koMove - Объект смещен
		public bool Move(int objRef)
		{
			Global.UpdateSlideBox(null);
			return true;
		}


		// koBeginRotate - Поворот объекта
		public bool BeginRotate(int objRef)
		{
			return true;
		}


		// koRotate - Поворот объекта
		public bool Rotate(int objRef)
		{
			Global.UpdateSlideBox(null);
			return true;
		}


		// koBeginScale - Маштабирование объекта
		public bool BeginScale(int objRef)
		{
			return true;
		}


		// koScale - Маштабирование объекта
		public bool scale(int objRef)
		{
			Global.UpdateSlideBox(null);
			return true;
		}


		// koBeginTransform - Трансформация объекта
		public bool BeginTransform(int objRef)
		{
			return true;
		}


		// koTransform - Трансформация объекта
		public bool Transform(int objRef)
		{
			Global.UpdateSlideBox(null);
			return true;
		}


		// koBeginCopy - Копирование объекта
		public bool BeginCopy(int objRef)
		{
			return true;
		}


		// koCopy - Копирование объекта
		public bool copy(int objRef)
		{
			Global.UpdateSlideBox(null);
			return true;
		}


		// koBeginSymmetry - Симметрия объекта
		public bool BeginSymmetry(int objRef)
		{
			return true;
		}


		// koSymmetry - Симметрия объекта
		public bool Symmetry(int objRef)
		{
			Global.UpdateSlideBox(null);
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

    // koBeginDestroyObject - Разрушить объект
    public bool BeginDestroyObject(int objRef)
    {
      Global.UpdateSlideBox(null);
      return true;
    }

    
    // koDestroyObject - Разрушить объект
    public bool DestroyObject(int objRef)
    {
      Global.UpdateSlideBox(null);
      return true;
    }

    public bool BeginPropertyChanged(int objRef)
    {
      return true;
    }

    public bool PropertyChanged(int objRef)
    {
      return true;
    }

		// Создать обработчик событий документа
		public static BaseEvent NewObj2DEvent(object doc, int docType, int objType) 
		{
			Object2DEvent res = null;
			if (doc != null) 
			{
				if (!BaseEvent.FindEvent(typeof(Object2DEvent), doc, objType)) 
				{
					ksDocument2D doc2d = (ksDocument2D)doc;
					if (doc != null)
					{
						DocumentEvent.NewDocumentEvent(doc);	// чтобы при закрытии документа отписаться
						res = new Object2DEvent(doc2d.GetObject2DNotify(objType), doc, objType);
						res.Advise();
					}
				}
			}
			return res;
		}
	}
}
