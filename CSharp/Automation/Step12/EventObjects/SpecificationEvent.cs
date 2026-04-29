////////////////////////////////////////////////////////////////////////////////
//
// SpecificationEvent  - обработчик событий от спецификации
//
////////////////////////////////////////////////////////////////////////////////

using Kompas6API5;

using System;
using System.Reflection;
using Kompas6Constants;
using KAPITypes;

using reference = System.Int32;

namespace Steps.NET
{
	public class SpecificationEvent : BaseEvent, ksSpecificationNotify
	{
		public SpecificationEvent(object obj, object doc)
			: base(obj, typeof(ksSpecificationNotify).GUID, doc) {}

		// ssTuningSpcStyleBeginChange - Начало изменения настроек спецификации
		public bool TuningSpcStyleBeginChange(string libName, int numb)
		{ 
			return true; 
		}


		// ssTuningSpcStyleChange - Настроейки спецификации изменились
		public bool TuningSpcStyleChange(string libName, int numb, bool isOk)
		{
			Global.UpdateSlideBox(null);
			return true;      
		}


		// ssChangeCurrentSpcDescription - Изменилось текущее описание спецификации
		public bool ChangeCurrentSpcDescription(string libName, int numb)
		{ 
			//Global.UpdateSlideBox();
			return true; 
		}


		// ssSpcDescriptionAdd - Добавилось описание спецификации
		public bool SpcDescriptionAdd(string libName, int numb)
		{ 
			return true;
		}


		// ssSpcDescriptionRemove - Удалилось описание спецификации
		public bool SpcDescriptionRemove(string libName, int numb)
		{ 
			return true;
		}


		// ssSpcDescriptionBeginEdit - Начало редактирования описания спецификации
		public bool SpcDescriptionBeginEdit(string libName, int numb)
		{ 
			return true; 
		}


		// ssSpcDescriptionEdit - Отредактировали описание спецификации
		public bool SpcDescriptionEdit(string libName, int numb, bool isOk)
		{
			return true; 
		}


		// ssSynchronizationBegin - Начало синхронизации
		public bool SynchronizationBegin()
		{ 
			return true;
		}


		// ssSynchronization - Синхронизация проведена
		public bool Synchronization()
		{ 
			return true;
		}


		// ssBeginCalcPositions - Начало  расчета позиций
		public bool BeginCalcPositions()
		{ 
			return true;
		}


		// ssCalcPositions - Проведен расчет позиций 
		public bool CalcPositions()
		{ 
			Global.UpdateSlideBox(null);
			return true;
		}


		// ssBeginCreateObject - Начало создания объекта спецификации (до диалога выбора раздела) 
		public bool BeginCreateObject(int typeObj)
		{
			return true;
		}


		// Создать обработчик события редактирования спецификации
		public static BaseEvent NewSpecificationEvent(object doc)
		{
			SpecificationEvent res = null;
			if (doc != null)
			{
				if (!BaseEvent.FindEvent(typeof(SpecificationEvent), doc))
				{
					object spcObj = doc.GetType().InvokeMember("GetSpecification", BindingFlags.InvokeMethod, null, doc, null);
					ksSpecification spc = (ksSpecification)spcObj;
					if (spc != null)
					{
						DocumentEvent.NewDocumentEvent(doc);	// чтобы при закрытии документа отписаться
						res = new SpecificationEvent(spc, doc);
						res.Advise();
					}
				}
			}
			return res;

		}

	}
}