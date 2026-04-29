////////////////////////////////////////////////////////////////////////////////
//
// SelectMngEvent  - обработчик событий от менеджера селектирования документа
//
////////////////////////////////////////////////////////////////////////////////

using Kompas6API5;

using System;
using Kompas6Constants;
using KAPITypes;

namespace Steps.NET
{
	public class SelectMngEvent : BaseEvent, ksSelectionMngNotify
	{
		public SelectMngEvent(object obj, object doc)
			: base(obj, typeof(ksSelectionMngNotify).GUID, doc) {}

		// ksmSelect - Объект селектирован
		public bool Select(object obj) 
		{ 
			Global.UpdateSlideBox(null);
			return true;
		}


		// ksmUnselect - Объект расселектирован
		public bool Unselect(object obj)
		{
			Global.UpdateSlideBox(null);
			return true;
		}


		// ksmUnselectAll - Все объекты расселектированы
		public bool UnselectAll() 
		{
			Global.UpdateSlideBox(null);
			return true;
		}	
	}
}
