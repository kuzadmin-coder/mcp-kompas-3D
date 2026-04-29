////////////////////////////////////////////////////////////////////////////////
//
// SpcDocumentEvent  - обработчик событий от документа спецификации
//
////////////////////////////////////////////////////////////////////////////////

using Kompas6API5;

using System;
using Kompas6Constants;
using KAPITypes;

namespace Steps.NET
{
	public class SpcDocumentEvent : BaseEvent, ksSpcDocumentNotify
	{
		public SpcDocumentEvent(object obj, object doc)
			: base(obj, typeof(ksSpcDocumentNotify).GUID, doc) {}

		// sdDocumentBeginAdd - Начало добавления документа сборочного чертежа
		public bool DocumentBeginAdd()
		{
			return true;
		}


		// sdDocumentAdd - Добавление документа сборочного чертежа
		public bool DocumentAdd(string docName)
		{
			return true;
		}
      

		// sdDocumentBeginRemove - Начало удаления документа сборочного чертежа
		public bool DocumentBeginRemove(string docName)
		{
			return true;
		}


		// sdDocumentRemove - Удаление документа сборочного чертежа
		public bool DocumentRemove(string docName)
		{
			return true;
		}


		// sdSpcStyleBeginChange - Начало изменения стиля спецификации
		public bool SpcStyleBeginChange(string libName, int numb)
		{
			return true;
		}


		// sdSpcStyleChange - Стиль спецификации изменился
		public bool SpcStyleChange(string libName, int numb)
		{
			Global.UpdateSlideBox(null);
			return true;
		}	


		// Создать обработчик событий документа cпецификации
		public static BaseEvent NewSpcDocEvent(object doc, int objType)
		{
			SpcDocumentEvent res = null;
			if (doc != null)
			{
				if (!BaseEvent.FindEvent(typeof(SpcDocumentEvent), doc)) 
				{
					ksSpcDocument spcDoc = (ksSpcDocument)doc;
					if (spcDoc != null)
					{
						DocumentEvent.NewDocumentEvent(doc);	// чтобы при закрытии документа отписаться
						res = new SpcDocumentEvent(spcDoc.GetSpcDocumentNotify(), doc);
						res.Advise();
					}
				}
			}
			return res;
		}
	}
}