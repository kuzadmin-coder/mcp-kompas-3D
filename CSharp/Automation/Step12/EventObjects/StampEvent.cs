////////////////////////////////////////////////////////////////////////////////
//
// StampEvent  - обработчик событий штампа
//
//////////////////////////////////////////////////////////////////////////////////

using Kompas6API5;

using System;
using Kompas6Constants;
using KAPITypes;

using reference = System.Int32;

namespace Steps.NET
{
	public class StampEvent : BaseEvent, ksStampNotify
	{
		public StampEvent(object obj, object doc)
			: base(obj, typeof(ksStampNotify).GUID, doc) {}

		// kdBeginEditStamp - Начало редактирования штампа
		public bool BeginEditStamp()
		{
			return true;
		}


		// kdEndEditStamp - Завершение редактирования штампа
		public bool EndEditStamp(bool editResult)
		{
			Global.UpdateSlideBox(null);
			return true;  
		}


		// kdStampCellDblClick - ДаблКлик по штампу
		public bool StampCellDblClick(int number)
		{
			return true; 
		}


		// kdStampCellBeginEdit - Начало редактирования ячейки штампа
		public bool StampCellBeginEdit(int number)
		{
			return true; 
		}

    // kdStampCellBeginEdit - Начало редактирования ячейки штампа
    public bool StampBeginClearCells(object numbers)
    {
      return true; 
    }			

    // Создать обработчик событий для штампа
		public static BaseEvent NewStampEvent(object doc, int docType)
		{
			StampEvent res = null;
			if (doc != null)
			{
				object stamp = null;
				ksDocument2D doc2d = null;
				ksSpcDocument spcDoc = null;
				ksDocumentTxt txtDoc = null;

				switch (docType)
				{
					case (int)DocType.lt_DocSheetStandart :
					case (int)DocType.lt_DocSheetUser :
						doc2d = (ksDocument2D)doc;
						stamp = doc2d.GetStamp();
						break;
					case (int)DocType.lt_DocSpc :
					case (int)DocType.lt_DocSpcUser :
						spcDoc = (ksSpcDocument)doc;
						stamp = spcDoc.GetStamp();
						break;
					case (int)DocType.lt_DocTxtStandart :
					case (int)DocType.lt_DocTxtUser : 
						txtDoc = (ksDocumentTxt)doc;
						stamp = txtDoc.GetStamp();
						break;
				}

				if (stamp != null && !BaseEvent.FindEvent(typeof(StampEvent), doc)) 
				{
					DocumentEvent.NewDocumentEvent(doc);	// чтобы при закрытии документа отписаться
					res = new StampEvent(stamp, doc);
					res.Advise();
				}
			}

			return res;
		}
 
	}
}