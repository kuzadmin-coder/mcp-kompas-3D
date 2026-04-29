////////////////////////////////////////////////////////////////////////////////
//
// StampEvent  - обработчик событий штампа
//
//////////////////////////////////////////////////////////////////////////////////

using Kompas6API5;
using System;
using Kompas6Constants;
using KAPITypes;

namespace Steps.NET
{
	public class StampEvent : BaseEvent, ksStampNotify
	{
		public StampEvent(object obj, object doc, bool selfAdvise)
			: base(obj, typeof(ksStampNotify).GUID, doc,
			-1, null, selfAdvise) {}

		// kdBeginEditStamp - Начало редактирования штампа
		public bool BeginEditStamp()
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbStampEvents.Checked)
			{
				string str = m_LibName + " --> BeginEditStamp";
				str += "\nИмя документа = " + GetDocName();
				return Global.Kompas.ksYesNo(str) == 1 ? true : false;
			}
			return true;
		}


		// kdEndEditStamp - Завершение редактирования штампа
		public bool EndEditStamp(bool editResult)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbStampEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> EndEditStamp\neditResult = {1}", m_LibName, editResult ? "true" : "false");
				str += "\nИмя документа = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true;  
		}


		// kdStampCellDblClick - ДаблКлик по штампу
		public bool StampCellDblClick(int number)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbStampEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> StampCellDblClick\nnumber = {1}", m_LibName, number);
				str += "\nИмя документа = " + GetDocName();
				return Global.Kompas.ksYesNo(str) == 1 ? true : false;
			}
			return true; 
		}


		// kdStampCellBeginEdit - Начало редактирования ячейки штампа
		public bool StampCellBeginEdit(int number)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbStampEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> StampCellBeginEdit\nnumber = {1}", m_LibName, number);
				str += "\nИмя документа = " + GetDocName();
				return Global.Kompas.ksYesNo(str) == 1 ? true : false;
			}
			return true; 
		}	

    // kdStampCellBeginEdit - Начало редактирования ячейки штампа
    public bool StampBeginClearCells(object numbers)
    {
      return true; 
    }	
	}
}
