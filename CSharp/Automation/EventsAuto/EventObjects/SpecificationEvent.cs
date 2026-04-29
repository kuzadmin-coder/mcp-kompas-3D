////////////////////////////////////////////////////////////////////////////////
//
// SpecificationEvent  - обработчик событий от спецификации
//
////////////////////////////////////////////////////////////////////////////////

using Kompas6API5;
using System;
using Kompas6Constants;
using KAPITypes;

namespace Steps.NET
{
	public class SpecificationEvent : BaseEvent, ksSpecificationNotify
	{
		public SpecificationEvent(object obj, object doc, bool selfAdvise)
			: base(obj, typeof(ksSpecificationNotify).GUID, doc,
			-1, null, selfAdvise) {}

		// ssTuningSpcStyleBeginChange - Начало изменения настроек спецификации
		public bool TuningSpcStyleBeginChange(string libName, int numb)
		{ 
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> TuningSpcStyleBeginChange\nlibName = {1}\nnumb = {2}", m_LibName, libName, numb); 
				str += "\nИмя документа = " + GetDocName();
				return Global.Kompas.ksYesNo(str) == 1 ? true : false;
			}
			return true; 
		}


		// ssTuningSpcStyleChange - Настроейки спецификации изменились
		public bool TuningSpcStyleChange(string libName, int numb, bool isOk)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> TuningSpcStyleChange\nlibName = {1}\nnumb = {2}\nisOk = {3}", m_LibName, libName, numb, isOk ? "true" : "false");
				str += "\nИмя документа = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true;      
		}


		// ssChangeCurrentSpcDescription - Изменилось текущее описание спецификации
		public bool ChangeCurrentSpcDescription(string libName, int numb)
		{ 
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> ChangeCurrentSpcDescription\nlibName = {1}\nnumb = {2}", m_LibName, libName, numb); 
				str += "\nИмя документа = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true; 
		}


		// ssSpcDescriptionAdd - Добавилось описание спецификации
		public bool SpcDescriptionAdd(string libName, int numb)
		{ 
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> SpcDescriptionAdd\nlibName = {1}\nnumb = {2}", m_LibName, libName, numb); 
				str += "\nИмя документа = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true;
		}


		// ssSpcDescriptionRemove - Удалилось описание спецификации
		public bool SpcDescriptionRemove(string libName, int numb)
		{ 
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> SpcDescriptionRemove\nlibName = {1}\nnumb = {2}", m_LibName, libName, numb); 
				str += "\nИмя документа = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true;
		}


		// ssSpcDescriptionBeginEdit - Начало редактирования описания спецификации
		public bool SpcDescriptionBeginEdit(string libName, int numb)
		{ 
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> SpcDescriptionBeginEdit\nlibName = {1}\nnumb = {2}", m_LibName, libName, numb); 
				str += "\nИмя документа = " + GetDocName();
				return Global.Kompas.ksYesNo(str) == 1 ? true : false;
			}
			return true; 
		}


		// ssSpcDescriptionEdit - Отредактировали описание спецификации
		public bool SpcDescriptionEdit(string libName, int numb, bool isOk)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> SpcDescriptionEdit\nlibName = {1}\nnumb = {2}\nisOk = {3}", m_LibName, libName, numb, isOk ? "true" : "false");
				str += "\nИмя документа = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true; 
		}


		// ssSynchronizationBegin - Начало синхронизации
		public bool SynchronizationBegin()
		{ 
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecEvents.Checked)
			{
				string str = m_LibName + " --> SynchronizationBegin";
				str += "\nИмя документа = " + GetDocName();
				return Global.Kompas.ksYesNo(str) == 1 ? true : false;
			}
			return true;
		}


		// ssSynchronization - Синхронизация проведена
		public bool Synchronization()
		{ 
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecEvents.Checked)
			{
				string str = m_LibName + " --> Synchronization";
				str += "\nИмя документа = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true;
		}


		// ssBeginCalcPositions - Начало  расчета позиций
		public bool BeginCalcPositions()
		{ 
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecEvents.Checked)
			{
				string str = m_LibName + " --> BeginCalcPositions";
				str += "\nИмя документа = " + GetDocName();
				return Global.Kompas.ksYesNo(str) == 1 ? true : false;
			}
			return true;
		}


		// ssCalcPositions - Проведен расчет позиций 
		public bool CalcPositions()
		{ 
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecEvents.Checked)
			{
				string str = m_LibName + " --> CalcPositions";
				str += "\nИмя документа = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true;
		}


		// ssBeginCreateObject - Начало создания объекта спецификации (до диалога выбора раздела) 
		public bool BeginCreateObject(int typeObj)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> BeginCreateObject\ntypeObj = {1}", m_LibName, typeObj);
				str += "\nИмя документа = " + GetDocName();
				return Global.Kompas.ksYesNo(str) == 1 ? true : false;
			}
			return true;
		}
	}
}
