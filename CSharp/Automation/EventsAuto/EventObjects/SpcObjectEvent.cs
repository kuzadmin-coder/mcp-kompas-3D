////////////////////////////////////////////////////////////////////////////////
//
// SpcObjectEvent - обработчик событий объектов документа спецификации
//
////////////////////////////////////////////////////////////////////////////////

using Kompas6API5;

using System;
using Kompas6Constants;
using KAPITypes;

namespace Steps.NET
{
	public class SpcObjectEvent : BaseEvent, ksSpcObjectNotify
	{
		public SpcObjectEvent(object obj, object doc, int objType, bool selfAdvise)
			: base(obj, typeof(ksSpcObjectNotify).GUID, doc,
			objType, null, selfAdvise) {}

		// koBeginDelete - Попытка удаления объекта
		public bool BeginDelete(int objRef)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecObjEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> BeginDelete\nobjRef = {1}", m_LibName, objRef);
				str += "\nИмя документа = " + GetDocName();
				return Global.Kompas.ksYesNo(str) == 1 ? true : false;
			}
			return true;
		}


		// koDelete - Объект удален
		public bool Delete(int objRef)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecObjEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> Delete\nobjRef = {1}", m_LibName, objRef);
				str += "\nИмя документа = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			// Если удаляется объект удаляем и подписку на его события
			// Либо тип объекта : 0, SPC_BASE_OBJECT, SPC_COMMENT, либо указатель на объект
			if (objRef != 0 && objRef == m_ObjType)
				this.Delete(objRef); 
			return true;
		}
       

		// soCellDblClick - Двойной клик в ячейке 
		public bool CellDblClick(int objRef, int number)
		{ 
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecObjEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> CellDblClick\nobjRef = {1}\nnumber = {2}", m_LibName, objRef, number);
				str += "\nИмя документа = " + GetDocName(); 
				return Global.Kompas.ksYesNo(str) == 1 ? true : false;
			}
			return true;  
		}


		// soCellBeginEdit - Начало редактирования в ячейке   
		public bool CellBeginEdit(int objRef, int number) 
		{ 
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecObjEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> CellBeginEdit\nobjRef = {1}\nnumber = {2}", m_LibName, objRef, number);
				str += "\nИмя документа = " + GetDocName();
				return Global.Kompas.ksYesNo(str) == 1 ? true : false;
			}
			return true;
		}


		// soChangeCurrent - Изменился текущий объект  
		public bool ChangeCurrent(int objRef)
		{ 
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecObjEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> ChangeCurrent\nobjRef = {1}", m_LibName, objRef);
				str += "\nИмя документа = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true;  
		}


		// soDocumentBeginAdd - Начало добавления документа
		public bool DocumentBeginAdd(int objRef)
		{ 
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecObjEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> DocumentBeginAdd\nobjRef = {1}", m_LibName, objRef);
				str += "\nИмя документа = " + GetDocName();
				return Global.Kompas.ksYesNo(str) == 1 ? true : false;
			}
			return true;
		}


		// soDocumentAdd - Добавление документа в объекте СП   
		public bool DocumentAdd(int objRef, string docName) 
		{ 
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecObjEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> DocumentAdd\nobjRef = {1}\ndocName = {2}", m_LibName, objRef, docName);
				str += "\nИмя документа = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true; 
		}


		// soDocumentRemove - Удаление документа из объекта СП  
		public bool DocumentRemove(int objRef, string docName)
		{ 
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecObjEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> DocumentRemove\nobjRef = {1}\ndocName = {2}", m_LibName, objRef, docName);
				str += "\nИмя документа = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true;
		}


		// soBeginGeomChange - Начало измения геометрии объекта СП
		public bool BeginGeomChange(int objRef)
		{ 
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecObjEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> BeginGeomChange\nobjRef = {1}", m_LibName, objRef);
				str += "\nИмя документа = " + GetDocName();
				return Global.Kompas.ksYesNo(str) == 1 ? true : false;
			}
			return true;
		}


		// soGeomChange - Геометрия объекта СП изменилась    
		public bool GeomChange(int objRef) 
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecObjEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> GeomChange\nobjRef = {1}", m_LibName, objRef);
				str += "\nИмя документа = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true;  
		}


		// koBeginProcess - Начало редактирования\создания объекта
		public bool BeginProcess(int pType, int objRef)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecObjEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> BeginProcess\npType = {1}\nobjRef = {2}", m_LibName, pType, objRef);
				str += "\nИмя документа = " + GetDocName();
				return Global.Kompas.ksYesNo(str) == 1 ? true : false;
			}
			return true;  
		}
       

		// koEndProcess - Конец редактирования\создания объекта
		public bool EndProcess(int pType)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecObjEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> EndProcess\npType = {1}", m_LibName, pType);
				str += "\nИмя документа = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true; 
		}


		// koCreate - Создание объектов
		public bool CreateObject(int objRef)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecObjEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> CreateObject\nobjRef = {1}", m_LibName, objRef);
				str += "\nИмя документа = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true; 
		}

    
		// koUpdateObject - Редактирование объекта
		public bool UpdateObject(int objRef)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chbSpecObjEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> UpdateObject\nobjRef = {1}", m_LibName, objRef);
				str += "\nИмя документа = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}   
			return true;
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
