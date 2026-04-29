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
		private ksObject2DNotifyResult m_res;

		public Object2DEvent(object obj, object doc, int objType,
			ksObject2DNotifyResult res, bool selfAdvise)
			: base(obj, typeof(ksObject2DNotify).GUID, doc,
			objType, null, selfAdvise) {m_res = res;}

		
		string OutRes()
		{
			string str = string.Empty;
			if (m_res != null)
			{
				int notifyType = m_res.GetNotifyType();
				double angle = m_res.GetAngle();
				int copyObject = m_res.GetCopyObject();
				double sx = -1, sy = -1;
				m_res.GetScale(out sx, out sy);
				double x = -1, y = -1, x1 = -1, y1 = -1;
				m_res.GetSheetPoint(false, out x, out y);
				m_res.GetSheetPoint(true, out x1, out y1);
				int copy = m_res.IsCopy() ? 1 : 0;
				str = string.Format("Object2DNotifyResult: GetNotifyType = {0},\nGetAngle = {1}, GetCopyObject = {2},\nGetScale({3}, {4}),\nGetSheetPoint(1, {5}, {6}),\nGetSheetPoint(0, {7}, {8}),\nIsCopy = {9}", 
					notifyType,
					angle, 
					copyObject,               
					sx, sy,
					x1, y1,     
					x, y,
					copy);
			}
			return str;
		}


		// kdChangeActive - Переключение вида/слоя в текущий
		public bool ChangeActive(int viewRef)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chb2DObjEvents.Checked)
			{
				ksDocument2D doc2D = (ksDocument2D)m_Doc;
				if (doc2D != null && viewRef != 0) 
					doc2D.ksLightObj(viewRef, 1);

				string str = string.Empty;
				str = string.Format("{0} --> ChangeActive\nviewRef = {1}\n", m_LibName, viewRef);
				str += OutRes();
				str += "\nИмя документа = " + GetDocName();
				string strType;
				strType = string.Format("\nТип/указатель объекта = {0}", m_ObjType);
				str += strType;
				Global.Kompas.ksMessage(str);

				if (doc2D != null && viewRef != 0) 
					doc2D.ksLightObj(viewRef, 0);
			}
			return true;
		}


		// koBeginDelete - Попытка удаления объекта
		public bool BeginDelete(int objRef)
		{
			bool res = false;
			if (m_SelfAdvise && FrmConfig.Instance.chb2DObjEvents.Checked)
			{
				ksDocument2D doc2D = (ksDocument2D)m_Doc;
				if (doc2D != null && objRef != 0) 
					doc2D.ksLightObj(objRef, 1);

				string str = string.Empty;
				str = string.Format("{0} --> BeginDelete\nobjRef = {1}\n", m_LibName, objRef);
				str += OutRes();
				str += "\nИмя документа = " + GetDocName();
				string strType;
				strType = string.Format("\nТип/указатель объекта = {0}", m_ObjType);
				str += strType;                    
				res = Global.Kompas.ksYesNo(str) == 1 ? true : false;

				if (doc2D != null && objRef != 0) 
					doc2D.ksLightObj(objRef, 0);
			}
			return res;
		}


		// koDelete - Объект удален
		public bool Delete(int objRef)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chb2DObjEvents.Checked)
			{
				ksDocument2D doc2D = (ksDocument2D)m_Doc;
				if (doc2D != null && objRef != 0) 
					doc2D.ksLightObj(objRef, 1);

				string str = string.Empty;
				str = string.Format("{0} --> Delete\nobjRef = {1}\n", m_LibName, objRef);
				str += OutRes();
				str += "\nИмя документа = " + GetDocName();
				string strType;
				strType = string.Format("\nТип/указатель объекта = {0}", m_ObjType);
				str += strType;
				Global.Kompas.ksMessage(str);

				if (doc2D != null && objRef != 0) 
					doc2D.ksLightObj(objRef, 0);
			}
			// Если удаляется объект удаляем и подписку на его события
			// Либо тип объекта : ALL_OBJ...AXISLINE_OBJ, VIEW_OBJ, либо указатель на объект
			if (objRef != 0 && objRef == m_ObjType)
				this.Delete(objRef);
			return true;
		}
       

		// koBeginMove - Начало смещения объекта
		public bool BeginMove(int objRef)
		{
			bool res = false;
			if (m_SelfAdvise && FrmConfig.Instance.chb2DObjEvents.Checked)
			{
				ksDocument2D doc2D = (ksDocument2D)m_Doc;
				if (doc2D != null && objRef != 0) 
					doc2D.ksLightObj(objRef, 1);

				string str = string.Empty;
				str = string.Format("{0} --> BeginMove\nobjRef = {1}\n", m_LibName, objRef);
				str += OutRes();
				str += "\nИмя документа = " + GetDocName();
				string strType;
				strType = string.Format("\nТип/указатель объекта = {0}", m_ObjType);
				str += strType;
				res = Global.Kompas.ksYesNo(str) == 1 ? true : false;

				if (doc2D != null && objRef != 0) 
					doc2D.ksLightObj(objRef, 0);
			}
			return res;
		}


		// koMove - Объект смещен
		public bool Move(int objRef)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chb2DObjEvents.Checked)
			{
				ksDocument2D doc2D = (ksDocument2D)m_Doc;
				if (doc2D != null && objRef != 0) 
					doc2D.ksLightObj(objRef, 1);

				string str = string.Empty;
				str = string.Format("{0} --> Move\nobjRef = {1}\n", m_LibName, objRef);
				str += OutRes();
				str += "\nИмя документа = " + GetDocName();
				string strType;
				strType = string.Format("\nТип/указатель объекта = {0}", m_ObjType);
				str += strType;
				Global.Kompas.ksMessage(str);

				if (doc2D != null && objRef != 0) 
					doc2D.ksLightObj(objRef, 0);
			}
			return true;
		}


		// koBeginRotate - Поворот объекта
		public bool BeginRotate(int objRef)
		{
			bool res = false;
			if (m_SelfAdvise && FrmConfig.Instance.chb2DObjEvents.Checked)
			{
				ksDocument2D doc2D = (ksDocument2D)m_Doc;
				if (doc2D != null && objRef != 0) 
					doc2D.ksLightObj(objRef, 1);

				string str = string.Empty;
				str = string.Format("{0} --> BeginRotate\nobjRef = {1}\n", m_LibName, objRef);
				str += OutRes();
				str += "\nИмя документа = " + GetDocName();
				string strType;
				strType = string.Format("\nТип/указатель объекта = {0}", m_ObjType);
				str += strType;
				res = Global.Kompas.ksYesNo(str) == 1 ? true : false;

				if (doc2D != null && objRef != 0) 
					doc2D.ksLightObj(objRef, 0);
			}
			return res;
		}


		// koRotate - Поворот объекта
		public bool Rotate(int objRef)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chb2DObjEvents.Checked)
			{
				ksDocument2D doc2D = (ksDocument2D)m_Doc;
				if (doc2D != null && objRef != 0) 
					doc2D.ksLightObj(objRef, 1);

				string str = string.Empty;
				str = string.Format("{0} --> Rotate\nobjRef = {1}\n", m_LibName, objRef);
				str += OutRes();
				str += "\nИмя документа = " + GetDocName();
				string strType;
				strType = string.Format("\nТип/указатель объекта = {0}", m_ObjType);
				str += strType;
				Global.Kompas.ksMessage(str);

				if (doc2D != null && objRef != 0) 
					doc2D.ksLightObj(objRef, 0);
			}
			return true;
		}


		// koBeginScale - Маштабирование объекта
		public bool BeginScale(int objRef)
		{
			bool res = false;
			if (m_SelfAdvise && FrmConfig.Instance.chb2DObjEvents.Checked)
			{
				ksDocument2D doc2D = (ksDocument2D)m_Doc;
				if (doc2D != null && objRef != 0) 
					doc2D.ksLightObj(objRef, 1);

				string str = string.Empty;
				str = string.Format("{0} --> BeginScale\nobjRef = {1}\n", m_LibName, objRef);
				str += OutRes();
				str += "\nИмя документа = " + GetDocName();
				string strType;
				strType = string.Format("\nТип объекта/указатель = {0}", m_ObjType);
				str += strType;
				res = Global.Kompas.ksYesNo(str) == 1 ? true : false;

				if (doc2D != null && objRef != 0) 
					doc2D.ksLightObj(objRef, 0);
			}
			return res;
		}


		// koScale - Маштабирование объекта
		public bool scale(int objRef)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chb2DObjEvents.Checked)
			{
				ksDocument2D doc2D = (ksDocument2D)m_Doc;
				if (doc2D != null && objRef != 0) 
					doc2D.ksLightObj(objRef, 1);

				string str = string.Empty;
				str = string.Format("{0} --> Scale\nobjRef = {1}\n", m_LibName, objRef);
				str += OutRes();
				str += "\nИмя документа = " + GetDocName();
				string strType;
				strType = string.Format("\nТип объекта/указатель = {0}", m_ObjType);
				str += strType;
				Global.Kompas.ksMessage(str);

				if (doc2D != null && objRef != 0) 
					doc2D.ksLightObj(objRef, 0);
			}
			return true;
		}


		// koBeginTransform - Трансформация объекта
		public bool BeginTransform(int objRef)
		{
			bool res = false;
			if (m_SelfAdvise && FrmConfig.Instance.chb2DObjEvents.Checked)
			{    
				ksDocument2D doc2D = (ksDocument2D)m_Doc;
				if (doc2D != null && objRef != 0) 
					doc2D.ksLightObj(objRef, 1);

				string str = string.Empty;
				str = string.Format("{0} --> BeginTransform\nobjRef = {1}\n", m_LibName, objRef);
				str += OutRes();
				str += "\nИмя документа = " + GetDocName();
				string strType;
				strType = string.Format("\nТип объекта/указатель = {0}", m_ObjType);
				str += strType;
				res = Global.Kompas.ksYesNo(str) == 1 ? true : false;

				if (doc2D != null && objRef != 0) 
					doc2D.ksLightObj(objRef, 0);
			}
			return res;
		}


		// koTransform - Трансформация объекта
		public bool Transform(int objRef)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chb2DObjEvents.Checked)
			{
				ksDocument2D doc2D = (ksDocument2D)m_Doc;
				if (doc2D != null && objRef != 0) 
					doc2D.ksLightObj(objRef, 1);

				string str = string.Empty;
				str = string.Format("{0} --> Transform\nobjRef = {1}\n", m_LibName, objRef);
				str += OutRes();
				str += "\nИмя документа = " + GetDocName();
				string strType;
				strType = string.Format("\nТип объекта/указатель = {0}", m_ObjType);
				str += strType;
				Global.Kompas.ksMessage(str);

				if (doc2D != null && objRef != 0) 
					doc2D.ksLightObj(objRef, 0);
			}
			return true;
		}


		// koBeginCopy - Копирование объекта
		public bool BeginCopy(int objRef)
		{
			bool res = false;
			if (m_SelfAdvise && FrmConfig.Instance.chb2DObjEvents.Checked)
			{
				ksDocument2D doc2D = (ksDocument2D)m_Doc;
				if (doc2D != null && objRef != 0) 
					doc2D.ksLightObj(objRef, 1);

				string str = string.Empty;
				str = string.Format("{0} --> BeginCopy\nobjRef = {1}\n", m_LibName, objRef);
				str += OutRes();
				str += "\nИмя документа = " + GetDocName();
				string strType;
				strType = string.Format("\nТип объекта/указатель = {0}", m_ObjType);
				str += strType;
				res = Global.Kompas.ksYesNo(str) == 1 ? true : false;

				if (doc2D != null && objRef != 0) 
					doc2D.ksLightObj(objRef, 0);
			}
			return res;
		}


		// koCopy - Копирование объекта
		public bool copy(int objRef)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chb2DObjEvents.Checked)
			{
				ksDocument2D doc2D = (ksDocument2D)m_Doc;
				if (doc2D != null && objRef != 0) 
					doc2D.ksLightObj(objRef, 1);

				string str = string.Empty;
				str = string.Format("{0} --> Copy\nobjRef = {1}\n", m_LibName, objRef);
				str += "\nИмя документа = " + GetDocName();
				string strType;
				strType = string.Format("\nТип объекта/указатель = {0}", m_ObjType);
				str += strType;
				Global.Kompas.ksMessage(str);

				if (doc2D != null && objRef != 0) 
					doc2D.ksLightObj(objRef, 0);
			}
			return true;
		}


		// koBeginSymmetry - Симметрия объекта
		public bool BeginSymmetry(int objRef)
		{
			bool res = false;
			if (m_SelfAdvise && FrmConfig.Instance.chb2DObjEvents.Checked)
			{
				ksDocument2D doc2D = (ksDocument2D)m_Doc;
				if (doc2D != null && objRef != 0) 
					doc2D.ksLightObj(objRef, 1);

				string str = string.Empty;
				str = string.Format("{0} --> BeginSymmetry\nobjRef = {1}\n", m_LibName, objRef);
				str += OutRes();
				str += "\nИмя документа = " + GetDocName();
				string strType;
				strType = string.Format("\nТип объекта/указатель = {0}", m_ObjType);
				str += strType;
				res = Global.Kompas.ksYesNo(str) == 1 ? true : false;
    
				if (doc2D != null && objRef != 0) 
					doc2D.ksLightObj(objRef, 0);
			}
			return res;
		}


		// koSymmetry - Симметрия объекта
		public bool Symmetry(int objRef)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chb2DObjEvents.Checked)
			{
				ksDocument2D doc2D = (ksDocument2D)m_Doc;
				if (doc2D != null && objRef != 0) 
					doc2D.ksLightObj(objRef, 1);

				string str = string.Empty;
				str = string.Format("{0} --> Symmetry\nobjRef = {1}\n", m_LibName, objRef);
				str += OutRes();
				str += "\nИмя документа = " + GetDocName();
				string strType;
				strType = string.Format("\nТип объекта/указатель = {0}", m_ObjType);
				str += strType;
				Global.Kompas.ksMessage(str);

				if (doc2D != null && objRef != 0) 
					doc2D.ksLightObj(objRef, 0);
			}
			return true;
		}


		// koBeginProcess - Начало редактирования\создания объекта
		public bool BeginProcess(int pType, int objRef)
		{
			bool res = false;
			if (m_SelfAdvise && FrmConfig.Instance.chb2DObjEvents.Checked)
			{
				ksDocument2D doc2D = (ksDocument2D)m_Doc;
				if (doc2D != null && objRef != 0) 
					doc2D.ksLightObj(objRef, 1);

				string str = string.Empty;
				str = string.Format("{0} --> BeginProcess\npType = {1}\nobjRef = {2}\n", m_LibName, pType, objRef); 
				str += OutRes();
				str += "\nИмя документа = " + GetDocName();
				string strType;
				strType = string.Format("\nТип объекта/указатель = {0}", m_ObjType);
				str += strType;
				res = Global.Kompas.ksYesNo(str) == 1 ? true : false;

				if (doc2D != null && objRef != 0) 
					doc2D.ksLightObj(objRef, 0);
			}
			return res;
		}
     

		// koEndProcess - Конец редактирования\создания объекта
		public bool EndProcess(int pType)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chb2DObjEvents.Checked)
			{
				ksDocument2D doc2D = (ksDocument2D)m_Doc;
				string str = string.Empty;
				str = string.Format("{0} --> EndProcess\npType = {1}\n", m_LibName, pType); 
				str += OutRes();
				str += "\nИмя документа = " + GetDocName();
				string strType;
				strType = string.Format("\nТип объекта/указатель = {0}", m_ObjType);
				str += strType;
				Global.Kompas.ksMessage(str);
			}
			return true;
		}


		// koCreate - Создание объектов
		public bool CreateObject(int objRef)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chb2DObjEvents.Checked)
			{
				ksDocument2D doc2D = (ksDocument2D)m_Doc;
				if (doc2D != null && objRef != 0) 
					doc2D.ksLightObj(objRef, 1);

				string str = string.Empty;
				str = string.Format("{0} --> CreateObject\nobjRef = {1}\n", m_LibName, objRef); 
				str += OutRes();
				str += "\nИмя документа = " + GetDocName();
				string strType;
				strType = string.Format("\nТип объекта/указатель = {0}", m_ObjType);
				str += strType;
				Global.Kompas.ksMessage(str);

				if (doc2D != null && objRef != 0) 
					doc2D.ksLightObj(objRef, 0);
			}  
			return true;
		}

    
		// koUpdateObject - Редактирование объекта
		public bool UpdateObject(int objRef)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chb2DObjEvents.Checked)
			{
				ksDocument2D doc2D = (ksDocument2D)m_Doc;
				if (doc2D != null && objRef != 0) 
					doc2D.ksLightObj(objRef, 1);

				string str = string.Empty;
				str = string.Format("{0} --> UpdateObject\nobjRef = {1}\n", m_LibName, objRef); 
				str += OutRes();
				str += "\nИмя документа = " + GetDocName();
				string strType;
				strType = string.Format("\nТип объекта/указатель = {0}", m_ObjType);
				str += strType;
				Global.Kompas.ksMessage(str);

				if (doc2D != null && objRef != 0) 
					doc2D.ksLightObj(objRef, 0);
			}
			return true;
		}	
    // koUpdateObject - Редактирование объекта
    public bool BeginDestroyObject(int objRef)
    {
      return true;
    }	
    // koUpdateObject - Редактирование объекта
    public bool DestroyObject(int objRef)
    {
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
  }
}
