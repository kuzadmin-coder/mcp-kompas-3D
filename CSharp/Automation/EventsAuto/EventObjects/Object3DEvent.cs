////////////////////////////////////////////////////////////////////////////////
//
// Object3DEvent - обработчик событий объектов 3D документа
//
////////////////////////////////////////////////////////////////////////////////

using Kompas6API5;
using System;
using Kompas6Constants;
using KAPITypes;

namespace Steps.NET
{
	public class Object3DEvent : BaseEvent, ksObject3DNotify
	{
		private ksObject3DNotifyResult m_res = null;

		public Object3DEvent(object obj, object doc, int objType, ksFeature obj3D,
				ksObject3DNotifyResult res, bool selfAdvise)
			: base(obj, typeof(ksObject3DNotify).GUID, doc,
			objType, obj3D, selfAdvise)
		{ }

		string OutRes()
		{
			string str = string.Empty;
			if (m_res != null)
			{
				int notifyType = m_res.GetNotifyType();
				ksFeatureCollection featureCollection = m_res.GetFeatureCollection();
				ksPlacement placement = m_res.GetPlacement();
				double x = -1, y = -1, z = -1;
				if (placement != null)
					placement.GetOrigin(out x, out y, out z);
				str = string.Format("Object3DNotifyResult: GetNotifyType = {0},\nGetFeatureCollection->GetCount() = {1}\nGetPlacement->GetOrigin({2}, {3}, {4})",
					notifyType,
					featureCollection != null ? featureCollection.GetCount() : 0, x, y, z);
			}
			return str;
		}


		// o3BeginDelete - Начало удаления объектов
		public bool BeginDelete(object obj)
		{
			bool res = false;
			if (m_SelfAdvise && FrmConfig.Instance.chb3DObjEvents.Checked)
			{
				ksDocument3D doc3D = (ksDocument3D)m_Doc;
				ksChooseMng chooseMng = null;
				if (doc3D != null && m_Obj3D != null &&
					(chooseMng = (ksChooseMng)doc3D.GetChooseMng()) != null)
					chooseMng.Choose(obj);

				string str = string.Empty;
				str = string.Format("{0} --> BeginDelete\nobj = {1}\n", m_LibName, obj);
				str += OutRes();
				str += "\nИмя документа = " + GetDocName();
				string strType;
				strType = string.Format("\nТип объекта = {0}", m_ObjType);
				str += strType;
				string strObj3DName = string.Empty;
				if (m_Obj3D != null)
					strObj3DName = m_Obj3D.name;
				str += "\nИмя объекта = ";
				str += strObj3DName;
				res = Global.Kompas.ksYesNo(str) == 1 ? true : false;

				if (chooseMng != null)
					chooseMng.UnChoose(obj);
			}
			return res;
		}


		// o3Delete - Oбъекты удалены
		public bool Delete(object obj)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chb3DObjEvents.Checked)
			{
				ksDocument3D doc3D = (ksDocument3D)m_Doc;
				ksChooseMng chooseMng = null;
				if (doc3D != null && m_Obj3D != null &&
					(chooseMng = (ksChooseMng)doc3D.GetChooseMng()) != null)
					chooseMng.Choose(obj);

				string str = string.Empty;
				str = string.Format("{0} --> Delete\nobj = {1}\n", m_LibName, obj);
				str += OutRes();
				str += "\nИмя документа = " + GetDocName();
				string strType;
				strType = string.Format("\nТип объекта = {0}", m_ObjType);
				str += strType;
				string strObj3DName = string.Empty;
				if (m_Obj3D != null)
					strObj3DName = m_Obj3D.name;
				str += "\nИмя объекта = ";
				str += strObj3DName;
				Global.Kompas.ksMessage(str);

				if (chooseMng != null)
					chooseMng.UnChoose(obj);
			}

			// Если удаляется объект удаляем и подписку на его события
			if (obj != null && m_Obj3D == obj)
				this.Delete(obj);
			return true;
		}


		// o3Excluded - Oбъект исключен/включен в расчет
		public bool excluded(object obj, bool excluded)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chb3DObjEvents.Checked)
			{
				ksDocument3D doc3D = (ksDocument3D)m_Doc;
				ksChooseMng chooseMng = null;
				if (doc3D != null && m_Obj3D != null &&
					(chooseMng = (ksChooseMng)doc3D.GetChooseMng()) != null)
					chooseMng.Choose(obj);

				string str = string.Empty;
				str = string.Format("{0} --> Excluded\nobj = {1}\nexcluded = {2}\n", m_LibName, obj, excluded);
				str += OutRes();
				str += "\nИмя документа = " + GetDocName();
				string strType;
				strType = string.Format("\nТип объекта = {0}", m_ObjType);
				str += strType;
				string strObj3DName = string.Empty;
				if (m_Obj3D != null)
					strObj3DName = m_Obj3D.name;
				str += "\nИмя объекта = ";
				str += strObj3DName;
				Global.Kompas.ksMessage(str);

				if (chooseMng != null)
					chooseMng.UnChoose(obj);
			}
			return true;
		}


		// o3Hidden - Oбъект скрыт/показан
		public bool hidden(object obj, bool _hidden)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chb3DObjEvents.Checked)
			{
				ksDocument3D doc3D = (ksDocument3D)m_Doc;
				ksChooseMng chooseMng = null;
				if (doc3D != null && m_Obj3D != null &&
					(chooseMng = (ksChooseMng)doc3D.GetChooseMng()) != null)
					chooseMng.Choose(obj);

				string str = string.Empty;
				str = string.Format("{0} --> Hidden\nobj = {1}\n_hidden = {2}\n", m_LibName, obj, _hidden);
				str += OutRes();
				str += "\nИмя документа = " + GetDocName();
				string strType;
				strType = string.Format("\nТип объекта = {0}", m_ObjType);
				str += strType;
				string strObj3DName = string.Empty;
				if (m_Obj3D != null)
					strObj3DName = m_Obj3D.name;
				str += "\nИмя объекта = ";
				str += strObj3DName;
				Global.Kompas.ksMessage(str);

				if (chooseMng != null)
					chooseMng.UnChoose(obj);
			}
			return true;
		}


		// o3BeginPropertyChanged - Начало изменения свойств объета
		public bool BeginPropertyChanged(object obj)
		{
			bool res = false;
			if (m_SelfAdvise && FrmConfig.Instance.chb3DObjEvents.Checked)
			{
				ksDocument3D doc3D = (ksDocument3D)m_Doc;
				ksChooseMng chooseMng = null;
				if (doc3D != null && m_Obj3D != null &&
					(chooseMng = (ksChooseMng)doc3D.GetChooseMng()) != null)
					chooseMng.Choose(obj);

				string str = string.Empty;
				str = string.Format("{0} --> BeginPropertyChanged\nobj = {1}\n", m_LibName, obj);
				str += OutRes();
				str += "\nИмя документа = " + GetDocName();
				string strType;
				strType = string.Format("\nТип объекта = {0}", m_ObjType);
				str += strType;
				string strObj3DName = string.Empty;
				if (m_Obj3D != null)
					strObj3DName = m_Obj3D.name;
				str += "\nИмя объекта = ";
				str += strObj3DName;
				res = Global.Kompas.ksYesNo(str) == 1 ? true : false;

				if (chooseMng != null)
					chooseMng.UnChoose(obj);
			}
			return res;
		}


		// o3PropertyChanged - Изменены свойства объета
		public bool PropertyChanged(object obj)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chb3DObjEvents.Checked)
			{
				ksDocument3D doc3D = (ksDocument3D)m_Doc;
				ksChooseMng chooseMng = null;
				if (doc3D != null && m_Obj3D != null &&
					(chooseMng = (ksChooseMng)doc3D.GetChooseMng()) != null)
					chooseMng.Choose(obj);

				string str = string.Empty;
				str = string.Format("{0} --> PropertyChanged\nobj = {1}\n", m_LibName, obj);
				str += OutRes();
				str += "\nИмя документа = " + GetDocName();
				string strType;
				strType = string.Format("\nТип объекта = {0}", m_ObjType);
				str += strType;
				string strObj3DName = string.Empty;
				if (m_Obj3D != null)
					strObj3DName = m_Obj3D.name;
				str += "\nИмя объекта = ";
				str += strObj3DName;
				Global.Kompas.ksMessage(str);

				if (chooseMng != null)
					chooseMng.UnChoose(obj);
			}
			return true;
		}


		// o3BeginPlacementChanged - Начало изменения положения объета
		public bool BeginPlacementChanged(object obj)
		{
			bool res = false;
			if (m_SelfAdvise && FrmConfig.Instance.chb3DObjEvents.Checked)
			{
				ksDocument3D doc3D = (ksDocument3D)m_Doc;
				ksChooseMng chooseMng = null;
				if (doc3D != null && m_Obj3D != null &&
					(chooseMng = (ksChooseMng)doc3D.GetChooseMng()) != null)
					chooseMng.Choose(obj);

				string str = string.Empty;
				str = string.Format("{0} --> BeginPlacementChanged\nobj = {1}\n", m_LibName, obj);
				str += OutRes();
				str += "\nИмя документа = " + GetDocName();
				string strType;
				strType = string.Format("\nТип объекта = {0}", m_ObjType);
				str += strType;
				string strObj3DName = string.Empty;
				if (m_Obj3D != null)
					strObj3DName = m_Obj3D.name;
				str += "\nИмя объекта = ";
				str += strObj3DName;
				res = Global.Kompas.ksYesNo(str) == 1 ? true : false;

				if (chooseMng != null)
					chooseMng.UnChoose(obj);
			}
			return res;
		}


		// o3PlacementChanged - Изменено положения объета
		public bool PlacementChanged(object obj)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chb3DObjEvents.Checked)
			{
				ksDocument3D doc3D = (ksDocument3D)m_Doc;
				ksChooseMng chooseMng = null;
				if (doc3D != null && m_Obj3D != null &&
					(chooseMng = (ksChooseMng)doc3D.GetChooseMng()) != null)
					chooseMng.Choose(obj);

				string str = string.Empty;
				str = string.Format("{0} --> PlacementChanged\nobj = {1}\n", m_LibName, obj);
				str += OutRes();
				str += "\nИмя документа = " + GetDocName();
				string strType;
				strType = string.Format("\nТип объекта = {0}", m_ObjType);
				str += strType;
				string strObj3DName = string.Empty;
				if (m_Obj3D != null)
					strObj3DName = m_Obj3D.name;
				str += "\nИмя объекта = ";
				str += strObj3DName;
				Global.Kompas.ksMessage(str);

				if (chooseMng != null)
					chooseMng.UnChoose(obj);
			}
			return true;
		}


		// o3BeginProcess - Начало редактирования\создания объекта
		public bool BeginProcess(int pType, object obj)
		{
			bool res = false;
			if (m_SelfAdvise && FrmConfig.Instance.chb3DObjEvents.Checked)
			{
				ksDocument3D doc3D = (ksDocument3D)m_Doc;
				ksChooseMng chooseMng = null;
				if (doc3D != null && m_Obj3D != null &&
					(chooseMng = (ksChooseMng)doc3D.GetChooseMng()) != null)
					chooseMng.Choose(obj);

				string str = string.Empty;
				str = string.Format("{0} --> BeginProcess\npType = {1}\nobj = {2}\n", m_LibName, pType, obj);
				str += OutRes();
				str += "\nИмя документа = " + GetDocName();
				string strType;
				strType = string.Format("\nТип объекта = {0}", m_ObjType);
				str += strType;
				string strObj3DName = string.Empty;
				if (m_Obj3D != null)
					strObj3DName = m_Obj3D.name;
				str += "\nИмя объекта = ";
				str += strObj3DName;
				res = Global.Kompas.ksYesNo(str) == 1 ? true : false;

				if (chooseMng != null)
					chooseMng.UnChoose(obj);
			}
			return res;
		}


		// o3EndProcess - Конец редактирования\создания объекта
		public bool EndProcess(int pType)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chb3DObjEvents.Checked)
			{
				ksDocument3D doc3D = (ksDocument3D)m_Doc;

				string str = string.Empty;
				str = string.Format("{0} --> EndProcess\npType = {1}\n", m_LibName, pType);
				str += OutRes();
				str += "\nИмя документа = " + GetDocName();
				string strType;
				strType = string.Format("\nТип объекта = {0}", m_ObjType);
				str += strType;
				string strObj3DName = string.Empty;
				if (m_Obj3D != null)
					strObj3DName = m_Obj3D.name;
				str += "\nИмя объекта = ";
				str += strObj3DName;
				Global.Kompas.ksMessage(str);
			}
			return true;
		}


		// o3CreateObject - Создание объекта
		public bool CreateObject(object obj)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chb3DObjEvents.Checked)
			{
				ksDocument3D doc3D = (ksDocument3D)m_Doc;
				ksChooseMng chooseMng = null;
				if (doc3D != null && m_Obj3D != null &&
					(chooseMng = (ksChooseMng)doc3D.GetChooseMng()) != null)
					chooseMng.Choose(obj);

				string str = string.Empty;
				str = string.Format("{0} --> CreateObject\nobj = {1}\n", m_LibName, obj);
				str += OutRes();
				str += "\nИмя документа = " + GetDocName();
				string strType;
				strType = string.Format("\nТип объекта = {0}", m_ObjType);
				str += strType;
				string strObj3DName = string.Empty;
				if (m_Obj3D != null)
					strObj3DName = m_Obj3D.name;
				str += "\nИмя объекта = ";
				str += strObj3DName;
				Global.Kompas.ksMessage(str);

				if (chooseMng != null)
					chooseMng.UnChoose(obj);
			}
			return true;
		}


		// o3UpdateObject - Редактирование объекта
		public bool UpdateObject(object obj)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chb3DObjEvents.Checked)
			{
				ksDocument3D doc3D = (ksDocument3D)m_Doc;
				ksChooseMng chooseMng = null;
				if (doc3D != null && m_Obj3D != null &&
					(chooseMng = (ksChooseMng)doc3D.GetChooseMng()) != null)
					chooseMng.Choose(obj);

				string str = string.Empty;
				str = string.Format("{0} --> UpdateObject\nobj = {1}\n", m_LibName, obj);
				str += OutRes();
				str += "\nИмя документа = " + GetDocName();
				string strType;
				strType = string.Format("\nТип объекта = {0}", m_ObjType);
				str += strType;
				string strObj3DName = string.Empty;
				if (m_Obj3D != null)
					strObj3DName = m_Obj3D.name;
				str += "\nИмя объекта = ";
				str += strObj3DName;
				Global.Kompas.ksMessage(str);

				if (chooseMng != null)
					chooseMng.UnChoose(obj);
			}
			return true;
		}

		public bool BeginLoadStateChange(object obj, int loadState)
		{
			return true;
		}
		public bool LoadStateChange(object obj, int loadState)
    {
      return true;
    }
  }
}
