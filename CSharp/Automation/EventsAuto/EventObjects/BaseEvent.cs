/////////////////////////////////////////////////////////////////////////////
//
// Базовый клас для обработчиков событий
//
/////////////////////////////////////////////////////////////////////////////

using Kompas6API5;

using System;
using System.Resources;
using System.Diagnostics;
using System.Collections;
using System.Runtime.InteropServices.ComTypes;
using Kompas6Constants;
using KAPITypes;

namespace Steps.NET
{
	public class BaseEvent
	{
		protected int m_Cookie;
		protected object m_Container;
		protected Guid m_Events;
		protected object m_Doc;
		protected int m_ObjType;
		protected ksFeature m_Obj3D;
		protected bool m_SelfAdvise;
		protected IConnectionPoint m_ConnPt;
		protected string m_LibName;
		protected string str = string.Empty;

		public BaseEvent(object obj, Guid events, object doc, int objType,
			ksFeature obj3d, bool selfAdvise)
		{
			m_Cookie = 0;
			m_Container = obj;
			m_Events = events;
			m_Doc = doc;
			m_ObjType = objType;
			m_Obj3D = obj3d;
			m_SelfAdvise = selfAdvise;
			m_ConnPt = null;

			ResourceManager rm = new ResourceManager(typeof (EventsAuto));
			m_LibName = rm.GetString("LibraryName");

			Global.EventList.Add(this);
		}


		~BaseEvent()
		{
			Global.EventList.Remove(this);

			Unadvise();

			m_Container = null;
			m_Doc = null;
			m_Obj3D = null;
		}


		// Получить имя документа
		protected string GetDocName()
		{
			string result = string.Empty;
			if (m_Doc != null)
			{
				ksDocument2D doc2d = m_Doc as ksDocument2D;
				if (doc2d != null)
				{
					ksDocumentParam docPar = Global.Kompas.GetParamStruct((short)StructType2DEnum.ko_DocumentParam) as ksDocumentParam;
					doc2d.ksGetObjParam(doc2d.reference, docPar, ldefin2d.ALLPARAM);
					result = docPar.fileName;
				}
				else
				{
					ksDocument3D doc3d = m_Doc as ksDocument3D;
					if (doc3d != null)
					{
						result = doc3d.fileName;
					}
					else
					{
						ksSpcDocument spcDoc = m_Doc as ksSpcDocument;
						if (spcDoc != null)
						{
							ksDocumentParam docPar = Global.Kompas.GetParamStruct((short)StructType2DEnum.ko_DocumentParam) as ksDocumentParam;
							spcDoc.ksGetObjParam(spcDoc.reference, docPar, ldefin2d.ALLPARAM);
							result = docPar.fileName;
						}
						else
						{
							ksDocumentTxt txtDoc = m_Doc as ksDocumentTxt;
							if (txtDoc != null)
							{
								ksTextDocumentParam docPar = Global.Kompas.GetParamStruct((short)StructType2DEnum.ko_TextDocumentParam) as ksTextDocumentParam;
								txtDoc.ksGetObjParam(spcDoc.reference, txtDoc, ldefin2d.ALLPARAM);
								result = docPar.fileName;
							}
						}
					}
				}
			}
			return result;
		}


		// Подписаться на получение событий
		public int Advise()
		{
			Debug.Assert(m_Cookie == 0, "Повторно подписываться нельзя");

			// Подписаться на получение событий
			if (m_Container != null) 
			{
				IConnectionPointContainer cpContainer = m_Container as IConnectionPointContainer;
				if (cpContainer != null)
				{
					cpContainer.FindConnectionPoint(ref m_Events, out m_ConnPt);
					if (m_ConnPt != null)
						m_ConnPt.Advise(this, out m_Cookie);
				}
			}

			if (m_Cookie == 0)
				return 0;

			if (this.m_SelfAdvise)
			{
				if (this.GetType() == typeof(ApplicationEvent) && FrmConfig.Instance.chbAppEvents.Checked)
				{ 
					str = string.Format("Подписаться на получение событий\n{0} --> ApplicationEvent.Advise", m_LibName);
					Global.Kompas.ksMessage(str);
				}
				else if (this.GetType() == typeof(Document2DEvent) && FrmConfig.Instance.chb2DDocEvents.Checked)
				{
					str = string.Format("Подписаться на получение событий\n{0} --> Document2DEvent.Advise\nИмя документа = {1}", m_LibName, GetDocName());
					Global.Kompas.ksMessage(str);
				}
				else if (this.GetType() == typeof(Document3DEvent) && FrmConfig.Instance.chb3DDocEvents.Checked)
				{
					str = string.Format("Подписаться на получение событий\n{0} --> Document3DEvent.Advise\nИмя документа = {1}", m_LibName, GetDocName());
					Global.Kompas.ksMessage(str);
				}
				else if (this.GetType() == typeof(DocumentEvent) && FrmConfig.Instance.chbDocEvents.Checked)
				{
					str = string.Format("Подписаться на получение событий\n{0} --> DocumentEvent.Advise\nИмя документа = {1}", m_LibName, GetDocName());
					Global.Kompas.ksMessage(str);
				}
				else if (this.GetType() == typeof(Object2DEvent) && FrmConfig.Instance.chb2DObjEvents.Checked)
				{
					ksDocument2D doc2D = (ksDocument2D)m_Doc;
					if (doc2D != null && doc2D.ksExistObj(m_ObjType) != 0) 
						doc2D.ksLightObj(m_ObjType, 1);

					string strType = string.Format("\nТип/указатель объекта = {0}", m_ObjType);
					str = string.Format("Подписаться на получение событий\n{0} --> Object2DEvent.Advise\nИмя документа = {1}{2}", m_LibName, GetDocName(), strType);
					Global.Kompas.ksMessage(str);
      
					if (doc2D != null && doc2D.ksExistObj(m_ObjType) != 0) 
						doc2D.ksLightObj(m_ObjType, 0);
				}
				else if (this.GetType() == typeof(Object3DEvent) && FrmConfig.Instance.chb3DObjEvents.Checked)
				{
					ksDocument3D doc3D = (ksDocument3D)m_Doc;
          ksChooseMng chooseMng = (doc3D.GetChooseMng()) as ksChooseMng;
					if (doc3D != null  && m_Obj3D != null &&
            (chooseMng  != null)    )     
						chooseMng.Choose(m_Obj3D);

					string strType = string.Format("\nТип объекта = {0}", m_ObjType);
					string strObj3DName = string.Empty;
					if (m_Obj3D != null)
						strObj3DName = m_Obj3D.name;
					str = string.Format("Подписаться на получение событий\n{0} --> Object3DEvent.Advise\nИмя документа = {1}{2}\nИмя 3D объекта = {3}", m_LibName, GetDocName(), strType, strObj3DName);
					Global.Kompas.ksMessage(str);

					if (chooseMng != null)
						chooseMng.UnChoose(m_Obj3D); 
				}
				else if (this.GetType() == typeof(SelectMngEvent) && FrmConfig.Instance.chbSelectEvents.Checked)
				{
					str = string.Format("Подписаться на получение событий\n{0} --> SelectMngEvent.Advise\nИмя документа = {1}", m_LibName, GetDocName());
					Global.Kompas.ksMessage(str);
				}
				else if (this.GetType() == typeof(SpcDocumentEvent) && FrmConfig.Instance.chbSpecDocEvents.Checked)
				{
					str = string.Format("Подписаться на получение событий\n{0} --> SpcDocumentEvent.Advise\nИмя документа = {1}", m_LibName, GetDocName());
					Global.Kompas.ksMessage(str);
				}
				else if (this.GetType() == typeof(SpcObjectEvent) && FrmConfig.Instance.chbSpecObjEvents.Checked)
				{
					str = string.Format("Подписаться на получение событий\n{0} --> SpcObjectEvent.Advise\nИмя документа = {1}", m_LibName, GetDocName());
					Global.Kompas.ksMessage(str);
				}
				else if (this.GetType() == typeof(SpecificationEvent) && FrmConfig.Instance.chbSpecEvents.Checked)
				{
					str = string.Format("Подписаться на получение событий\n{0} --> SpecificationEvent.Advise\nИмя документа = {1}", m_LibName, GetDocName());
					Global.Kompas.ksMessage(str);
				}
				else if (this.GetType() == typeof(StampEvent) && FrmConfig.Instance.chbStampEvents.Checked)
				{
					str = string.Format("Подписаться на получение событий\n{0} --> StampEvent.Advise\nИмя документа = {1}", m_LibName, GetDocName());
					Global.Kompas.ksMessage(str);
				}
			} 

			return m_Cookie;
		}


		// Отписаться от получения событий
		void Unadvise()
		{
      if (m_Cookie == 0)
        return;

			if (m_ConnPt != null)				// Подписка была
			{
				m_ConnPt.Unadvise(m_Cookie);	// Отписаться от получения событий
				m_ConnPt = null;
			} 
			m_Cookie = 0;

      if (m_SelfAdvise && Global.Kompas != null )
			{
				if (this.GetType() == typeof(ApplicationEvent) && FrmConfig.Instance.chbAppEvents.Checked)
				{ 
					str = string.Format("Отписаться от получения событий\n{0} --> ApplicationEvent.Unadvise", m_LibName);
					Global.Kompas.ksMessage(str);
				}
				else if (this.GetType() == typeof(Document2DEvent) && FrmConfig.Instance.chb2DDocEvents.Checked)
				{
					str = string.Format("Отписаться от получения событий\n{0} --> Document2DEvent.Unadvise\nИмя документа = {1}", m_LibName, GetDocName());
					Global.Kompas.ksMessage(str);
				}
				else if (this.GetType() == typeof(Document3DEvent) && FrmConfig.Instance.chb3DDocEvents.Checked)
				{
					str = string.Format("Отписаться от получения событий\n{0} --> Document3DEvent.Unadvise\nИмя документа = {1}", m_LibName, GetDocName());
					Global.Kompas.ksMessage(str);
				}
				else if (this.GetType() == typeof(DocumentEvent) && FrmConfig.Instance.chbDocEvents.Checked)
				{
					str = string.Format("Отписаться от получения событий\n{0} --> DocumentEvent.Unadvise\nИмя документа = {1}", m_LibName, GetDocName());
					Global.Kompas.ksMessage(str);
				}
				else if (this.GetType() == typeof(Object2DEvent) && FrmConfig.Instance.chb2DObjEvents.Checked)
				{
					ksDocument2D doc2D = (ksDocument2D)m_Doc;
					if (doc2D != null  && doc2D.ksExistObj(m_ObjType) != 0) 
						doc2D.ksLightObj(m_ObjType, 1);

					string strType;
					strType = string.Format("\nТип объекта/указатель = {0}", m_ObjType);
					str = string.Format("Отписаться от получения событий\n{0} --> Object2DEvent.Unadvise\nИмя документа = {1}", m_LibName, GetDocName(), strType);
					Global.Kompas.ksMessage(str);

					if (doc2D != null && doc2D.ksExistObj(m_ObjType) != 0) 
						doc2D.ksLightObj(m_ObjType, 0);
				}
				else if (this.GetType() == typeof(Object3DEvent) && FrmConfig.Instance.chb3DObjEvents.Checked)
				{
					ksDocument3D doc3D = (ksDocument3D)m_Doc;
					ksChooseMng chooseMng = null;
					if (doc3D != null  && m_Obj3D != null && 
						(chooseMng = (ksChooseMng)doc3D.GetChooseMng()) != null)         
						chooseMng.Choose(m_Obj3D);

					string strType;
					strType = string.Format("\nТип объекта = {0}", m_ObjType);
					string strObj3DName = string.Empty;
					if (m_Obj3D != null)
						strObj3DName = m_Obj3D.name;
					str = string.Format("Отписаться от получения событий\n{0} --> Object3DEvent.Unadvise\nИмя документа = {1}{2}\nИмя 3D объекта = {3}", m_LibName, GetDocName(), strType, strObj3DName);
					Global.Kompas.ksMessage(str);

					if (chooseMng != null)
						chooseMng.UnChoose(m_Obj3D);
				}
				else if (this.GetType() == typeof(SelectMngEvent) && FrmConfig.Instance.chbSelectEvents.Checked)
				{
					str = string.Format("Отписаться от получения событий\n{0} --> SelectMngEvent.Unadvise\nИмя документа = {1}", m_LibName, GetDocName());
					Global.Kompas.ksMessage(str);
				}
				else if (this.GetType() == typeof(SpcDocumentEvent) && FrmConfig.Instance.chbSpecDocEvents.Checked)
				{
					str = string.Format("Отписаться от получения событий\n{0} --> SpcDocumentEvent.Unadvise\nИмя документа = {1}", m_LibName, GetDocName());
					Global.Kompas.ksMessage(str);
				}
				else if (this.GetType() == typeof(SpcObjectEvent) && FrmConfig.Instance.chbSpecObjEvents.Checked)
				{
					str = string.Format("Отписаться от получения событий\n{0} --> SpcObjectEvent.Unadvise\nИмя документа = {1}", m_LibName, GetDocName());
					Global.Kompas.ksMessage(str);
				}
				else if (this.GetType() == typeof(SpecificationEvent) && FrmConfig.Instance.chbSpecEvents.Checked)
				{
					str = string.Format("Отписаться от получения событий\n{0} --> SpecificationEvent.Unadvise\nИмя документа = {1}", m_LibName, GetDocName());
					Global.Kompas.ksMessage(str);
				}
				else if (this.GetType() == typeof(StampEvent) && FrmConfig.Instance.chbStampEvents.Checked)
				{
					str = string.Format("Отписаться от получения событий\n{0} --> StampEvent.Unadvise\nИмя документа = {1}", m_LibName, GetDocName());
					Global.Kompas.ksMessage(str);
				}
			} 
		}


		// Отписать все события
		public static void TerminateEvents()
		{
			int count = Global.EventList.Count;
			for (int i = 0; i < count ; i ++)
			{
				BaseEvent headEvent = (BaseEvent)Global.EventList[0];
				headEvent.Disconnect();
				Global.EventList.Remove(headEvent);
			}
		}


		// Отписать все события по GUID и документу
		public static void TerminateEvents(Type type, object doc, int objType, ksFeature obj3D)
		{
			int count = Global.EventList.Count;
			for (int i = count - 1; i >= 0; i --)
			{
				object obj = Global.EventList[i];
				BaseEvent evt = (BaseEvent)obj;  

				if (evt != null &&
					(evt.GetType() == type || type == null) &&
					(doc == null || evt.m_Doc == doc) &&
					(obj3D == null || evt.m_Obj3D == obj3D) &&
					(objType == -1 || evt.m_ObjType == objType))
				{
					evt.Disconnect();	// В деструкторе будет удален из списка RemoveAt(pos)
					Global.EventList.Remove(evt);
				}
			}
		}


		// Освободить ссылки
		void Clear()
		{
			if (m_Container != null) 
			{
				m_Container = null;     
			}

			if (m_Doc != null) 
			{
				m_Doc = null;
			}

			if (m_Obj3D != null)
			{
				m_Obj3D = null;
			}

			m_Events = Guid.Empty;
		}


		// Отсоединиться
		void Disconnect()
		{
			Unadvise();
			Clear();
		}


		public static bool FindEvent(Type type, object doc, int objType, ksFeature obj3D)
		{
			int count = Global.EventList.Count;
			for (int i = count - 1; i >= 0; i--)
			{
				object obj = Global.EventList[i];
				BaseEvent evt = (BaseEvent)obj;  
				if (evt != null &&
					evt.GetType() == type &&
					(doc == null || evt.m_Doc == doc) &&
					((obj3D == null && evt.m_Obj3D == null) || (evt.m_Obj3D == obj3D)) &&
					((objType == -1 && evt.m_ObjType == 0) || (evt.m_ObjType == objType)))
					return true;
			}
			return false;
		}

		public static void ListEvents()
		{
			string str = string.Format("Подписанные события:");

			int count = Global.EventList.Count;
			for (int i = count - 1; i >= 0; i --)
			{
				object obj = Global.EventList[i];
				BaseEvent evt = (BaseEvent)obj;     
    
				if (evt.GetType() == typeof(ApplicationEvent))
				{ 
					str += "\nApplicationEvent, подписка пользователем: ";
					str += evt.m_SelfAdvise ? "да" : "нет";
				}
				else if (evt.GetType() == typeof(Document2DEvent))
				{
					str += "\nDocument2DEvent, подписка пользователем: ";
					str += evt.m_SelfAdvise ? "да," : "нет,";
					str += " имя документа = " + evt.GetDocName();
				}
				else if (evt.GetType() == typeof(Document3DEvent))
				{
					str += "\nDocument3DEvent, подписка пользователем: ";
					str += evt.m_SelfAdvise ? "да," : "нет,";
					str += " имя документа = " + evt.GetDocName(); 
				}
				else if (evt.GetType() == typeof(DocumentEvent))
				{
					str += "\nDocumentEvent, подписка пользователем: ";
					str += evt.m_SelfAdvise ? "да," : "нет,";
					str += " имя документа = " + evt.GetDocName();
				}
				else if (evt.GetType() == typeof(Object2DEvent))
				{
					str += "\nObject2DEvent, подписка пользователем: ";
					str += evt.m_SelfAdvise ? "да," : "нет,";
					str += " имя документа = " + evt.GetDocName();
				}
				else if (evt.GetType() == typeof(Object3DEvent))
				{
					str += "\nObject3DEvent, подписка пользователем: ";
					str += evt.m_SelfAdvise ? "да," : "нет,";
					str += " имя документа = " + evt.GetDocName();       
				}
				else if (evt.GetType() == typeof(SelectMngEvent))
				{
					str += "\nSelectMngEvent, подписка пользователем: ";
					str += evt.m_SelfAdvise ? "да," : "нет,";
					str += " имя документа = " + evt.GetDocName();
				}
				else if (evt.GetType() == typeof(SpcDocumentEvent))
				{
					str += "\nSpcDocumentEvent, подписка пользователем: ";
					str += evt.m_SelfAdvise ? "да," : "нет,";
					str += " имя документа = " + evt.GetDocName();
				}
				else if (evt.GetType() == typeof(SpcObjectEvent))
				{
					str += "\nSpcObjectEvent, подписка пользователем: ";
					str += evt.m_SelfAdvise ? "да," : "нет,";
					str += " имя документа = " + evt.GetDocName();  
				}
				else if (evt.GetType() == typeof(SpecificationEvent))
				{
					str += "\nSpecificationEvent, подписка пользователем: ";
					str += evt.m_SelfAdvise ? "да," : "нет,";
					str += " имя документа = " + evt.GetDocName();
				}
				else if (evt.GetType() == typeof(StampEvent))
				{
					str += "\nStampEvent, подписка пользователем: ";
					str += evt.m_SelfAdvise ? "да," : "нет,";
					str += " имя документа = " + evt.GetDocName();
				}
			}
			Global.Kompas.ksMessage(str);
		}
	}
}
