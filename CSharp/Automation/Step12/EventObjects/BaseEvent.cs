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

using reference = System.Int32;

namespace Steps.NET
{
	public class BaseEvent
	{
		protected int m_Cookie;
		protected object m_Container;
		protected Guid m_Events;
		protected object m_Doc;
		protected IConnectionPoint m_ConnPt;
		protected reference m_Reference;
		protected object m_3DObj;
		protected int m_ObjType;

		public BaseEvent(object obj, Guid events)
		{
			m_Cookie = 0;
			m_Container = obj;
			m_Events = events;
			m_Doc = null;
			m_ConnPt = null;
			m_3DObj = null;
			m_Reference = 0;
			m_ObjType = 0;

			Global.EventList.Add(this);
		}

		public BaseEvent(object obj, Guid events, object doc)
		{
			m_Cookie = 0;
			m_Container = obj;
			m_Events = events;
			m_Doc = doc;
			m_ConnPt = null;
			m_3DObj = null;
			m_Reference = 0;
			m_ObjType = 0;

			Global.EventList.Add(this);
		}

		public BaseEvent(object obj, Guid events, object doc, int objType, reference objRef)
		{
			m_Cookie = 0;
			m_Container = obj;
			m_Events = events;
			m_Doc = doc;
			m_ConnPt = null;
			m_3DObj = null;
			m_Reference = objRef;
			m_ObjType = objType;

			Global.EventList.Add(this);
		}

		public BaseEvent(object obj, Guid events, object doc, int objType, object obj3d)
		{
			m_Cookie = 0;
			m_Container = obj;
			m_Events = events;
			m_Doc = doc;
			m_ConnPt = null;
			m_3DObj = obj3d;
			m_Reference = 0;
			m_ObjType = objType;

			Global.EventList.Add(this);
		}


		~BaseEvent()
		{
			Global.EventList.Remove(this);

			Unadvise();

			m_Container = null;
			m_Doc = null;
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

			return m_Cookie;
		}


		// Отписаться от получения событий
		void Unadvise()
		{
			if (m_ConnPt != null)				// Подписка была
			{
				m_ConnPt.Unadvise(m_Cookie);	// Отписаться от получения событий
				m_ConnPt = null;
			} 
			m_Cookie = 0;
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
		public static void TerminateEvents(Type type, object doc)
		{
			int count = Global.EventList.Count;
			for (int i = count - 1; i >= 0; i --)
			{
				object obj = Global.EventList[i];
				BaseEvent evt = (BaseEvent)obj;  

				if (evt != null &&
					(evt.GetType() == type || type == null) &&
					(doc == null || evt.m_Doc == doc))
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

			m_Events = Guid.Empty;
		}


		// Отсоединиться
		void Disconnect()
		{
			Unadvise();
			Clear();
		}


		public static bool FindEvent(Type type, object doc)
		{
			int count = Global.EventList.Count;
			for (int i = count - 1; i >= 0; i--)
			{
				object obj = Global.EventList[i];
				BaseEvent evt = (BaseEvent)obj;  
				if (evt != null &&
					evt.GetType() == type &&
					(doc == null || evt.m_Doc == doc))
					return true;
			}
			return false;
		}

		public static bool FindEvent(Type type, object doc, int objType)
		{
			int count = Global.EventList.Count;
			for (int i = count - 1; i >= 0; i--)
			{
				object obj = Global.EventList[i];
				BaseEvent evt = (BaseEvent)obj;  
				if (evt != null &&
					evt.GetType() == type &&
					(doc == null || evt.m_Doc == doc) &&
					(objType == -1) || evt.m_ObjType == objType)
					return true;
			}
			return false;
		}
		public static bool FindEvent(Type type, object doc, int objType, object obj3d)
		{
			int count = Global.EventList.Count;
			for (int i = count - 1; i >= 0; i--)
			{
				object obj = Global.EventList[i];
				BaseEvent evt = (BaseEvent)obj;  
				if (evt != null &&
					evt.GetType() == type &&
					(doc == null || evt.m_Doc == doc) &&
					(objType == -1 || evt.m_ObjType == objType) &&
					(obj3d == null || evt.m_3DObj == obj3d))
					return true;
			}
			return false;
		}
	}
}
