//#define __LIGHT_VERSION__
using Kompas6API5;
using System;
using System.Collections;

namespace Steps.NET
{
	public class Global
	{
		private static ArrayList eventList = new ArrayList();
		public static ArrayList EventList
		{
			get
			{
				return eventList;
			}
			set
			{
				eventList = value;
			}
		}

		private static KompasObject kompas;
		public static KompasObject Kompas
		{
			get
			{
				return kompas;
			}
			set
			{
				kompas = value;
			}
		}
	}
}
