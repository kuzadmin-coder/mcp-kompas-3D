using Kompas6API5;

using System;
using Kompas6Constants;

namespace Steps.NET
{
	public class Global
	{
		private Global()
		{
		}


		private static Global instance;
		public static Global Instance
		{
			get
			{
				if (instance == null)
					instance = new Global();
				return instance;
			}
		}


		private KompasObject kompasObject;
		public KompasObject KompasObject
		{
			get
			{
				return kompasObject;
			}
			set
			{
				kompasObject = value;
			}
		}


		private ksDocument2D document2d;
		public ksDocument2D Document2D
		{
			get
			{
				return document2d;
			}
			set
			{
				document2d = value;
			}
		}

	}
}
