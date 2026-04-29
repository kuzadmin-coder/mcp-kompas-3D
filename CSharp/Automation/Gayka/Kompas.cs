
using Kompas6API5;
using KompasAPI7;

using System;
using Kompas6Constants;

namespace Steps.NET
{
	public class Kompas
	{
		private Kompas()
		{
		}

		private static Kompas instance;
		public static Kompas Instance
		{
			get
			{
				if (instance == null)
					instance = new Kompas();
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

		private _Application kompasApp;
		public _Application KompasApp
		{
			get
			{
				return kompasApp;
			}
			set
			{
				kompasApp = value;
			}
		}
		private ksMathematic2D math;
		public ksMathematic2D Math
		{
			get
			{
				if (math == null)
					math = (ksMathematic2D)Kompas.Instance.KompasObject.GetMathematic2D();
				return math;
			}
		}
	}
}
