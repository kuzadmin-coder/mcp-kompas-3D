////////////////////////////////////////////////////////////////////////////////
//
// Document2DEvent - обработчик событий от 2D документа
//
////////////////////////////////////////////////////////////////////////////////

using Kompas6API5;

using System;
using Kompas6Constants;
using KAPITypes;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Steps.NET
{
	public class Document2DEvent : BaseEvent, ksDocument2DNotify
	{
		public Document2DEvent(object obj, object doc)
			: base(obj, typeof(ksDocument2DNotify).GUID, doc) {}

		// d3BeginRebuild - Начало перестроения модели
		public bool BeginRebuild()
		{ 
			return true;
		}


		// d3Rebuild - Модель перестроена
		public bool Rebuild()
		{   
			Global.UpdateSlideBox(null);
			return true;
		}


		// d3BeginChoiceMaterial - Начало выбора материала
		public bool BeginChoiceMaterial()
		{ 
			return true;
		}


		// d3СhoiceMaterial - Закончен выбор материала
		public bool ChoiceMaterial(string material, double density)
		{
			return true;
		}


		// d2BeginInsertFragment - Начало вставки фрагмента (до диалога выбора имени)
		public bool BeginInsertFragment()
		{
			return true;
		}


		// d2LocalFragmentEdit
		public bool LocalFragmentEdit(object newDoc, bool newFrw)
		{
			return true;
		}

    public bool BeginChoiceProperty(int objRef, double propID)
    {
      return true;
    }

    public bool ChoiceProperty(int objRef, double propID)
    {
      return true;
    }

    public bool BeginDeleteProperty(int objRef, double propID)
    {
      return true;
    }

    public bool DeleteProperty(int objRef, double propID)
    {
      return true;
    }


  }
}