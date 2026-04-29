////////////////////////////////////////////////////////////////////////////////
//
// Document3DEvent - обработчик событий от 3D документа
//
////////////////////////////////////////////////////////////////////////////////

using Kompas6API5;

using System;
using Kompas6Constants;
using Kompas6Constants3D;
using KAPITypes;
using System.Runtime.InteropServices;

namespace Steps.NET
{
	public class Document3DEvent : BaseEvent, ksDocument3DNotify
	{
		public Document3DEvent(object obj, object doc)
			: base(obj, typeof(ksDocument3DNotify).GUID, doc) {}

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


		// d3BeginChoiceMarking - Начало выбора обозначения
		public bool BeginChoiceMarking()
		{
			return true;
		}


		// d3СhoiceMarking - Закончен выбор обозначения
		public bool ChoiceMarking(string marking)
		{
			return true;
		}


		// d3BeginSetPartFromFile - Начало установки компонента в сборку (до диалога выбора имени)
		public bool BeginSetPartFromFile()
		{
			return true;
		}


		// d3BeginCreatePartFromFile - Начало создания компонента в сборке (до диалога выбора имени)
		public bool BeginCreatePartFromFile(bool part, entity plane)
		{
			return true;
		}

    public bool CreateEmbodiment(string marking)
    {
      return true;
    }

    public bool DeleteEmbodiment(string marking)
    {
      return true;
    }
    public bool ChangeCurrentEmbodiment(string marking)
    {
      return true;
    }
    public bool BeginChoiceProperty(object obj, double propID)
    {
      return true;
    }

    public bool ChoiceProperty(object obj, double propID)
    {
      return true;
    }

    public bool BeginRollbackFeatures()
    {
      return true;
    }
    public bool RollbackFeatures()
    {
      return true;
    }
    public bool BedinLoadCombinationChange(int index)
    {
      return true;
    }
    public bool LoadCombinationChange(int index)
    {
      return true;
    }
    public bool BeginDeleteMaterial()
    {
      return true;
    }
    public bool DeleteMaterial()
    {
      return true;
    }
    public bool BeginDeleteProperty([MarshalAs(UnmanagedType.IDispatch)] object obj, double propID)
    {
      return true;
    }
    public bool DeleteProperty([MarshalAs(UnmanagedType.IDispatch)] object obj, double propID)
    {
      return true;
    }
  }

}