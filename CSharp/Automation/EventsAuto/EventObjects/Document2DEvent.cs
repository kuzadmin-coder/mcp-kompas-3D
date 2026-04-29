////////////////////////////////////////////////////////////////////////////////
//
// Document2DEvent - обработчик событий от 2D документа
//
////////////////////////////////////////////////////////////////////////////////

using Kompas6API5;
using System;
using Kompas6Constants;
using KAPITypes;

namespace Steps.NET
{
	public class Document2DEvent : BaseEvent, ksDocument2DNotify
	{
		public Document2DEvent(object obj, object doc, bool selfAdvise)
			: base(obj, typeof(ksDocument2DNotify).GUID, doc,
			-1, null, selfAdvise) {}

		// d3BeginRebuild - Начало перестроения модели
		public bool BeginRebuild()
		{ 
			if (m_SelfAdvise && FrmConfig.Instance.chb2DDocEvents.Checked)
			{
				string str = m_LibName + " --> BeginRebuild";
				str += "\nИмя документа = " + GetDocName();
				return Global.Kompas.ksYesNo(str) == 1 ? true : false;
			}
			return true;
		}


		// d3Rebuild - Модель перестроена
		public bool Rebuild()
		{   
			if (m_SelfAdvise && FrmConfig.Instance.chb2DDocEvents.Checked)
			{
				string str = m_LibName + " --> Rebuild";
				str += "\nИмя документа = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true;
		}


		// d3BeginChoiceMaterial - Начало выбора материала
		public bool BeginChoiceMaterial()
		{ 
			bool res = false;
			if (m_SelfAdvise && FrmConfig.Instance.chb2DDocEvents.Checked)
			{
				string str = m_LibName + " --> BeginChoiceMaterial";
				str += "\nИмя документа = " + GetDocName();
				res = Global.Kompas.ksYesNo(str) == 1 ? true : false;

				if (!res)
				{   
					ksTextLineParam parLine = (ksTextLineParam)Global.Kompas.GetParamStruct((short)StructType2DEnum.ko_TextLineParam);
					ksTextItemParam parItem = (ksTextItemParam)Global.Kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemParam);

					// Cоздали массив строк текста
					ksDynamicArray arMaterial = (ksDynamicArray)Global.Kompas.GetDynamicArray(ldefin2d.TEXT_LINE_ARR);
					if (arMaterial != null && parLine != null && parItem != null)
					{
						parLine.Init();
						parItem.Init();

						// Массив компонент строки текста
						ksDynamicArray item = (ksDynamicArray)parLine.GetTextItemArr();
						if (item != null) 
						{
							ksTextItemFont font = (ksTextItemFont)parItem.GetItemFont();
							if (font != null) 
							{
								// Создаем первую строку текста
								font.height = 10;   // Высота текста
								font.ksu = 1;       // Сужение текста
								font.color = 1000;  // Цвет
								font.bitVector = 1; // Битовый вектор (наклон, толщина, подчеркивание, тип составной части(дробь, отклонение, выражение типа суммы))
								parItem.s = "(1-я компонента)";
								// Добавили 1-ю компоненту  в массив компонент
								item.ksAddArrayItem(-1, parItem);

								font.height = 20;   // Высота текста
								font.ksu = 2;       // Сужение текста
								font.color = 2000;  // Цвет
								font.bitVector = 2; // Битовый вектор (наклон, толщина, подчеркивание, тип составной части(дробь, отклонение, выражение типа суммы))
								parItem.s = "(2-я компонента)";
								// Добавили 2-ю компоненту  в массив компонент
								item.ksAddArrayItem(-1, parItem);
				    
								parLine.style = 1;

								// 1-я строка текста состоит из двух компонент добавим 
								// строку текста в массив строк текста
								arMaterial.ksAddArrayItem(-1, parLine);

								ksDocument2D doc2D = (ksDocument2D)m_Doc;
								if (doc2D != null) 
								{
									doc2D.ksSetMaterialParam(arMaterial, 36.6);
								}
							}
						}
					}
				}
			}
			return res;
		}


		// d3СhoiceMaterial - Закончен выбор материала
		public bool ChoiceMaterial(string material, double density)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chb2DDocEvents.Checked)
			{
				string str = string.Empty;
				str = string.Format("{0} --> ChoiceMaterial\nmaterial = {1}\ndensity = {2}", m_LibName, material, density);
				str += "\nИмя документа = " + GetDocName();
				Global.Kompas.ksMessage(str);
			}
			return true;
		}


		// d2BeginInsertFragment - Начало вставки фрагмента (до диалога выбора имени)
		public bool BeginInsertFragment()
		{
			if (m_SelfAdvise && FrmConfig.Instance.chb2DDocEvents.Checked)
			{
				string str = m_LibName + " --> BeginInsertFragment";
				str += "\nИмя документа = " + GetDocName();
				return Global.Kompas.ksYesNo(str) == 1 ? true : false;
			}
			return true;
		}


		// d2LocalFragmentEdit
		public bool LocalFragmentEdit(object newDoc, bool newFrw)
		{
			if (m_SelfAdvise && FrmConfig.Instance.chb2DDocEvents.Checked)
			{
				string str = m_LibName + " --> LocalFragmentEdit";
				str += "\nИмя документа = " + GetDocName();
				return Global.Kompas.ksYesNo(str) == 1 ? true : false;
			}
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
