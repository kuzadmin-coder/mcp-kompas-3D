////////////////////////////////////////////////////////////////////////////////
//
// ApplicationEvent  - обработчик событий от приложения
//
////////////////////////////////////////////////////////////////////////////////

using System;
using Kompas6API5;
using Kompas6Constants;
using KAPITypes;
using System.Runtime.InteropServices;

namespace Steps.NET
{
  public class ApplicationEvent : BaseEvent, ksKompasObjectNotify
  {
    public ApplicationEvent(object obj, bool selfAdvise)
      : base(obj, typeof(ksKompasObjectNotify).GUID,
      null, -1, null, selfAdvise)
    { }


    // koApplicatinDestroy - Закрытие приложения
    public bool ApplicationDestroy()
    {
      if (m_SelfAdvise && FrmConfig.Instance.chbAppEvents.Checked)
      {
        string str = m_LibName + " --> ApplicationEvent.ApplicationDestroy";
        Global.Kompas.ksMessage(str);
      }

      // Самоудаление
      TerminateEvents();
      Global.Kompas = null;
      return true;
    }


    // koBeginCloseAllDocument - Начало закрытия всех открытых документов
    public bool BeginCloseAllDocument()
    {
      bool res = true;
      if (m_SelfAdvise && FrmConfig.Instance.chbAppEvents.Checked)
      {
        string str = string.Empty;
        str = string.Format("{0} --> ApplicationEvent.BeginCloseAllDocument\nЗакрыть все?", m_LibName);
        res = Global.Kompas.ksYesNo(str) == 1 ? true : false;
      }
      return res;
    }


    // koBeginCreate - Начало создания документа(до диалога выбора типа)
    public bool BeginCreate(int docType)
    {
      bool res = true;
      if (m_SelfAdvise && FrmConfig.Instance.chbAppEvents.Checked)
      {
        string str = string.Empty;
        str = string.Format("{0} --> ApplicationEvent.BeginCreate\ndocType = {1}\n", m_LibName, docType);
        str += "Да - Создать чертеж\n" +
          "Нет - запустить стандартный диалог открытия файла\n" +
          "Отмена - не открывать файл";
        int comm = Global.Kompas.ksYesNo(str);
        switch (comm)
        {
          case 1:
            {
              ksDocumentParam docParam = (ksDocumentParam)Global.Kompas.GetParamStruct((short)StructType2DEnum.ko_DocumentParam);
              docParam.Init();
              docParam.type = (short)DocType.lt_DocSheetStandart;
              ksDocument2D doc = (ksDocument2D)Global.Kompas.Document2D();
              doc.ksCreateDocument(docParam);
              res = (doc.reference == 0); // Если документ создан стандартный диалог не нужен
              break;
            }
          case -1:
            res = false;
            break;
        }
      }
      return res;
    }


    // koOpenDocumenBegin - Начало открытия документа
    public bool BeginOpenDocument(string fileName)
    {
      if (m_SelfAdvise && FrmConfig.Instance.chbAppEvents.Checked)
      {
        string str = m_LibName + " --> ApplicationEvent.BeginOpenDocumen\nfileName = " + fileName;
        return Global.Kompas.ksYesNo(str) == 1 ? true : false;
      }
      else
        return true;
    }


    // koBeginOpenFile - Начало открытия документа(до диалога выбора имени)
    public bool BeginOpenFile()
    {
      bool res = true;
      if (m_SelfAdvise && FrmConfig.Instance.chbAppEvents.Checked)
      {
        string str = m_LibName + " --> ApplicationEvent.BeginOpenFile\n" +
          "Да - Создать чертеж\n" +
          "Нет - запустить стандартный диалог открытия файла\n" +
          "Отмена - не открывать файл";
        int comm = Global.Kompas.ksYesNo(str);
        switch (comm)
        {
          case 1:
            {
              ksDocumentParam docParam = (ksDocumentParam)Global.Kompas.GetParamStruct((short)StructType2DEnum.ko_DocumentParam);
              docParam.Init();
              docParam.type = (short)DocType.lt_DocSheetStandart;
              ksDocument2D doc = (ksDocument2D)Global.Kompas.Document2D();
              res = !doc.ksCreateDocument(docParam);
              break;
            }
          case -1:
            res = false;
            break;
        }
      }
      return res;
    }


    // koActiveDocument - Переключение на другой активный документ
    public bool ChangeActiveDocument(object newDoc, int docType)
    {
      if (m_SelfAdvise && FrmConfig.Instance.chbAppEvents.Checked)
      {
        string str = string.Empty;
        str = string.Format("{0} --> ApplicationEvent.ChangeActiveDocument\nnewDoc = {1}\ndocType = {2}", m_LibName, newDoc, docType);
        Global.Kompas.ksMessage(str);
      }
      return true;
    }


    // koCreateDocument - Документ создан
    public bool CreateDocument(object newDoc, int docType)
    {
      if (m_SelfAdvise && FrmConfig.Instance.chbAppEvents.Checked)
      {
        string str = string.Empty;
        str = string.Format("{0} --> ApplicationEvent.CreateDocument\nnewDoc = {1}\ndocType = {2}", m_LibName, newDoc, docType);
        Global.Kompas.ksMessage(str);
      }
      return true;
    }


    // koOpenDocumen - Документ открыт
    public bool OpenDocument(object newDoc, int docType)
    {
      if (m_SelfAdvise && FrmConfig.Instance.chbAppEvents.Checked)
      {
        string str = string.Empty;
        str = string.Format("{0} --> ApplicationEvent.OpenDocumen\nnewDoc = {1}\ndocType = {2}", m_LibName, newDoc, docType);
        Global.Kompas.ksMessage(str);
      }
      return true;
    }

    // koKeyDown - Событие клавиатуры
    public bool KeyDown(ref int key, int flags, bool system)
    {
      return true;
    }

    // koKeyUp - Событие клавиатуры
    public bool KeyUp(ref int key, int flags, bool system)
    {
      return true;
    }

    // koKeyPress - Событие клавиатуры
    public bool KeyPress(ref int key, bool system)
    {
      return true;
    }

    public bool BeginReguestFiles(int type, ref object files)
    {
      return true;
    }

    public bool BeginChoiceMaterial(int MaterialPropertyId)
    {
      return true;
    }
    public bool ChoiceMaterial(int MaterialPropertyId, string material, double density)
    {
      return true;
    }

    public bool IsNeedConvertToSavePrevious(object pDoc, int docType, int saveVersion, object saveToPreviusParam, ref bool needConvert)
    {
      return true;
    }

    public bool BeginConvertToSavePrevious(object pDoc, int docType, int saveVersion, object saveToPreviusParam)
    {
      return true;
    }

    public bool EndConvertToSavePrevious(object pDoc, int docType, int saveVersion, object saveToPreviusParam)
    {
      return true;
    }

    public bool ChangeTheme(int newTheme)
    {
      return true;
    }

    public bool BeginDragOpenFiles([MarshalAs(UnmanagedType.Struct)] ref object filesList, bool insert, ref bool filesListChanged)
    {
      return true;
    }
    public bool BeginChoiceMarking([MarshalAs(UnmanagedType.BStr)] string oldMarking, [MarshalAs(UnmanagedType.BStr)] out string newMarking)
    {
      newMarking = oldMarking;
      return true;
    }

  }
}
