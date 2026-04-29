////////////////////////////////////////////////////////////////////////////////
//
// ApplicationEvent  - обработчик событий от приложения
//
////////////////////////////////////////////////////////////////////////////////

using Kompas6API5;

using System;
using Kompas6Constants;
using KAPITypes;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;

namespace Steps.NET
{
  public class ApplicationEvent : BaseEvent, ksKompasObjectNotify
  {
    public ApplicationEvent(object obj)
      : base(obj, typeof(ksKompasObjectNotify).GUID, null) { }


    // koApplicatinDestroy - Закрытие приложения
    public bool ApplicationDestroy()
    {
      TerminateEvents();
      return true;
    }


    // koBeginCloseAllDocument - Начало закрытия всех открытых документов
    public bool BeginCloseAllDocument()
    {
      Global.UpdateSlideBox(null);
      return true;
    }


    // koBeginCreate - Начало создания документа(до диалога выбора типа)
    public bool BeginCreate(int docType)
    {
      return true;
    }


    // koOpenDocumenBegin - Начало открытия документа
    public bool BeginOpenDocument(string fileName)
    {
      return true;
    }


    // koBeginOpenFile - Начало открытия документа(до диалога выбора имени)
    public bool BeginOpenFile()
    {
      return true;
    }


    // koActiveDocument - Переключение на другой активный документ
    public bool ChangeActiveDocument(object newDoc, int docType)
    {
      Global.UpdateSlideBox(newDoc);
      return true;
    }


    // koCreateDocument - Документ создан
    public bool CreateDocument(object newDoc, int docType)
    {
      Global.AdviseDoc(newDoc, docType);
      Global.UpdateSlideBox(newDoc);
      return true;
    }


    // koOpenDocumen - Документ открыт
    public bool OpenDocument(object newDoc, int docType)
    {
      Global.AdviseDoc(newDoc, docType);
      Global.UpdateSlideBox(newDoc);
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

    public static BaseEvent NewApplicationEvent()
    {
      ApplicationEvent res = null;
      if (!BaseEvent.FindEvent(typeof(ApplicationEvent), null))
      {
        res = new ApplicationEvent(Global.Kompas);
        res.Advise();
      }
      return res;
    }
  }
}
