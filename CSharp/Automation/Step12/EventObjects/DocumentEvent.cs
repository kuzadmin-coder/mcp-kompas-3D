////////////////////////////////////////////////////////////////////////////////
//
// DocumentEvent  - обработчик событий от документа
//
////////////////////////////////////////////////////////////////////////////////

using Kompas6API5;
using KompasAPI7;

using System;
using Kompas6Constants;
using KAPITypes;


using reference = System.Int32;
using System.Runtime.InteropServices;

namespace Steps.NET
{
  public class DocumentEvent : BaseEvent, Kompas6API5.ksDocumentFileNotify
  {
    public DocumentEvent(object doc)
      : base(doc, typeof(Kompas6API5.ksDocumentFileNotify).GUID, doc) { }


    // kdBeginCloseDocument - Начало закрытия документа
    public bool BeginCloseDocument()
    {
      return true;
    }


    // kdCloseDocument - Документ закрыт
    public bool CloseDocument()
    {
      BaseEvent.TerminateEvents(null, this.m_Doc);
      Global.UpdateSlideBox(null);
      return true;
    }


    // kdBeginSaveDocument - Начало сохранения документа
    public bool BeginSaveDocument(string fileName)
    {
      return true;
    }


    // kdSaveDocument - Документ сохранен
    public bool SaveDocument()
    {
      return true;
    }


    // kdActiveDocument - Документ активизирован.
    public bool Activate()
    {
      Global.UpdateSlideBox(null);
      return true;
    }


    // kdDeactiveDocument - Документ деактивизирован.
    public bool Deactivate()
    {
      Global.UpdateSlideBox(null);
      return true;
    }


    // kdBeginSaveAsDocument - Начало сохранения документа c другим именем (до диалога выбора имени)
    public bool BeginSaveAsDocument()
    {
      return true;
    }


    // kdDocumentFrameOpen - Окно документа открылось
    public bool DocumentFrameOpen(object v)
    {
      Global.UpdateSlideBox(null);
      return true;
    }


    public bool ProcessActivate(int id)
    {
      return true;
    }


    public bool ProcessDeactivate(int id)
    {
      Global.UpdateSlideBox(null);
      return true;
    }

    public bool BeginProcess(int iD)
    {
      return true;
    }
    public bool EndProcess(int iD, bool Success)
    {
      return true;
    }

    public bool BeginAutoSaveDocument(string fileName)
    {
      return true;
    }

    public bool AutoSaveDocument()
    {
      return true;
    }
    public bool ActivateToolbarSet([MarshalAs(UnmanagedType.BStr)] string OldToolbarSetId, [MarshalAs(UnmanagedType.BStr)] string NewToolbarSetId)
    {
      return true;
    }


    // Создать обработчик событий документа
    public static BaseEvent NewDocumentEvent(object doc)
    {
      DocumentEvent res = null;
      if (doc != null)
      {
        if (!BaseEvent.FindEvent(typeof(DocumentEvent), doc))
        {
          res = new DocumentEvent(doc);
          res.Advise();
        }
      }
      return res;
    }


    // Создать обработчик событий 2D документа
    public static BaseEvent NewDocument2DEvent(object doc)
    {
      Document2DEvent res = null;
      if (doc != null)
      {
        if (BaseEvent.FindEvent(typeof(Document2DEvent), doc))
        {
          ksDocument2D doc2d = (ksDocument2D)doc;
          if (doc2d != null)
          {
            NewDocumentEvent(doc);  // чтобы при закрытии документа отписаться
            res = new Document2DEvent(doc2d.GetDocument2DNotify(), doc);
            res.Advise();
          }
        }
      }
      return res;
    }

    // Создать обработчик событий 3D документа
    public static BaseEvent NewDocument3DEvent(object doc)
    {
      Document3DEvent res = null;
      if (doc != null)
      {
        if (BaseEvent.FindEvent(typeof(Document3DEvent), doc))
        {
          ksDocument3D doc3d = (ksDocument3D)doc;
          if (doc3d != null)
          {
            NewDocumentEvent(doc);  // чтобы при закрытии документа отписаться
            res = new Document3DEvent(doc3d.GetDocument3DNotify(), doc);
            res.Advise();
          }
        }
      }
      return res;
    }

  }
}
