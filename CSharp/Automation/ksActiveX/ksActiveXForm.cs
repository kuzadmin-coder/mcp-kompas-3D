using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Kompas6API5;
using Kompas6Constants;
using Kompas6Constants3D;

namespace ksActiveX
{
  public partial class ksAxtiveXForm : Form
  {
    public ksAxtiveXForm()
    {
      InitializeComponent();
    }

    private void button1_Click(object sender, EventArgs e)
    {
      KompasObject kompas = axKGAX3D.GetKompasObject();
      if (kompas == null)
      {
        MessageBox.Show("BAD KompasObject");
        return;
      }

      ksDocument3D doc = (ksDocument3D)kompas.ActiveDocument3D();
      if (doc == null)
      {
        MessageBox.Show("BAD Document3D");
        return;
      }

      ksPart part = (ksPart)doc.GetPart((short)Part_Type.pTop_Part);	//новый компонент
      if (part == null)
      {
        MessageBox.Show("BAD Part");
        return;
      }

      ksEntity entitySketch = (ksEntity)part.NewEntity((short)Obj3dType.o3d_sketch);
      if (entitySketch == null)
      {
        MessageBox.Show("BAD EntitySketch");
        return;
      }

      // интерфейс свойств эскиза
      ksSketchDefinition sketchDef = (ksSketchDefinition)entitySketch.GetDefinition();
      if (sketchDef == null)
      {
        MessageBox.Show("BAD SketchDefinition");
        return;
      }

      // получим интерфейс базовой плоскости XOY
      ksEntity basePlane = (ksEntity)part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);
      sketchDef.SetPlane(basePlane);	// установим плоскость XOY базовой для эскиза
      sketchDef.angle = 45;			// угол поворота эскиза
      entitySketch.Create();			// создадим эскиз

      // интерфейс редактора эскиза
      ksDocument2D sketchEdit = (ksDocument2D)sketchDef.BeginEdit();
      // введем новый эскиз - квадрат
      sketchEdit.ksLineSeg(50, 50, -50, 50, 1);
      sketchEdit.ksLineSeg(50, -50, -50, -50, 1);
      sketchEdit.ksLineSeg(50, -50, 50, 50, 1);
      sketchEdit.ksLineSeg(-50, -50, -50, 50, 1);
      sketchDef.EndEdit();	//завершение редактирования эскиза

      ksEntity entityExtr = (ksEntity)part.NewEntity((short)Obj3dType.o3d_bossExtrusion);
      if (entityExtr == null)
      {
        MessageBox.Show("BAD EntityExtrusion");
        return;
      }

      // интерфейс свойств базовой операции выдавливания
      ksBossExtrusionDefinition extrusionDef = (ksBossExtrusionDefinition)entityExtr.GetDefinition();
      if (extrusionDef == null)
      {
        MessageBox.Show("BAD ExtrusionDefinition");
        return;
      }

      extrusionDef.directionType = (short)Direction_Type.dtNormal;	// направление выдавливания
      extrusionDef.SetSideParam(true, (short)End_Type.etBlind, 200, 0, false);
      extrusionDef.SetThinParam(true, (short)Direction_Type.dtBoth, 10, 10);		// тонкая стенка в два направления
      extrusionDef.SetSketch(entitySketch);					// эскиз операции выдавливания
      entityExtr.Create();									// создать операцию

      axKGAX3D.ZoomEntireDocument();
    }

    private void button2_Click(object sender, EventArgs e)
    {
      axKGAX3D.ZoomEntireDocument();
    }

    private void Form1_FormClosed(object sender, FormClosedEventArgs e)
    {
       KompasObject kompas = axKGAX3D.GetKompasObject();
       if (kompas != null )
       {
         axKGAX3D.CloseAll();
         kompas.Quit();
       }
    }

    private void radioButton1_CheckedChanged(object sender, EventArgs e)
    {
      axKGAX3D.Document3DDrawMode =  KGAXLib.KDocument3DDrawMode.vt_WireframeMode;
    }

    private void radioButton2_CheckedChanged(object sender, EventArgs e)
    {
      axKGAX3D.Document3DDrawMode = KGAXLib.KDocument3DDrawMode.vt_HiddenRemovedMode;
    }

    private void radioButton3_CheckedChanged(object sender, EventArgs e)
    {
      axKGAX3D.Document3DDrawMode = KGAXLib.KDocument3DDrawMode.vt_HiddenThinMode;
    }

    private void radioButton4_CheckedChanged(object sender, EventArgs e)
    {
      axKGAX3D.Document3DDrawMode = KGAXLib.KDocument3DDrawMode.vt_ShadedMode;
    }

    private void axKGAX3D_OnKgMouseDown(object sender, AxKGAXLib._DKGAXEvents_OnKgMouseDownEvent e)
    {
      e.proceed = true;
    }
  }
}