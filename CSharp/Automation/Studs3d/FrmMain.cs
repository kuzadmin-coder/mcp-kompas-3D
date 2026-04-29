using Kompas6API5;
using KompasAPI7;


using System;
using System.IO;
using System.Drawing;
using System.Resources;
using System.Reflection;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Kompas6Constants;
using Kompas6Constants3D;
using KAPITypes;

using reference = System.Int32;

namespace Steps.NET
{
	public class FrmMain : System.Windows.Forms.Form
	{
		#region Designer declarations
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.CheckBox chbSpecObject;
		private System.Windows.Forms.ListView lvListView;
		private System.Windows.Forms.Button btnOK;
		private System.Windows.Forms.Button btnHelp;
		private System.Windows.Forms.Button btnCancel;
		private System.Windows.Forms.ComboBox cbEndLength;
		private System.Windows.Forms.ComboBox cbDiameter;
		private System.Windows.Forms.ComboBox cbLength;
		private System.Windows.Forms.RadioButton rbVar1;
		private System.Windows.Forms.RadioButton rbVar2;
		private System.Windows.Forms.ComboBox cbAccClass;
		private System.Windows.Forms.ComboBox cbGOST;
		private System.Windows.Forms.ComboBox cbMaterial;
		private System.Windows.Forms.CheckBox chbShortStep;
		private System.Windows.Forms.Panel panel;
		private System.Windows.Forms.CheckBox chbSimpled;
		#endregion

		#region Custom declarations
		private float p;						// шаг - вспомогательное поле
		private float massa;					// масса - вспомогательное поле
		private int bitMapId;					// идентификатор слайда

		private static int MAX_LENGTH = 1000;	// максимальная длина стержня
		private static int MIN_LENGTH = 0;		// минимальная длина стержня

		private Pin pinObj;
		#endregion

		#region Instance etc...
		private FrmMain()
		{
			InitializeComponent();

		}


		private static FrmMain instance;
		public static FrmMain Instance
		{
			get
			{
				if (instance == null)
					instance = new FrmMain();
				return instance;

			}
		}

		protected override void Dispose(bool disposing)
		{
			if(disposing)
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose(disposing);
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.label1 = new System.Windows.Forms.Label();
			this.cbEndLength = new System.Windows.Forms.ComboBox();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cbLength = new System.Windows.Forms.ComboBox();
			this.label3 = new System.Windows.Forms.Label();
			this.cbDiameter = new System.Windows.Forms.ComboBox();
			this.label2 = new System.Windows.Forms.Label();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.rbVar2 = new System.Windows.Forms.RadioButton();
			this.rbVar1 = new System.Windows.Forms.RadioButton();
			this.panel = new System.Windows.Forms.Panel();
			this.cbAccClass = new System.Windows.Forms.ComboBox();
			this.label4 = new System.Windows.Forms.Label();
			this.cbGOST = new System.Windows.Forms.ComboBox();
			this.label5 = new System.Windows.Forms.Label();
			this.cbMaterial = new System.Windows.Forms.ComboBox();
			this.label6 = new System.Windows.Forms.Label();
			this.chbShortStep = new System.Windows.Forms.CheckBox();
			this.chbSpecObject = new System.Windows.Forms.CheckBox();
			this.chbSimpled = new System.Windows.Forms.CheckBox();
			this.lvListView = new System.Windows.Forms.ListView();
			this.btnOK = new System.Windows.Forms.Button();
			this.btnHelp = new System.Windows.Forms.Button();
			this.btnCancel = new System.Windows.Forms.Button();
			this.groupBox1.SuspendLayout();
			this.groupBox2.SuspendLayout();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(8, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(184, 19);
			this.label1.TabIndex = 0;
			this.label1.Text = "Длина ввинчиваемого конца b1";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cbEndLength
			// 
			this.cbEndLength.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbEndLength.Location = new System.Drawing.Point(184, 8);
			this.cbEndLength.Name = "cbEndLength";
			this.cbEndLength.Size = new System.Drawing.Size(64, 21);
			this.cbEndLength.TabIndex = 1;
			this.cbEndLength.SelectedIndexChanged += new System.EventHandler(this.cbEndLength_SelectedIndexChanged);
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.cbLength);
			this.groupBox1.Controls.Add(this.label3);
			this.groupBox1.Controls.Add(this.cbDiameter);
			this.groupBox1.Controls.Add(this.label2);
			this.groupBox1.Location = new System.Drawing.Point(8, 32);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(168, 80);
			this.groupBox1.TabIndex = 2;
			this.groupBox1.TabStop = false;
			// 
			// cbLength
			// 
			this.cbLength.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbLength.Location = new System.Drawing.Point(80, 46);
			this.cbLength.Name = "cbLength";
			this.cbLength.Size = new System.Drawing.Size(72, 21);
			this.cbLength.TabIndex = 7;
			this.cbLength.SelectedIndexChanged += new System.EventHandler(this.cbLength_SelectedIndexChanged);
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(8, 46);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(56, 19);
			this.label3.TabIndex = 6;
			this.label3.Text = "Длина";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cbDiameter
			// 
			this.cbDiameter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbDiameter.Location = new System.Drawing.Point(80, 16);
			this.cbDiameter.Name = "cbDiameter";
			this.cbDiameter.Size = new System.Drawing.Size(72, 21);
			this.cbDiameter.TabIndex = 5;
			this.cbDiameter.SelectedIndexChanged += new System.EventHandler(this.cbDiameter_SelectedIndexChanged);
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(8, 16);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(56, 19);
			this.label2.TabIndex = 4;
			this.label2.Text = "Диаметр";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.rbVar2);
			this.groupBox2.Controls.Add(this.rbVar1);
			this.groupBox2.Location = new System.Drawing.Point(8, 112);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(168, 72);
			this.groupBox2.TabIndex = 3;
			this.groupBox2.TabStop = false;
			// 
			// rbVar2
			// 
			this.rbVar2.Location = new System.Drawing.Point(16, 40);
			this.rbVar2.Name = "rbVar2";
			this.rbVar2.TabIndex = 1;
			this.rbVar2.Text = "Исполнение 2";
			this.rbVar2.CheckedChanged += new System.EventHandler(this.rbVar2_CheckedChanged);
			// 
			// rbVar1
			// 
			this.rbVar1.Checked = true;
			this.rbVar1.Location = new System.Drawing.Point(16, 16);
			this.rbVar1.Name = "rbVar1";
			this.rbVar1.TabIndex = 0;
			this.rbVar1.TabStop = true;
			this.rbVar1.Text = "Исполнение 1";
			this.rbVar1.CheckedChanged += new System.EventHandler(this.rbVar1_CheckedChanged);
			// 
			// panel
			// 
			this.panel.BackColor = System.Drawing.SystemColors.Window;
			this.panel.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.panel.Location = new System.Drawing.Point(184, 37);
			this.panel.Name = "panel";
			this.panel.Size = new System.Drawing.Size(168, 147);
			this.panel.TabIndex = 4;
			// 
			// cbAccClass
			// 
			this.cbAccClass.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbAccClass.Location = new System.Drawing.Point(104, 192);
			this.cbAccClass.Name = "cbAccClass";
			this.cbAccClass.Size = new System.Drawing.Size(72, 21);
			this.cbAccClass.TabIndex = 9;
			this.cbAccClass.SelectedIndexChanged += new System.EventHandler(this.cbAccClass_SelectedIndexChanged);
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(8, 192);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(96, 19);
			this.label4.TabIndex = 8;
			this.label4.Text = "Класс точности";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cbGOST
			// 
			this.cbGOST.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbGOST.Location = new System.Drawing.Point(248, 192);
			this.cbGOST.Name = "cbGOST";
			this.cbGOST.Size = new System.Drawing.Size(104, 21);
			this.cbGOST.TabIndex = 11;
			this.cbGOST.SelectedIndexChanged += new System.EventHandler(this.cbGOST_SelectedIndexChanged);
			// 
			// label5
			// 
			this.label5.Location = new System.Drawing.Point(184, 192);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(56, 19);
			this.label5.TabIndex = 10;
			this.label5.Text = "ГОСТ";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cbMaterial
			// 
			this.cbMaterial.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbMaterial.Location = new System.Drawing.Point(248, 220);
			this.cbMaterial.Name = "cbMaterial";
			this.cbMaterial.Size = new System.Drawing.Size(104, 21);
			this.cbMaterial.TabIndex = 13;
			this.cbMaterial.SelectedIndexChanged += new System.EventHandler(this.cbMaterial_SelectedIndexChanged);
			// 
			// label6
			// 
			this.label6.Location = new System.Drawing.Point(184, 220);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(64, 19);
			this.label6.TabIndex = 12;
			this.label6.Text = "Материал";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// chbShortStep
			// 
			this.chbShortStep.Location = new System.Drawing.Point(16, 220);
			this.chbShortStep.Name = "chbShortStep";
			this.chbShortStep.TabIndex = 14;
			this.chbShortStep.Text = "Мелкий шаг";
			this.chbShortStep.CheckedChanged += new System.EventHandler(this.chbShortStep_CheckedChanged);
			// 
			// chbSpecObject
			// 
			this.chbSpecObject.Checked = true;
			this.chbSpecObject.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chbSpecObject.Location = new System.Drawing.Point(16, 248);
			this.chbSpecObject.Name = "chbSpecObject";
			this.chbSpecObject.Size = new System.Drawing.Size(208, 24);
			this.chbSpecObject.TabIndex = 15;
			this.chbSpecObject.Text = "Создать объект спецификации";
			this.chbSpecObject.CheckedChanged += new System.EventHandler(this.chbSpecObject_CheckedChanged);
			// 
			// chbSimpled
			// 
			this.chbSimpled.Location = new System.Drawing.Point(248, 248);
			this.chbSimpled.Name = "chbSimpled";
			this.chbSimpled.Size = new System.Drawing.Size(88, 24);
			this.chbSimpled.TabIndex = 16;
			this.chbSimpled.Text = "Упрощенно";
			this.chbSimpled.CheckedChanged += new System.EventHandler(this.chbSimpled_CheckedChanged);
			// 
			// lvListView
			// 
			this.lvListView.Location = new System.Drawing.Point(8, 280);
			this.lvListView.Name = "lvListView";
			this.lvListView.Size = new System.Drawing.Size(344, 48);
			this.lvListView.TabIndex = 17;
			this.lvListView.View = System.Windows.Forms.View.Details;
			// 
			// btnOK
			// 
			this.btnOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
			this.btnOK.Location = new System.Drawing.Point(216, 336);
			this.btnOK.Name = "btnOK";
			this.btnOK.Size = new System.Drawing.Size(64, 23);
			this.btnOK.TabIndex = 18;
			this.btnOK.Text = "ОК";
			this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
			// 
			// btnHelp
			// 
			this.btnHelp.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnHelp.Location = new System.Drawing.Point(8, 336);
			this.btnHelp.Name = "btnHelp";
			this.btnHelp.Size = new System.Drawing.Size(64, 23);
			this.btnHelp.TabIndex = 19;
			this.btnHelp.Text = "Справка";
			this.btnHelp.Click += new System.EventHandler(this.btnHelp_Click);
			// 
			// btnCancel
			// 
			this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btnCancel.Location = new System.Drawing.Point(288, 336);
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.Size = new System.Drawing.Size(64, 23);
			this.btnCancel.TabIndex = 20;
			this.btnCancel.Text = "Отмена";
			// 
			// FrmMain
			// 
			this.AcceptButton = this.btnOK;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.CancelButton = this.btnCancel;
			this.ClientSize = new System.Drawing.Size(362, 367);
			this.Controls.Add(this.btnCancel);
			this.Controls.Add(this.btnHelp);
			this.Controls.Add(this.btnOK);
			this.Controls.Add(this.lvListView);
			this.Controls.Add(this.chbSimpled);
			this.Controls.Add(this.chbSpecObject);
			this.Controls.Add(this.chbShortStep);
			this.Controls.Add(this.cbMaterial);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.cbGOST);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.panel);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.cbEndLength);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.cbAccClass);
			this.Controls.Add(this.label4);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "FrmMain";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Шпильки";
			this.Load += new System.EventHandler(this.FrmMain_Load);
			this.Paint += new System.Windows.Forms.PaintEventHandler(this.FrmMain_Paint);
			this.groupBox1.ResumeLayout(false);
			this.groupBox2.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		public DialogResult Execute(Pin pin)
		{
			pinObj = pin;
			InitParamList();
			return this.ShowDialog();
		}
		#endregion

		private void CalcMassa()
		{
			massa = pinObj.PinParam.f != 0 & Pin.ISPOLN != 0
				? pinObj.PinParam.m2 : pinObj.PinParam.m1;
			massa *= (float)(pinObj.PinParam.indexMassa == 0
				? 1 : pinObj.PinParam.indexMassa == 1
				? 0.356 : pinObj.PinParam.indexMassa == 3
				? 0.97 : 1.08);
		}

		private void ChoicePitch()
		{
			p = pinObj.PinParam.f != 0 & Pin.PITCH != 0
				? pinObj.PinParam.p2 : pinObj.PinParam.p1;
		}

		private void InitParamList()
		{
			CalcMassa();
			ChoicePitch();

			int width = Convert.ToInt32(lvListView.Width / 5);
			lvListView.Columns.Clear();
			lvListView.Columns.Add(pinObj.LoadString("STR247"), width, HorizontalAlignment.Left);	//l0 Гаечный конец
			lvListView.Columns.Add(pinObj.LoadString("STR248"), width, HorizontalAlignment.Left);	//l1 Ввинчиваемый конец
			lvListView.Columns.Add(pinObj.LoadString("STR231"), width, HorizontalAlignment.Left);	//"Шаг резьбы"
			lvListView.Columns.Add(pinObj.LoadString("STR236"), width, HorizontalAlignment.Left);	//"Фаска"
			lvListView.Columns.Add(pinObj.LoadString("STR237"), width, HorizontalAlignment.Left);	//"Масса 1000 шт"

			CalcMassa();

			lvListView.Items.Clear();
			ListViewItem item = new ListViewItem();
			item.Text = pinObj.PinParam.b.ToString();
			item.SubItems.Add(pinObj.PinParam.b1.ToString());
			item.SubItems.Add(p.ToString());
			item.SubItems.Add(pinObj.PinParam.c.ToString());
			item.SubItems.Add(massa.ToString());
			lvListView.Items.Add(item);
		}


		/// <summary>
		/// Выдаёт сообщение о ошибке, закрывает диалог
		/// </summary>
		/// <param name="id"></param>
		void ErrorDialog(string id)
		{
			string buf = pinObj.LoadString(id);			// строка из ресурса
			Kompas.Instance.KompasObject.ksError(buf);	// сообщение об ошибке
			this.DialogResult = DialogResult.Cancel;
			Close();
		}


		/// <summary>
		/// Заполняет список диаметров
		/// </summary>
		void FillDiametr()
		{
			bool failDB = true;	// предполагаем ошибку
			ksUserParam tmpL = (ksUserParam)Kompas.Instance.KompasObject.GetParamStruct((short)StructType2DEnum.ko_UserParam);
			ksLtVariant item = (ksLtVariant)Kompas.Instance.KompasObject.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
			ksDynamicArray arr = (ksDynamicArray)Kompas.Instance.KompasObject.GetDynamicArray(ldefin2d.LTVARIANT_ARR);
			if (tmpL != null && item != null && arr != null)
			{
				tmpL.Init();
				tmpL.SetUserArray(arr);
				item.Init();
				item.shortVal = 0;
				arr.ksAddArrayItem(-1, item);
				item.strVal = string.Empty;
				arr.ksAddArrayItem(-1, item);

				cbDiameter.Items.Clear();	// очищаем список

				float lMin = MAX_LENGTH;	// минимальное значение в списке
				float lMax = MIN_LENGTH;	// максимальное значение в списке
				bool enterInRange = false;	// Текущее значение диаметра не входит в новые значения

				if (pinObj.Data.ksCondition(pinObj.Base.bg,
					pinObj.Base.rg1, "t = 1") == 1)
				{ 
					int i = 1;
					do
					{
						// Просматриваем все записи базы данных
						i = pinObj.Data.ksReadRecord(pinObj.Base.bg,
							pinObj.Base.rg1,
							tmpL);	// Считываем данное
						if (i != 0)
						{
							// если запись найдена
							arr.ksGetArrayItem(1, item);
							string c = item.strVal;
							cbDiameter.Items.Add(c);		// Заносим его в список
							float ln = Convert.ToSingle(c.Replace('.', ','));	// Преобразуем его в число
							if (pinObj.PinParam.d == ln)
								enterInRange = true;	// Текущее значение диаметра входит в новые значения
							if (ln < lMin)
								lMin = ln;				// находим минимальное значение
							if (ln > lMax)
								lMax = ln;				// находим максимальное значение

							failDB = false;				// ошибки нет
						}
					}
					while (i != 0);

					if (!enterInRange)			// если текущее значение диаметра не входит в в новые значения
						pinObj.PinParam.d = lMin;	// присваиваем диаметру минимальное значение
					if (pinObj.PinParam.d < lMin)	// если диаметр выходит за минимальную границу новых значений
						pinObj.PinParam.d = lMin;	// присваиваем ему минимальное значение
					if (pinObj.PinParam.d > lMax)	// если диаметр выходит за максимальную границу новых значений
						pinObj.PinParam.d = lMax;	// присваиваем ему максимальное значение

					string s = string.Format("{0}",
						pinObj.PinParam.d);					// преобразуем текущее значение диаметра в строку
					for (int j = 0; j < cbDiameter.Items.Count; j ++)
					{
						string str = cbDiameter.Items[j].ToString();
						if (str == s)
							cbDiameter.SelectedIndex = j;	// выделяем его в списке
					}
				}
			}
			if (failDB)						// если работа с БД не была удачной
				ErrorDialog("DIAM_ERROR");	// сообщаем об ошибке
		}


		/// <summary>
		/// Заполняет список длин
		/// </summary>
		void FillLenght()
		{
			string buf = string.Empty;
			string buf1 = string.Empty;
			bool failDB = true;	// предполагаем ошибку
	
			ksUserParam tmpL = (ksUserParam)Kompas.Instance.KompasObject.GetParamStruct((short)StructType2DEnum.ko_UserParam);
			ksLtVariant item = (ksLtVariant)Kompas.Instance.KompasObject.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
			ksDynamicArray arr = (ksDynamicArray)Kompas.Instance.KompasObject.GetDynamicArray(ldefin2d.LTVARIANT_ARR);
			if (tmpL != null && item != null && arr != null)
			{
				tmpL.Init();
				tmpL.SetUserArray(arr);
				item.Init();
				item.shortVal = 0;
				arr.ksAddArrayItem(-1, item);
				item.strVal = string.Empty;
				arr.ksAddArrayItem(-1, item);
				item.floatVal = 0;
				arr.ksAddArrayItem(-1, item);

				// формируем условие запроса к БД
				buf = pinObj.LoadString("STR_245");
				buf1 = string.Format(buf, pinObj.PinParam.d);
				if (pinObj.Data.ksCondition(pinObj.Base.bs, pinObj.Base.rs1, buf1) == 1)
				{
					cbLength.Items.Clear();	// очищаем список

					float lMin = MAX_LENGTH;	// минимальное значение в списке
					float lMax = MIN_LENGTH;	// максимальное значение в списке
					bool enterInRange = false;	// Текущее значение длины не входит в новые значения

					int i = 1;
					do
					{
						// Просматриваем все записи базы данных
						i = pinObj.Data.ksReadRecord(pinObj.Base.bs, pinObj.Base.rs1, tmpL);	// Считываем данное
						if (i != 0)
						{
							// если запись найдена
							arr.ksGetArrayItem(1, item);
							string c = item.strVal;
							cbLength.Items.Add(c);			// Заносим его в список
							float ln = Convert.ToSingle(c.Replace('.', ','));	// Преобразуем его в число
							if (pinObj.PinParam.l == ln)
								enterInRange = true;		// Текущее значение длины входит в новые значения
							if (ln < lMin)
								lMin = ln;					// находим минимальное значение
							if (ln > lMax)
								lMax = ln;					// находим максимальное значение

							failDB = false;					// ошибки нет
						}
					}
					while (i != 0);

					if (!enterInRange)			// если текущее значение длины не входит в в новые значения
						pinObj.PinParam.l = lMin;	// присваиваем длине минимальное значение
					if (pinObj.PinParam.l < lMin)	// если диаметр выходит за минимальную границу новых значений
						pinObj.PinParam.l = lMin;	// присваиваем ему минимальное значение
					if (pinObj.PinParam.l > lMax)	// если диаметр выходит за максимальную границу новых значений
						pinObj.PinParam.l = lMax;	// присваиваем ему максимальное значение

					buf1 = string.Format("{0}", pinObj.PinParam.l);	// преобразуем текущее значение длины в строку
					for (int j = 0; j < cbLength.Items.Count; j ++)
					{
						string str = cbLength.Items[j].ToString();
						if (str == buf1)
							cbLength.SelectedIndex = j;				// выделяем его в списке
					}
				}
			}

			if (failDB)						// если работа с БД не была удачной
				ErrorDialog("LEN_ERROR");	// сообщаем об ошибке
		}


		private void FrmMain_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{
			// Вывод слайда
			Kompas.Instance.KompasObject.ksDrawSlide(panel.Handle.ToInt32(), bitMapId);
		}

		private void cbGOST_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			// Изменение госта в списке ГОСТов
			string str = cbGOST.Text;	// считываем строку
			int index = str.IndexOf('-');
			if (index > 0)                   // удаляем из ГОСТа год
				str = str.Substring(0, index);
			short gost = Convert.ToInt16(str); // преобразуем в число
			if (gost != pinObj.PinParam.gost)
			{
				pinObj.PinParam.gost = gost;

				cbEndLength.SelectedIndex = pinObj.GetTypeForGost(pinObj.PinParam.gost);	// выделение длины ввинчиваемого конца l1
				cbAccClass.SelectedIndex = pinObj.GetKlassForGost(pinObj.PinParam.gost);	// выделение класса точности

				if (pinObj.BaseOpened)		// если БД открыта
					pinObj.CloseBase();	// закрываем старую БД

				// Открываем новую БД и считываем текущие параметры
				if (!pinObj.OpenBase(true))
				{
					DialogResult = DialogResult.Cancel;
					Close();
					return;
				}
				if (pinObj.ReadPinBase(pinObj.PinParam.d) == 0)
				{
					DialogResult = DialogResult.Cancel;
					Close();
					return;
				}

				if (pinObj.ReadPinStBase() == -1)
				{
					DialogResult = DialogResult.Cancel;
					Close();
					return;
				}

				FrmMain_Load(null, null);
			}
		}

		private void chbSpecObject_CheckedChanged(object sender, System.EventArgs e)
		{
			if (chbSpecObject.Checked)
				pinObj.MacroParam.flagAttr = 1;
			else
				pinObj.MacroParam.flagAttr = 0;
		}

		private void cbMaterial_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			int k = cbMaterial.SelectedIndex;
			if (k > -1)
			{
				pinObj.PinParam.indexMassa = (byte)k;
				InitParamList();
			}
		}

		private void btnHelp_Click(object sender, System.EventArgs e)
		{
			// Вызов справки
      string err = pinObj.LoadString("STR230");	// "Файл constr.chm не найден"
			Kompas.Instance.KompasObject.ksError(err);
		}

		private void rbVar2_CheckedChanged(object sender, System.EventArgs e)
		{
			if (rbVar1.Checked)
				pinObj.PinParam.f = Convert.ToInt16(pinObj.PinParam.f != 0 | Pin.ISPOLN != 0);
			else
				pinObj.PinParam.f = Convert.ToInt16((pinObj.PinParam.f != 0 & Pin.ISPOLN != 0) ? 0 : 1);

			bitMapId = pinObj.ChoiceBMP(pinObj.PinParam.gost, pinObj.PinParam.f);	// выбор слайда
			InitParamList();			// обновляем список параметров
			FrmMain_Paint(null, null);	// перерисовываем слайд
		}

		private void rbVar1_CheckedChanged(object sender, System.EventArgs e)
		{
			if (rbVar1.Checked)
				pinObj.PinParam.f = Convert.ToInt16((pinObj.PinParam.f != 0 & Pin.ISPOLN != 0) ? 0 : 1);
			else
				pinObj.PinParam.f = Convert.ToInt16(pinObj.PinParam.f != 0 | Pin.ISPOLN != 0);

			bitMapId = pinObj.ChoiceBMP(pinObj.PinParam.gost, pinObj.PinParam.f);	// выбор слайда
			InitParamList();			// обновляем список параметров
			FrmMain_Paint(null, null);	// перерисовываем слайд
		}

		private void chbShortStep_CheckedChanged(object sender, System.EventArgs e)
		{
			if (chbShortStep.Checked)
				pinObj.PinParam.f = (short)(Convert.ToInt32(pinObj.PinParam.f) | Convert.ToInt32(Pin.PITCH));
			else
				pinObj.PinParam.f &= (short)~ Pin.PITCH;
			InitParamList();
		}

		private void cbLength_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			// Изменение в списке длин - считываем новые параметры из БД
			string str = cbLength.Text;			// считываем строку
			float l = Convert.ToSingle(str);	// преобразуем в число
			if (Math.Abs(l - pinObj.PinParam.l) > 0.001)
			{
				pinObj.PinParam.l = l; //нашли длину для k индекса в списке длин
				if (pinObj.ReadPinStBase() == -1)
				{
					DialogResult = DialogResult.Cancel;
					Close();
					return;
				}
				// Чтение БД

				bool ew = pinObj.PinParam.f != 0 & Pin.ALLST != 0 ? false : true;
				rbVar2.Enabled = ew;
				rbVar1.Checked = !ew;

				bitMapId = pinObj.ChoiceBMP(pinObj.PinParam.gost, pinObj.PinParam.f); // выбор слайда
				InitParamList();                        // обновляем список параметров
				FrmMain_Paint(null, null);              // перерисовываем слайд
			}
		}

		private void cbDiameter_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			// Изменение в списке диаметров - обновляем список длин и считываем новые параметры из БД
			string c = cbDiameter.Text;		// считываем строку
			float dr = Convert.ToSingle(c);	// преобразуем в число
			if (Math.Abs(dr - pinObj.PinParam.d) > 0.001)
			{
				if (pinObj.ReadPinBase(dr) == 0)
				{
					DialogResult = DialogResult.Cancel;
					Close();
					return;
				}
				// Чтение БД

				FillLenght();	// заполнение списка длин

				if (pinObj.ReadPinStBase() == -1)
				{
					DialogResult = DialogResult.Cancel;
					Close();
					return;
				}
				// Чтение БД

				bool ew = pinObj.PinParam.f != 0 & Pin.ALLST != 0 ? false : true;
				rbVar2.Enabled = ew;
				rbVar1.Checked = !ew;

				bitMapId = pinObj.ChoiceBMP(pinObj.PinParam.gost, pinObj.PinParam.f);	// выбор слайда
				InitParamList();					// обновляем список параметров
				this.FrmMain_Paint(null, null);	// перерисовываем слайд
			}
		}

		private void cbEndLength_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			// Изменение в списке "Длина ввинчиваемого конца"
			int k = cbEndLength.SelectedIndex;
			int k1 = cbAccClass.SelectedIndex;
			if ((k > -1) && (k1 > -1))
			{
				cbGOST.SelectedIndex = pinObj.GetGostForTypeAndKlass(k, k1);
			}
		}

		private void chbSimpled_CheckedChanged(object sender, System.EventArgs e)
		{
			// Изменение состояния checkbox'а "Упрощенно"
			if (chbSimpled.Checked)
				pinObj.PinParam.f = (short)(Convert.ToInt32(pinObj.PinParam.f) | Convert.ToInt32(Pin.SIMPLE));
			else
				pinObj.PinParam.f &= (short)~Pin.SIMPLE;
		}

		private void cbAccClass_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			// Изменение в списке "Класс точности"
			int k = cbAccClass.SelectedIndex;
			int k1 = cbEndLength.SelectedIndex;
			if ((k != -1) && (k1 != -1))
			{
				cbGOST.SelectedIndex = pinObj.GetGostForTypeAndKlass(k1, k);
			}
		}

		private void FrmMain_Load(object sender, System.EventArgs e)
		{
			if (pinObj.PinParam.gost > 22041)
			{
				if (Math.Abs(pinObj.PinParam.b) > 0.001)
					pinObj.PinParam.f &= (short)~ Pin.ALLST;
				else
				{
					int res = Convert.ToInt32(pinObj.PinParam.f) | Convert.ToInt32(Pin.ALLST);
					pinObj.PinParam.f = (short)res;
				}
			}
			else
				pinObj.PinParam.f &= (short)~Pin.ALLST;

			chbShortStep.Checked = (pinObj.PinParam.f & Pin.PITCH) != 0 ? true : false;
			if (pinObj.PinParam.f != 0 & Pin.ISPOLN != 0)
				rbVar1.Checked = true;
			else
				rbVar2.Checked = true;

			FillDiametr(); // заполнение списка диаметра
			FillLenght();  // заполнение списка длин

			bool ew = (pinObj.PinParam.f & Pin.ALLST) != 0 ? false : true;
			rbVar2.Enabled = ew;
			if (ew)
				rbVar1.Checked = true;

			CalcMassa();
			ChoicePitch();
			bitMapId = pinObj.ChoiceBMP(pinObj.PinParam.gost, pinObj.PinParam.f);

			// Заполнение списка ГОСТов
			cbGOST.Items.Clear();
			string[] ids = new string[]{"STR59", "STR60", "STR61", "STR62", "STR63", "STR64", "STR65", "STR66", "STR67", "STR68", "STR69", "STR70"};
			foreach (string id in ids)
				cbGOST.Items.Add(pinObj.LoadString(id));
			cbGOST.SelectedValue = pinObj.LoadString(pinObj.NumberStr(pinObj.PinParam.gost));	// выделяем его в списке

			// Заполнение списка "Длина ввинчиваемого конца l1"
			cbEndLength.Items.Clear();
			ids = new string[]{"IDS_STUD_1D", "IDS_STUD_125D", "IDS_STUD_16D", "IDS_STUD_2D", "IDS_STUD_25D", "IDS_STUD_L0"};
			foreach (string id in ids)
				cbEndLength.Items.Add(pinObj.LoadString(id));
			cbEndLength.SelectedIndex = pinObj.GetTypeForGost(pinObj.PinParam.gost);	// выделение длины ввинчиваемого конца l1

			cbMaterial.Items.Clear();
			ids = new string[]{"STR238", "STR239", "STR240", "STR246"};
			foreach (string id in ids)
				cbMaterial.Items.Add(pinObj.LoadString(id));
			cbMaterial.SelectedIndex = pinObj.PinParam.indexMassa;

			// Заполнение списка "Класс точности"
			cbAccClass.Items.Clear();
			cbAccClass.Items.Add(pinObj.LoadString("IDS_STUD_A"));	// "A"
			cbAccClass.Items.Add(pinObj.LoadString("IDS_STUD_B"));	// "B"
			cbAccClass.SelectedIndex = pinObj.GetKlassForGost(pinObj.PinParam.gost);	// выделение класса точности

			// Изменение состояния checkbox'а "Упрощенно"
			chbSimpled.Checked = (pinObj.PinParam.f & Pin.SIMPLE) != 0 ? true : false;
			chbSpecObject.Checked = pinObj.FlagAttr != 0 ? true: false;

			InitParamList();
		}

		private void btnOK_Click(object sender, System.EventArgs e)
		{
		
		}

	}


	/// <summary>
	/// Структура параметров и класс шпилек
	/// </summary>
	public class BaseMakroParam 
	{
		public float ang;
		public ushort flagAttr;
		public short drawType;
		public byte typeSwitch;	// тип запроса положения базовой точки элемента 0 Placement 1 Curso
								// 0 - точка + направление оси 0X (Placement);
								// 1 - точка, направление совпадает с осью 0X текущей СК (Cursor).
	}


	/// <summary>
	/// Структура параметров
	/// </summary>
	public class PIN 
	{
		public float d;				// диаметр резьбы
		public float p1;			// шаг резьбы
		public float p2;			// шаг резьбы
		public float b1;			// ввинчиваемый конец  l1
		public float c;				// размер фаски
		public float l;				// длина шпильки
		public float b;				// гаечный конец  l0
		public short f;				// битовые маски
		public short klass;			// класс точности
		public short gost;			// номер госта
		public short ver;			// версия макро
		public float m1;			// масса 1 исполнение
		public float m2;			// масса 2 исполнение
		public short indexMassa;	// 0-металл 1- алюмин сплав 3-бронза 2-латунь
	}


	public class PinBase 
	{
		public reference bg, rg1, rg2;
		public reference bs, rs1, rs2;
	}


	/// <summary>
	/// Класс шпилек
	/// </summary>
	public class Pin
	{
		#region Private fields
		private ResourceManager resManager = new ResourceManager(typeof(Pin));

		private bool openBase;
		private ksDocument3D doc;				// интерфейс текущего 3D-документа
		private ksSpecification specification;	// интерфейс спецификации
		private ksPart part;					// деталь
		private ksEntityCollection collect;		// массив всех entity детали
		private bool collectChanged;			// true - массив изменился
		private string fileName;
		private uint countObj;
		private int flagMode;

		private ksUserParam paramSTmp;			// параметры для чтения БД
		private ksUserParam paramTmp;			// параметры для чтения БД
		private ksUserParam param;				// параметры для Get/SetMacroParam
		private ksDataBaseObject data;			// интерфейс БД

		private BaseMakroParam par = new BaseMakroParam();
		private PIN tmp = new PIN();
		private PinBase base_ = new PinBase();	// работа с БД
		#endregion

		#region Constants
		private const short ID_VID = 1;
		private const short ID_SIDEVID = 2;
		private const short ID_TOPVID = 3;
		private const short ID_VIDSEC = 4;

		public static uint PITCH = 0x2;				// крупный шаг >0 мелкий шаг
		public static short ALLST = 0x4;			// резьба на b >0 резьба до головки
		public static short SIMPLE = 0x10;			// нормальная отрисовка >0 упрощенная отрисовка
		public static int ISPOLN = 0x80;			// исполнение 1 >0 исполнение 2
		private static int TAKEISPOLN = 0x100;		// 0 не учитывать исполнение  >0 учитывать исполнение
		private static int PITCHOFF = 0x800;		// есть крупный и мелкий шаг >0 есть только крупный

		private static int COUNT_MASSA = 1000;

		private static int MAX_COUNT_SPCOBJ = 4;	//максимальное число объектов СП за раз 4
		private static int[] spcObj = new int[MAX_COUNT_SPCOBJ];

		private static short STANDART_SECTION = 25;
		private static int SPC_NAME = 5;           // наименование

		private static string STUDS_FILE = "stud.l3d|Шпильки|Шпилька_1";
		#endregion

		#region Resource IDs
		private static int SH22032_1 = 12002;
		private static int SH22042_1 = 12004;
		private static int SH22042_3 = 12005;
		private static int SH22032_2 = 12006;
		private static int SH22042_2 = 12007;
		#endregion

		/// <summary>
		/// Определяет является ли объект поверхностью: planar = true - плоскостью,
		/// false - конической поверхностью
		/// </summary>
		/// <param name="entity"></param>
		/// <param name="planar"></param>
		/// <returns></returns>
		private bool IsSurface(ksEntity entity, bool planar)
		{
			bool res = false;
			if (entity != null) 
			{
				if (entity.IsIt((int)Obj3dType.o3d_face)) 
				{
					ksFaceDefinition faceDef = (ksFaceDefinition)entity.GetDefinition();
					if (faceDef != null) 
					{
						if (planar && faceDef.IsPlanar())
							res = true;

						if (!planar && faceDef.IsCylinder()) 
							res = true;
					}
				}
			}
			return res;
		}


		/// <summary>
		/// Определяет является ли объект осью
		/// </summary>
		/// <param name="?"></param>
		/// <returns></returns>
		private bool IsAxis(ksEntity entity) 
		{
			if (entity != null)
			{
				return entity.IsIt((int)Obj3dType.o3d_axis2Planes) ||	// ось по двум плоскостям
					entity.IsIt((int)Obj3dType.o3d_axisOperation) ||	// ось операций
					entity.IsIt((int)Obj3dType.o3d_axis2Points) ||		// ось по двум точкам
					entity.IsIt((int)Obj3dType.o3d_axisConeFace) ||		// ось конической поверхности
					entity.IsIt((int)Obj3dType.o3d_axisEdge);			// ось по ребру
			}
			return false;
		}


		/// <summary>
		/// Функция фильтр
		/// </summary>
		/// <param name="_entity"></param>
		/// <returns></returns>
		public bool SELECTFILTERPROC(object _entity) 
		{
			bool res = false;  
			if (_entity != null) 
			{                                    
				ksEntity entity = (ksEntity)_entity;
				if (entity != null) 
				{
					ksPart entPart = (ksPart)entity.GetParent();	// Деталь присланного объекта
					if (entPart != Kompas.Instance.PinObj.Part)		// Если деталь другая
						res = IsSurface(entity, false) || 
							IsSurface(entity, false) ||
							IsAxis(entity);							// Проверяем объект на принадлежность к
					// Плоскости, конической поверхности или оси
				}                                               
			}
			return res;
		}


		/// <summary>
		/// Функция обратной связи
		/// 0 - плоскость, 1 - цилиндр, 2 - ось - состав collection
		/// </summary>

		public int SELECTCALLBACKPROC(object _entity, object _info) 
		{
			int res = 0;
			ksEntity entity = (ksEntity)_entity;
			ksRequestInfo3D info = (ksRequestInfo3D)_info;
      ksEntityCollection collection = null;
      if ( info != null )
      {
        collection = (ksEntityCollection) info.GetEntityCollection();
      }
 
			if (entity != null && info != null) 
			{
				if (collection != null) 
				{
					if (IsSurface(entity, false)) 
					{
						// При повторном указании объекта снимаем с него выбор
            ksEntity oldPlane = (ksEntity)collection.GetByIndex( 2 );
            if ( oldPlane == entity )
              collection.SetByIndex( null, 2 );   // Обнулить ось
            else
            {
              collection.SetByIndex( entity, 2 ); // Заменить ось
              collection.SetByIndex( null, 1 );   // Обнулить коническую плоскость
            }
						res = 1;
					}
					else if (IsSurface(entity, false)) 
					{
						// При повторном указании объекта снимаем с него выбор
            ksEntity oldPlane = (ksEntity)collection.GetByIndex( 1 );
            if ( oldPlane == entity )
              collection.SetByIndex( null, 1 );   // Обнулить коническую плоскость
            else
            {
              collection.SetByIndex( entity, 1 ); // Заменить коническую плоскость
              collection.SetByIndex( null, 2 );   // Обнулить ось
            }
						res = 1;
					}
					else if (IsAxis(entity)) 
					{
						// При повторном указании объекта снимаем с него выбор
            ksEntity oldAxis = (ksEntity)collection.GetByIndex(2);
            if ( oldAxis == entity )
              collection.SetByIndex( null, 2 );   // Обнулить ось
            else
            {
              collection.SetByIndex( entity, 2 ); // Заменить ось
              collection.SetByIndex( null,   1 ); // Обнулить коническую плоскость
            }

						res = 1;
					}

					if (res != 0)
						Kompas.Instance.PinObj.collectChanged = true;

					string mes = LoadString("STR_SELECT_PLANE");	// Укажите точку привязки детали
					info.prompt = mes;
				}
			}
			else 
			{
				if (collection != null)
					for (int i = 0; i < 3; i ++)
						collection.SetByIndex(null, i);				// Обнулить

				string mes = LoadString("STR_SELECT_PLANE");	// Укажите точку привязки детали
				info.prompt = mes; 
				res = 1;
			}

			return res;
		}


		/// <summary>
		/// Добавить сопряжения
		/// </summary>
		/// <param name="object_"></param>
		/// <param name="doc"></param>
		/// <param name="collection"></param>
		private void AddMates(Pin object_, ksDocument3D doc, ksEntityCollection collection)
		{
			ksEntity ent = (ksEntity)collection.GetByIndex(0);
			if (ent != null && IsSurface(ent, false))
			{
				// поверхность
				ksEntity partObj = null;
				object_.GetEntityByName("Plane", partObj);
				if (!doc.AddMateConstraint((int)MateConstraintType.mc_Coincidence, ent, partObj,
					object_.DirectionElem, 1, 0))
				{
					collection.SetByIndex(null, 0);
				}
			}

			ent = (ksEntity)collection.GetByIndex(1);
			if (ent != null && IsSurface(ent, false/*planar*/)) 
			{
				// коническая поверхность
				ksEntity partObj = null;
				object_.GetEntityByName("Axis", partObj);
				if (!doc.AddMateConstraint((int)MateConstraintType.mc_Concentric, ent, partObj, 0, 1, 0)) 
				{
					collection.SetByIndex(null, 1);
				}
			}

			ent = (ksEntity)collection.GetByIndex(2);
			if (ent != null && IsAxis(ent)) 
			{
				// ось
				ksEntity partObj = null;
				object_.GetEntityByName("Axis", partObj);
				if (!doc.AddMateConstraint((int)MateConstraintType.mc_Concentric, ent, partObj, 0, 1, 0)) 
				{
					collection.SetByIndex(null, 2);
				}
			}
		}


		/// <summary>
		/// Удалить сопряжения
		/// </summary>
		/// <param name="?"></param>
		private void DelMates(Pin object_, ksDocument3D doc, ksEntityCollection collection) 
		{
			ksEntity ent = (ksEntity)collection.GetByIndex(0);
			if (ent != null && IsSurface(ent, false)) 
			{
				// поверхность
				ksEntity partObj = null;
				object_.GetEntityByName("Plane", partObj);
				doc.RemoveMateConstraint((int)MateConstraintType.mc_Coincidence, ent, partObj);
			}

			ent = (ksEntity)collection.GetByIndex(1);
			if (ent != null && IsSurface(ent, false/*planar*/)) 
			{
				// коническая поверхность
				ksEntity partObj = null;
				object_.GetEntityByName("Axis", partObj);
				doc.RemoveMateConstraint((int)MateConstraintType.mc_Concentric, ent, partObj);
			}

			ent = (ksEntity)collection.GetByIndex(2);
			if (ent != null && IsAxis(ent)) 
			{
				// ось
				ksEntity partObj = null;
				object_.GetEntityByName("Axis", partObj);
				doc.RemoveMateConstraint((int)MateConstraintType.mc_Concentric, ent, partObj);
			}
		}


		/// <summary>
		/// Заполняет адреса функций фильтра и обратной связи
		/// </summary>
		/// <param name="info"></param>
		private void SetCallBackAndFilterFunction(ksRequestInfo3D info) 
		{
			info.SetFilterCallBack("SELECTFILTERPROC", 0, this);
			info.SetCallBack("SELECTCALLBACKPROC", 0, this);
		}


		/// <summary>
		/// Конструктор
		/// </summary>
		public Pin()  
		{
			param = (ksUserParam)Kompas.Instance.KompasObject.GetParamStruct((short)StructType2DEnum.ko_UserParam);
			paramTmp = (ksUserParam)Kompas.Instance.KompasObject.GetParamStruct((short)StructType2DEnum.ko_UserParam);
			paramSTmp = (ksUserParam)Kompas.Instance.KompasObject.GetParamStruct((short)StructType2DEnum.ko_UserParam);
			data = (ksDataBaseObject)Kompas.Instance.KompasObject.DataBaseObject();
			collectChanged = false;
			par.flagAttr = 1;
			par.typeSwitch = 0;
			par.ang = 0;
			par.drawType = ID_VID;
			countObj = 1;

			InitUserParam();
			InitUserParamTmp();
			InitUserParamSTmp();

			if(GetParam() == 0)
				Init();
			else 
			{
				if (tmp.ver == 0)
					tmp.ver = 1;
			}
		}


		public void Init()
		{
			tmp.gost = (short)22032;
			par.drawType = ID_VID;
			par.ang = 0;
			tmp.f = 0;
			tmp.f = Convert.ToInt16(tmp.f != 0 | TAKEISPOLN != 0);
			tmp.d = 20;
			tmp.p1 = 2.5F;
			tmp.p2 = 1.5F;
			tmp.c = 2.5F;
			tmp.indexMassa = 0;
			tmp.ver = 1;
			tmp.b1 = 20;
			tmp.klass = 2; /*klass=B*/
			tmp.l = 90;
			tmp.b = 46;
			tmp.m1 = 245.9F;
			tmp.m2 = 228.9F;
		}


		/// <summary>
		/// Конструктор копии
		/// </summary>
		/// <returns></returns>
		public Pin Dublicate()
		{
			Pin result = new Pin();
			result.Assign(this);
			return result;
		}


		private void Assign(Pin other)
		{
			param = (ksUserParam)Kompas.Instance.KompasObject.GetParamStruct((short)StructType2DEnum.ko_UserParam);
			paramTmp = (ksUserParam)Kompas.Instance.KompasObject.GetParamStruct((short)StructType2DEnum.ko_UserParam);
			paramSTmp = (ksUserParam)Kompas.Instance.KompasObject.GetParamStruct((short)StructType2DEnum.ko_UserParam);
			data = (ksDataBaseObject)Kompas.Instance.KompasObject.DataBaseObject();
			doc = (ksDocument3D)Kompas.Instance.KompasObject.ActiveDocument3D();
			specification = (ksSpecification)doc.GetSpecification();
			openBase = false;
			if (other.part != null)
				part = other.part;

			collectChanged = other.collectChanged; // true - массив изменился
			countObj = other.countObj;
			flagMode = other.flagMode;
			par = other.par;
			fileName = other.fileName;
			specification = other.specification;

			tmp.d = other.tmp.d;
			tmp.p1 = other.tmp.p1;
			tmp.p2 = other.tmp.p2;
			tmp.b1 = other.tmp.b1;
			tmp.c = other.tmp.c;
			tmp.l = other.tmp.l;
			tmp.b = other.tmp.b;
			tmp.f = other.tmp.f;
			tmp.klass = other.tmp.klass;
			tmp.gost = other.tmp.gost;
			tmp.ver = other.tmp.ver;
			tmp.m1 = other.tmp.m1;
			tmp.m2 = other.tmp.m2;
			tmp.indexMassa = other.tmp.indexMassa;

			base_.bg = other.base_.bg;
			base_.rg1 = other.base_.rg1;
			base_.rg2 = other.base_.rg2;
			base_.bs = other.base_.bs;
			base_.rs1 = other.base_.rs1;
			base_.rs2 = other.base_.rs2;

			par.ang = other.par.ang;
			par.flagAttr = other.par.flagAttr;
			par.drawType = other.par.drawType;
			par.typeSwitch = other.par.typeSwitch;

			InitUserParam();
			InitUserParamTmp();
			InitUserParamSTmp();
		}

		public void _MessageBoxResult() 
		{
			if (Kompas.Instance.KompasObject.ksReturnResult() != 0)
			{
				if (Kompas.Instance.KompasObject.ksReturnResult() == (int)ErrorType.etError10)	//  10  "Ошибка! Вырожденный объект"
					Kompas.Instance.KompasObject.ksResultNULL();
				else
					Kompas.Instance.KompasObject.ksMessageBoxResult();
			}
		}

		public void InitUserParamTmp() 
		{
			if (paramTmp != null) 
			{
				ksLtVariant item = (ksLtVariant)Kompas.Instance.KompasObject.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
				ksDynamicArray arr = (ksDynamicArray)Kompas.Instance.KompasObject.GetDynamicArray(ldefin2d.LTVARIANT_ARR);
				if (item != null && arr != null) 
				{
					paramTmp.Init();
					paramTmp.SetUserArray(arr);
					item.Init();
					item.shortVal = 0;            // 0 - t
					arr.ksAddArrayItem(-1, item);
					item.floatVal = tmp.d;        // 1 - d
					arr.ksAddArrayItem(-1, item);
					item.floatVal = tmp.p1;       // 2 - p1
					arr.ksAddArrayItem(-1, item);
					item.floatVal = tmp.p2;       // 3 - p2
					arr.ksAddArrayItem(-1, item);
					item.floatVal = tmp.b1;       // 4 - b1
					arr.ksAddArrayItem(-1, item);
					item.floatVal = tmp.c;        // 5 - c
					arr.ksAddArrayItem(-1, item);
				}
			}
		}

		public void GetUserParamTmp()
		{
			if (paramTmp != null) 
			{
				ksLtVariant item = (ksLtVariant)Kompas.Instance.KompasObject.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
				ksLtVariant item1 = (ksLtVariant)Kompas.Instance.KompasObject.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
				ksDynamicArray arr = (ksDynamicArray)paramTmp.GetUserArray();
				if (item != null && item1 != null && arr != null 
					&& arr.ksGetArrayCount() >= 6) 
				{
					item.Init();
					arr.ksGetArrayItem(1, item); // d
					tmp.d = item.floatVal;
					arr.ksGetArrayItem(2, item); // p1
					tmp.p1 = item.floatVal;
					arr.ksGetArrayItem(3, item); // p2
					tmp.p2 = item.floatVal;
					arr.ksGetArrayItem(4, item); // b1
					tmp.b1 = item.floatVal;
					arr.ksGetArrayItem(5, item); // c
					tmp.c = item.floatVal;
				}
			}
		}

		public void InitUserParamSTmp() 
		{
			if (paramSTmp != null) 
			{
				ksLtVariant item = (ksLtVariant)Kompas.Instance.KompasObject.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
				ksDynamicArray arr = (ksDynamicArray)Kompas.Instance.KompasObject.GetDynamicArray(ldefin2d.LTVARIANT_ARR);
				if (item != null && arr != null) 
				{
					paramSTmp.Init();
					paramSTmp.SetUserArray(arr);
					item.Init();
					item.shortVal = 0;            // 0 - t
					arr.ksAddArrayItem(-1, item);
					item.floatVal = tmp.l;        // 1 - l
					arr.ksAddArrayItem(-1, item);
					item.floatVal = tmp.d;        // 2 - d
					arr.ksAddArrayItem(-1, item);
					item.floatVal = tmp.b;        // 3 - b
					arr.ksAddArrayItem(-1, item);
					item.floatVal = tmp.m1;       // 4 - m1
					arr.ksAddArrayItem(-1, item);
					item.floatVal = tmp.m2;       // 5 - m2
					arr.ksAddArrayItem(-1, item);
				}
			}
		}

		public void GetUserParamSTmp() 
		{
			if (paramSTmp != null) 
			{
				ksLtVariant item = (ksLtVariant)Kompas.Instance.KompasObject.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
				ksDynamicArray arr = (ksDynamicArray)paramSTmp.GetUserArray();
				if (item != null && arr != null 
					&& arr.ksGetArrayCount() >= 6) 
				{
					item.Init();
					arr.ksGetArrayItem(1, item); // l
					tmp.l = item.floatVal;
					arr.ksGetArrayItem(2, item); // d
					tmp.d = item.floatVal;
					arr.ksGetArrayItem(3, item); // s
					tmp.b = item.floatVal;
					arr.ksGetArrayItem(4, item); // m1
					tmp.m1 = item.floatVal;
					arr.ksGetArrayItem(5, item); // m2
					tmp.m2 = item.floatVal;
				}
				if (tmp.gost > 22041) 
				{
					if (Math.Abs(tmp.b) > 0.001) 
						tmp.f &= (short)~ALLST;
					else
					{
						int res = Convert.ToInt32(tmp.f) | Convert.ToInt32(ALLST);
						tmp.f = (short)res;
					}
				}
			}
		}

		public void NumberGost(int gost, ref string gNumb, ref string stNumb)
		{
			switch (gost)
			{
				case 22032 : gNumb = "STR_200"; stNumb = "STR_220"; break;
				case 22033 : gNumb = "STR_201"; stNumb = "STR_221"; break;
				case 22034 : gNumb = "STR_202"; stNumb = "STR_222"; break;
				case 22035 : gNumb = "STR_203"; stNumb = "STR_223"; break;
				case 22036 : gNumb = "STR_204"; stNumb = "STR_224"; break;
				case 22037 : gNumb = "STR_205"; stNumb = "STR_225"; break;
				case 22038 : gNumb = "STR_206"; stNumb = "STR_226"; break;
				case 22039 : gNumb = "STR_207"; stNumb = "STR_227"; break;
				case 22040 : gNumb = "STR_208"; stNumb = "STR_228"; break;
				case 22041 : gNumb = "STR_209"; stNumb = "STR_229"; break;
				case 22042 : gNumb = "STR_210"; stNumb = "STR_230"; break;
				case 22043 : gNumb = "STR_211"; stNumb = "STR_231"; break;
			}
		}

		public bool GetFullName(string name, ref string fName)
		{
			string bufName = name;			// буфер из которого будем все доставать
			string path = string.Empty;		// полный путь к файлу
			string modelName = string.Empty;// разделы и подразделы в библ-ке моделей
			bool rez = false;				// результат
			fName = Assembly.GetExecutingAssembly().Location;
			// путь к нашей dll
			int ind = bufName.IndexOf('|');	// ищем позицию с кот. нач-ся разделы          
			if (ind > 0)					// если раздел найден - копируем путь внутри библ-ки в буфер
				modelName = bufName.Substring(ind, bufName.Length - ind);

			string fileName = ind > 0 ? bufName.Substring(0, ind) : bufName; // имя файла без разделов
			if (fileName != string.Empty) 
			{
				if (fName != null && fName != string.Empty) 
				{
					FileInfo rootInfo = new FileInfo(fName);
					path = rootInfo.DirectoryName + @"\" + fileName;
					FileInfo fInfo = new FileInfo(path);
					if (fInfo.Exists)	// проверка файла
						rez = true;		// файл найден
					else 
					{
						// файл не найден
						path = rootInfo.DirectoryName + @"\LOAD\" + fileName;
						fInfo = new FileInfo(path);
						if (fInfo.Exists)	// проверка файла
							rez = true;		// файл найден
					}
				}
			}
			if (rez)						// файл найден
				fName = path + modelName;	// формируем полный путь

			return rez;
		}

		public int _ConnectDB(reference bd, ref string name)
		{
			string buf = string.Empty;

			return GetFullName(name, ref buf) ? Data.ksConnectDB(bd, buf) 
				: Data.ksConnectDB(bd, name);
		}


		/// <summary>
		/// Открытие БД и создание отношений к БД
		/// </summary>
		/// <param name="initParam"></param>
		/// <returns></returns>
		public bool OpenBase(bool initParam)
		{
			if (initParam) 
			{
				InitUserParam();
				InitUserParamTmp();
				InitUserParamSTmp();
			}

			if (data != null) 
			{
				string buf;
				buf = LoadString("BASE_STR1");
				Base.bg = Data.ksCreateDB(buf);   // TXT_DB
				Base.bs = Data.ksCreateDB(buf);   // TXT_DB

				string gNumb = string.Empty, stNumb = string.Empty;
				NumberGost(tmp.gost, ref gNumb, ref stNumb);

				buf = LoadString(gNumb);
				if(_ConnectDB(Base.bg, ref buf) != 0)
				{
					//"22032.loa"
					buf = LoadString(stNumb);
					if(_ConnectDB(Base.bs, ref buf) != 0)
					{
						//"st22032.loa"
						buf = LoadString("STR_241");
						Base.rg1 = Data.ksRelation(Base.bg);
						Data.ksRInt ("t");
						Data.ksRChar("d", 20, 0);
						Data.ksEndRelation();
						if (Data.ksDoStatement(Base.bg, Base.rg1, buf) == 1) 
						{
							//"1"
							Base.rg2 = Data.ksRelation(Base.bg);
							Data.ksRInt  ("t");
							Data.ksRFloat("d");
							Data.ksRFloat("");
							Data.ksRFloat("");
							Data.ksRFloat("");
							Data.ksRFloat("");
							Data.ksEndRelation();
							if (Data.ksDoStatement(Base.bg, Base.rg2, "") == 1) 
							{
								//TXT_ALL
								buf = LoadString("STR_243");
								Base.rs1 = Data.ksRelation(Base.bs);
								Data.ksRInt  ("t");
								Data.ksRChar ("L", 20, 0);
								Data.ksRFloat("d");
								Data.ksEndRelation();
								if (Data.ksDoStatement(Base.bs, Base.rs1, buf) == 1) 
								{
									//"1 2"
									Base.rs2 = Data.ksRelation(Base.bs);
									Data.ksRInt  ("t");
									Data.ksRFloat("L");
									Data.ksRFloat("d");
									Data.ksRFloat("");
									Data.ksRFloat("");
									Data.ksRFloat("");
									Data.ksEndRelation();
									if (Data.ksDoStatement(Base.bs, Base.rs2, "") == 1) 
									{
										// TXT_ALL
										openBase = true;
										return true;
									}
								}
							}
						}
					}
				}
			}
			return false;
		}


		/// <summary>
		/// Закрытие БД
		/// </summary>
		public void CloseBase() 
		{
			if (Base.bg != 0)
			{
				Data.ksDisconnectDB(Base.bg);
				Data.ksDeleteDB(Base.bg);
				Base.bg = 0;
			}
			if (Base.bs != 0) 
			{
				Data.ksDisconnectDB(Base.bs);
				Data.ksDeleteDB(Base.bs);
				Base.bs = 0;
			}
			openBase = false;
		}


		/// <summary>
		/// Чтение БД
		/// </summary>
		/// <param name="d"></param>
		/// <returns></returns>
		public int ReadPinBase(float d) 
		{
			string  buf;
			string s = string.Empty;
			buf = LoadString("STR_245");
			s = string.Format(buf, d);
			if (Data.ksCondition(Base.bg, Base.rg2, s) == 1) 
			{ 
				if (Data.ksReadRecord(Base.bg, Base.rg2, paramTmp) == 1) 
				{
					GetUserParamTmp();
					return 1;
				}
			}
			return 0;
		}


		/// <summary>
		/// Читать параметры стержня
		/// возвращает 1 - успех 0 - не найдено записи -1 - ошибка связи с БД
		/// </summary>
		/// <returns></returns>
		public int ReadPinStBase() 
		{
			string buf = string.Empty;
			string buf1 = string.Empty;
			buf = LoadString("STR_246");
			buf1 = string.Format(buf, tmp.l, tmp.d);

			if (Data.ksCondition(Base.bs, Base.rs2, buf1) == 0) { return -1; }
			int i = Data.ksReadRecord(Base.bs, Base.rs2, paramSTmp);
			if (i == 0)  
			{
				buf = LoadString("STR_245");
				buf1 = string.Format(buf, tmp.d);

				if (Data.ksCondition(Base.bs,  Base.rs2, buf1) == 0) { return -1;}
				Data.ksReadRecord(Base.bs, Base.rs2, paramSTmp);
			}
			GetUserParamSTmp();
			return 1;
		}


		public void InitUserParam() 
		{
			if (param != null) 
			{
				ksLtVariant item = (ksLtVariant)Kompas.Instance.KompasObject.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
				ksDynamicArray arr = (ksDynamicArray)Kompas.Instance.KompasObject.GetDynamicArray(ldefin2d.LTVARIANT_ARR);
				if (item != null && arr != null) 
				{
					param.Init();
					param.SetUserArray(arr);
					item.Init();
					item.floatVal = par.ang;        // 0 - ang
					arr.ksAddArrayItem(-1, item);
					item.shortVal = (short)par.flagAttr;   // 1 - flagAttr
					arr.ksAddArrayItem(-1, item);
					item.shortVal = par.drawType;   // 2 - drawType
					arr.ksAddArrayItem(-1, item);
					item.shortVal = par.typeSwitch; // 3 - typeSwitch
					arr.ksAddArrayItem(-1, item);

					item.floatVal = tmp.d;          // 4 - d
					arr.ksAddArrayItem(-1, item);
					item.floatVal = tmp.p1;         // 5 - p1
					arr.ksAddArrayItem(-1, item);
					item.floatVal = tmp.p2;         // 6 - p2
					arr.ksAddArrayItem(-1, item);
					item.floatVal = tmp.b1;         // 7 - b1
					arr.ksAddArrayItem(-1, item);
					item.floatVal = tmp.c;          // 8 - c
					arr.ksAddArrayItem(-1, item);
					item.floatVal = tmp.l;          // 9 - l
					arr.ksAddArrayItem(-1, item);
					item.floatVal = tmp.b;          // 10 - b
					arr.ksAddArrayItem(-1, item);
					item.shortVal = tmp.f;          // 11 - f
					arr.ksAddArrayItem(-1, item);
					item.shortVal = tmp.klass;      // 12 - klass
					arr.ksAddArrayItem(-1, item);
					item.shortVal = tmp.gost;       // 13 - gost
					arr.ksAddArrayItem(-1, item);
					item.shortVal = tmp.ver;        // 14 - ver
					arr.ksAddArrayItem(-1, item);
					item.floatVal = tmp.m1;         // 15 - m1
					arr.ksAddArrayItem(-1, item);
					item.floatVal = tmp.m2;         // 16 - m2
					arr.ksAddArrayItem(-1, item);
					item.shortVal = tmp.indexMassa; // 17 - indexMassa
					arr.ksAddArrayItem(-1, item);
				}
			}
		}

		public void SetUserParam() 
		{
			if (param != null) 
			{
				ksLtVariant item = (ksLtVariant)Kompas.Instance.KompasObject.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
				ksDynamicArray arr = (ksDynamicArray)param.GetUserArray();
				if (item != null && arr != null && arr.ksGetArrayCount() >= 18) 
				{
					item.Init();
					item.floatVal = par.ang;        // 0 - ang
					arr.ksSetArrayItem(0, item);
					item.shortVal = (short)par.flagAttr;   // 1 - flagAttr
					arr.ksSetArrayItem(1, item);
					item.shortVal = par.drawType;   // 2 - drawType
					arr.ksSetArrayItem(2, item);
					item.shortVal = par.typeSwitch; // 3 - typeSwitch
					arr.ksSetArrayItem(3, item);

					item.floatVal = tmp.d;          // 4 - d
					arr.ksSetArrayItem(4, item);
					item.floatVal = tmp.p1;         // 5 - p1
					arr.ksSetArrayItem(5, item);
					item.floatVal = tmp.p2;         // 6 - p2
					arr.ksSetArrayItem(6, item);
					item.floatVal = tmp.b1;         // 7 - b1
					arr.ksSetArrayItem(7, item);
					item.floatVal = tmp.c;          // 8 - c
					arr.ksSetArrayItem(8, item);
					item.floatVal = tmp.l;          // 9 - l
					arr.ksSetArrayItem(9, item);
					item.floatVal = tmp.b;          // 10 - b
					arr.ksSetArrayItem(10, item);
					item.shortVal = tmp.f;          // 11 - f
					arr.ksSetArrayItem(11, item);
					item.shortVal = tmp.klass;      // 12 - klass
					arr.ksSetArrayItem(12, item);
					item.shortVal = tmp.gost;       // 13 - gost
					arr.ksSetArrayItem(13, item);
					item.shortVal = tmp.ver;        // 14 - ver
					arr.ksSetArrayItem(14, item);
					item.floatVal = tmp.m1;         // 15 - m1
					arr.ksSetArrayItem(15, item);
					item.floatVal = tmp.m2;         // 16 - m2
					arr.ksSetArrayItem(16, item);
					item.shortVal = tmp.indexMassa; // 17 - indexMassa
					arr.ksSetArrayItem(17, item);
				}
			}
		}

		public void GetUserParam() 
		{
			if (param != null) 
			{
				ksLtVariant item = (ksLtVariant)Kompas.Instance.KompasObject.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
				ksDynamicArray arr = (ksDynamicArray)param.GetUserArray();
				if (item != null && arr != null && arr.ksGetArrayCount() >= 18) 
				{
					item.Init();
					arr.ksGetArrayItem(0, item);
					par.ang = item.floatVal;
					arr.ksGetArrayItem(1, item);
					par.flagAttr = (ushort)item.shortVal;
					arr.ksGetArrayItem(2, item);
					par.drawType = item.shortVal;
					arr.ksGetArrayItem(3, item);
					par.typeSwitch = (byte)item.shortVal;

					arr.ksGetArrayItem(4, item);
					tmp.d = item.floatVal;
					arr.ksGetArrayItem(5, item);
					tmp.p1 = item.floatVal;
					arr.ksGetArrayItem(6, item);
					tmp.p2 = item.floatVal;
					arr.ksGetArrayItem(7, item);
					tmp.b1 = item.floatVal;
					arr.ksGetArrayItem(8, item);
					tmp.c = item.floatVal;
					arr.ksGetArrayItem(9, item);
					tmp.l = item.floatVal;
					arr.ksGetArrayItem(10, item);
					tmp.b = item.floatVal;
					arr.ksGetArrayItem(11, item);
					tmp.f = item.shortVal;
					arr.ksGetArrayItem(12, item);
					tmp.klass = item.shortVal;
					arr.ksGetArrayItem(13, item);
					tmp.gost = item.shortVal;
					arr.ksGetArrayItem(14, item);
					tmp.ver = item.shortVal;
					arr.ksGetArrayItem(15, item);
					tmp.m1 = item.floatVal;
					arr.ksGetArrayItem(16, item);
					tmp.m2 = item.floatVal;
					arr.ksGetArrayItem(17, item);
					tmp.indexMassa = item.shortVal;
				}
			}
		}

		public int GetParam() 
		{
			if (part != null) 
			{
				part.GetUserParam(param);
				GetUserParam();
				return 1;
			}
			return 0; 
		}


		/// <summary>
		/// Находит у детали поверхность с заданным именем
		/// </summary>
		/// <param name="name"></param>
		/// <param name="entity"></param>
		public void GetEntityByName(string name, ksEntity entity) 
		{
			entity = null;
			if (part != null) 
			{
				if (collect == null)  // массив лежит в классе, и если он ещё не создан - создаём
					collect = (ksEntityCollection)part.EntityCollection((int)Obj3dType.o3d_constrElement);

				ksEntity findEntity = (ksEntity)collect.GetByName(name, true, true);
				if (findEntity != null) 
					entity = findEntity;
			}
		}


		/// <summary>
		/// Массив пуст - указали placement
		/// </summary>
		/// <param name="info"></param>
		/// <returns></returns>
		public bool IsCollectionEmpty(ksRequestInfo3D info) 
		{
			bool res = true;
			ksEntityCollection collection = (ksEntityCollection)info.GetEntityCollection(); // Новый массив
			if (collection != null)
				for(int i = 0, count = collection.GetCount(); i < count; i++) // Цикл по массиву
				{       
					ksEntity ent = (ksEntity)collection.GetByIndex(i);        // Текущий элемент
					if (ent != null)                          // Если он не null
					{                        
						res = false;                                   // Массив не пуст
						break;                                         // Завершить цикл
					}
				}
			return res; 
		}


		public bool IsCollectionChange(ksRequestInfo3D info) 
		{
			bool res = true;                                     
			if (!collectChanged) // Если новые объекты не указывали - сравниваем массивы
			{                             
				// Возможен вариант когда указали точку и массив в info заполнен null'ми а старый нет
				res = false; // Не надо перестраивать
				ksEntityCollection collection1 = (ksEntityCollection)info.GetEntityCollection(); // Новый массив
				ksEntityCollection collection2 = (ksEntityCollection)part.GetMateConstraintObjects(); // Старый массив
				if (collection1 != null && collection2 != null && 
					(collection1.GetCount() > 2) && (collection2.GetCount() > 2)) 
				{
					for(int i = 0; i < 3; i ++) // Цикл по массивам
					{                        
						ksEntity ent1 = (ksEntity)collection1.GetByIndex(i);	// Текущий элемент 1
						ksEntity ent2 = (ksEntity)collection2.GetByIndex(i);	// Текущий элемент 2
						if (ent1 != null && ent2 != null && ent1 != ent2)		// Если они не равны 
						{     
							res = true;                                 // Перестраиваем
							break;                                      // Завершить цикл
						}
					}
				}
			}
			return res; 
		}

		public void SetMates(ksDocument3D doc, ksRequestInfo3D info, bool delMates) 
		{
			ksEntityCollection collection = (ksEntityCollection)(!delMates 
				? info.GetEntityCollection() : part.GetMateConstraintObjects());
			if (collection != null) 
			{
				if (doc != null && collection.GetCount() > 2)
					if (!delMates)
						AddMates(this, doc, collection);
					else
						DelMates(this, doc, collection);

				if (!delMates)
					part.SetMateConstraintObjects(collection);
			}
		}

		public reference EditSpcObj(reference spcObj)
		{
			ksUserParam par1 = (ksUserParam)Kompas.Instance.KompasObject.GetParamStruct((short)StructType2DEnum.ko_UserParam);
			ksLtVariant item = (ksLtVariant)Kompas.Instance.KompasObject.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
			ksDynamicArray arr = (ksDynamicArray)Kompas.Instance.KompasObject.GetDynamicArray(ldefin2d.LTVARIANT_ARR);
			if (par1 == null || item == null || arr == null || specification == null)
				return 0;

			par1.Init();
			par1.SetUserArray(arr);
			item.Init();

			if (flagMode == 0 && par.flagAttr == 0 && specification == null)
				return 0;

			if (flagMode != 0)
			{
				spcObj = specification.D3GetSpcObjForGeomWithLimit(null,	// Имя библиотеки типов
					0,														// Номер типа спецификации
					part,
					1,														// 1 - первый объект 0 - следующий объект                                              		
					STANDART_SECTION,										// Номер раздела
					81764182256.0D);

				if(par.flagAttr == 0)
					return spcObj;

				if (spcObj != 0 && specification.ksSpcObjectEdit(spcObj) == 0)
					spcObj = 0;
			}

			bool flagMode_ = (spcObj != 0 ? true : false);
			if (flagMode_ | specification.ksSpcObjectCreate(null,	// Имя библиотеки типов
				0,													// Номер типа спецификации
				STANDART_SECTION, 0,								// Номер раздела и подраздела
				81764182256.0D, 0) != 0)							// Тип атрибута
			{
				// Исполнение
				uint uBuf = Convert.ToUInt32(!!(tmp.f != 0 & ISPOLN != 0));
				specification.ksSpcVisible(SPC_NAME, 2, (short)uBuf);
				if (uBuf != 0)
				{
					arr.ksClearArray();
					item.uIntVal = 2;
					arr.ksAddArrayItem(-1, item);
					specification.ksSpcChangeValue(SPC_NAME, 2, par1, ldefin2d.UINT_ATTR_TYPE);
				}

				// Изменим диаметр
				arr.ksClearArray();
				item.floatVal = tmp.d;
				arr.ksAddArrayItem(-1, item);
				specification.ksSpcChangeValue(SPC_NAME, 4, par1, ldefin2d.FLOAT_ATTR_TYPE);

				// Отследим мелкий шаг
				uBuf = Convert.ToUInt32(!(tmp.f != 0 & PITCHOFF != 0 || !(tmp.f != 0 & PITCH != 0)));
				specification.ksSpcVisible(SPC_NAME, 5, (short)uBuf);
				specification.ksSpcVisible(SPC_NAME, 6, (short)uBuf); // Шаг
				if(uBuf != 0)	// Выключить шаг и его разделитель
				{
					arr.ksClearArray();
					item.floatVal = tmp.p2;
					arr.ksAddArrayItem(-1, item);
					specification.ksSpcChangeValue(SPC_NAME, 6, par1, ldefin2d.FLOAT_ATTR_TYPE);
				}

				// Изменим длину
				arr.ksClearArray();
				item.uIntVal = (int)tmp.l;
				arr.ksAddArrayItem(-1, item);
				specification.ksSpcChangeValue(SPC_NAME , 9, par1, ldefin2d.UINT_ATTR_TYPE);

				if (flagMode == 0)
				{
					specification.ksSpcVisible(SPC_NAME, 7,  0); // Выключим поле допуска
					specification.ksSpcVisible(SPC_NAME, 10, 0); // Выключим класс прочности
					specification.ksSpcVisible(SPC_NAME, 11, 0); // Выключим материал
					specification.ksSpcVisible(SPC_NAME, 12, 0); // Выключим покрытие
				}

				// Изменим ГОСТ
				arr.ksClearArray();
				item.uIntVal = tmp.gost;
				arr.ksAddArrayItem(-1, item);
				specification.ksSpcChangeValue(SPC_NAME , 14, par1, ldefin2d.UINT_ATTR_TYPE);
				float massa  = tmp.f != 0 & ISPOLN != 0 ? tmp.m2 : tmp.m1;
				massa = (float)(massa * (tmp.indexMassa == 0
					? 1 : tmp.indexMassa == 1 
					? 0.356 : tmp.indexMassa == 3 
					? 0.97 : 1.08) / COUNT_MASSA);
				string buf = string.Empty;
				buf = string.Format("{0:0.#}", massa);
				specification.ksSpcMassa(buf);	// Масса детали

				// Подключим геометрию
				specification.D3SpcIncludePart(part, false);

				return specification.ksSpcObjectEnd();
			}
			return 0;
		}

		public bool DrawSpcObj() 
		{
			//memset(spcObj, 0, sizeof(reference) * MAX_COUNT_SPCOBJ);
			if (Kompas.Instance.KompasObject.ksReturnResult() == (int)ErrorType.etError10) // Вырожденный объект
				Kompas.Instance.KompasObject.ksResultNULL();

			spcObj[0] = Kompas.Instance.PinObj.EditSpcObj(spcObj[0]);
			if(Kompas.Instance.PinObj.par.flagAttr == 0 && spcObj[0] != 0) 
			{
				for (int i = 0, count = Kompas.Instance.PinObj.ObjCount; i < count; i++) 
				{
					if (spcObj[i] != 0) 
					{
						Kompas.Instance.PinObj.doc.ksDeleteObj(spcObj[i]);
						spcObj[i] = 0;
					}
				}
			}
			return spcObj[0] != 0 ? true : false;
		}


		/// <summary>
		/// Процесс создания детали
		/// </summary>
		/// <param name="_doc"></param>
		public void Draw3D(ksDocument3D _doc)
		{
			if (_doc != null) 
			{
				doc = _doc;
				specification = (ksSpecification)doc.GetSpecification();
				flagMode = Convert.ToInt32(doc.IsEditMode());
				Kompas.Instance.PinObj = this;

				// Взяли деталь
				part = (ksPart)doc.GetPart(flagMode != 0 ? (int)Part_Type.pEdit_Part : (int)Part_Type.pNew_Part);
				if (part != null) 
				{
					// После определения детали нужно считать параметры редактирования
					if (flagMode != 0)
						GetParam();
					else
						part.fileName = SetFileName();

					if (Kompas.Instance.PinObj.MacroElementParam()) 
					{
						if (flagMode != 0)
						{ 
							// Режим редактирования без изменения местоположения и mate'ов
							SetParam(part);		// Установим параметры
						}
						else
						{   
							// Создание нового
							ksRequestInfo3D info = (ksRequestInfo3D)doc.GetRequestInfo(part);
							if (info != null) 
							{
								string prompt;
					 
								prompt = LoadString(flagMode != 0 ? "STR_SELECT_PLANE" : "STR_SELECT_POINT");	// Укажите точку привязки детали
								info.prompt = prompt;
								SetCallBackAndFilterFunction(info);
								info.CreatePhantom();
								ksPart phantom = (ksPart)info.GetIPhantom();
								if (phantom != null)
									SetParam(phantom); 
            
								// Теперь временные српряжения
								if (doc.UserGetPlacementAndEntity(3)) 
								{
									ksPlacement lpace = (ksPlacement)info.GetPlacement();
									part.SetPlacement(lpace);
									if (flagMode != 0 || doc.SetPartFromFile(Kompas.Instance.PinObj.SetFileName(), part, false/*externalFile*/))
									{
										SetParam(part);

										if (IsCollectionChange(info))			// Массив Mate'ов изменился
										{ 
											if (flagMode != 0)					// Режим редактирования
												SetMates(doc, info, true);		// Удалить сопряжения

											SetMates(doc, info, false);			// Создать сопряжения
										}

										if (IsCollectionEmpty(info))			// Указали placement
											part.UpdatePlacement();				// Изменм положение детали
                
										// Устанавливаем признак, стандартное изделие
										part.standardComponent = true;
									}
									DrawSpcObj();		// Создать объект СП

									if (spcObj[0] != 0)	// Вывод объекта СП
									{                       
										// Олегово окно - редактируем параметры
										for (int i = 0, count = Kompas.Instance.PinObj.ObjCount; i < count; i ++) //{
											specification.ksEditWindowSpcObject(spcObj[i]);
										spcObj[0] = 0;
										spcObj[1] = 0;
										spcObj[2] = 0;
										spcObj[3] = 0;
									}
								}
							}
						}
					}
				}
			}
		}


		/// <summary>
		/// Присваивает переменной с именем varName значение val
		/// </summary>
		/// <param name="varArr"></param>
		/// <param name="varName"></param>
		/// <param name="val"></param>
		public void SetVarValue(ksVariableCollection varArr, string varName, double val)
		{
			ksVariable var = (ksVariable)varArr.GetByName(varName, true, false); // текущая переменная
			if (var != null)
				var.value = val;                     // сменить значение
		}


		public void SetParam(ksPart _part) 
		{
			// Редактируем внешние переменные
			if (_part != null) 
			{
				// Массив внешних переменных детали
				ksVariableCollection varArr = (ksVariableCollection)_part.VariableCollection();
				if (varArr != null) 
				{
					float l, b1 = 0, b; // b1 = 0 - условие резьбы на всю длину
					l = tmp.b1 + tmp.l; // Длина всей шпильки
					b = tmp.b;
					float step = (tmp.f & PITCH) != 0 ? tmp.p1 : tmp.p2;
					if(Math.Abs(tmp.b1) < 0.001) // ГОСТ 22042, ГОСТ 22043 
					{ 
						if(Math.Abs(tmp.b) > 0.001) 
							b1 = b;
					}
					else 
					{
						b1 = tmp.b1;
						if (Math.Abs(tmp.b) < 0.001)
							b = (float)(tmp.l - 0.5 * tmp.d - 2 * step);
					}
					// Основные параметры
					SetVarValue(varArr, "d", tmp.d);					// Диаметр резьбы
					float len = Math.Abs(b1 - b) < 0.001 ? tmp.l - b1 : tmp.l;
					SetVarValue(varArr, "l", len);						// Длина шпильки
					SetVarValue(varArr, "b1", b1);						// Длина ввинчиваемого конца
					SetVarValue(varArr, "c1", tmp.c);					// Размер фаски на конце стержня
					if (Math.Abs(l - b1 - b) > 0.001)
						SetVarValue(varArr, "l1",     l - b1 - b);		// Длина гладкой части стержня

					SetVarValue(varArr, "t", (step * 0.866 * 3 / 8));	// Толщина удаляемого материала

					int perf2 = ((tmp.f & ISPOLN) == 0 || (tmp.f & ALLST) != 0 || Math.Abs(l - b1 - b) < 0.001) ? 1 : 0;
					SetVarValue(varArr, "cut1", perf2);					// Вкл/выкл исполнение 2
					SetVarValue(varArr, "f1", (tmp.f & SIMPLE) == 0 ? 0.0 : 1.0); // Вкл/выкл подголовок

					// Задать название детали
					string buf1;
					buf1 = LoadString("IDS_STUD_NAME");					// Заготовка имени детали из ресурса
					string buf = string.Empty;
					buf = string.Format(buf1, tmp.d, tmp.l, tmp.gost);	// Создать имя детали
					_part.name = buf;									// Присвоить детали новое имя
					_part.Update();										// Обновить деталь

					// Перестроим модель с измененными внешними переменными
					_part.RebuildModel();
				}

				SetUserParam();
				_part.SetUserParam(param);
			}
		}


		/// <summary>
		/// Выбор меню
		/// </summary>
		/// <returns></returns>
		public string ChoiceMenuId() 
		{
			string result = string.Empty;

			switch (par.drawType) 
			{
				case ID_VID    : // Вид
					result = par.typeSwitch == 0 ? "MENU_22032_1" : "MENU_22032_3";
					break;
				case ID_TOPVID : // Вид сверху
					result = par.typeSwitch == 0 ? "MENU_22032_2" : "MENU_22032_4";
					break;
			}

			return result;
		}


		public string SetFileName()
		{
			string name = string.Empty;
			GetFullName(STUDS_FILE, ref name);
			fileName = name;
			return fileName;
		}

		public string LoadString(string name)
		{
			return resManager.GetString(name);
		}

		public bool MacroElementParam()
		{
			bool res = false;
			Pin bufS = Kompas.Instance.PinObj.Dublicate();
			if (bufS != null && bufS.OpenBase(true))
			{
				Kompas.Instance.KompasObject.ksEnableTaskAccess(0);	// закрыть доступ к компасу

				try
				{
					if (FrmMain.Instance.Execute(bufS) == DialogResult.OK) 
						res = true;
				}
				catch (Exception ex) {Kompas.Instance.KompasObject.ksMessage(ex.Message);}

				Kompas.Instance.KompasObject.ksEnableTaskAccess(1);	// открыть доступ к компасу

				if (res) 
					Kompas.Instance.PinObj.Assign(bufS);
				bufS.CloseBase();
			}
			return res;
		}


		/// <summary>
		/// Выбор слайда
		/// </summary>
		/// <param name="gost"></param>
		/// <param name="f"></param>
		/// <returns></returns>
		public int ChoiceBMP(short gost, short f) 
		{
			int bmp = 0;
			if (gost > 22041) 
				bmp = f != 0 & ALLST != 0 ? SH22042_3 : f != 0 & ISPOLN != 0 ? SH22042_2 : SH22042_1;
			else 
				bmp = f != 0 & ISPOLN != 0 ? SH22032_2 : SH22032_1;
			return bmp;
		}


		/// <summary>
		/// Определяет длину ввинчиваемого конца по ГОСТу
		/// </summary>
		/// <param name="gost"></param>
		/// <returns></returns>
		public int GetTypeForGost(short gost) 
		{
			int n = 0;
			switch (gost) 
			{
				case 22032 :
				case 22033 : n = 0;  break; // 1d
				case 22034 :
				case 22035 : n = 1;  break; // 1,25d
				case 22036 :
				case 22037 : n = 2;  break; // 1,6d
				case 22038 :
				case 22039 : n = 3;  break; // 2d
				case 22040 :
				case 22041 : n = 4;  break; // 2,5d
				case 22042 :
				case 22043 : n = 5;  break; // l0
			}
			return n;
		}


		/// <summary>
		/// Определяет класс точности по ГОСТу
		/// </summary>
		/// <param name="gost"></param>
		/// <returns></returns>
		public int GetKlassForGost(short gost) 
		{
			int n = 0;
			switch (gost) 
			{
				case 22032 :
				case 22034 :
				case 22036 :
				case 22038 :
				case 22040 :
				case 22042 :  n = 1;  break;	// B

				case 22033 :
				case 22035 :
				case 22037 :
				case 22039 :
				case 22041 :
				case 22043 :  n = 0;  break;	// A
			}
			return n;
		}


		/// <summary>
		/// Определяет класс точности по ГОСТу
		/// </summary>
		/// <param name="type"></param>
		/// <param name="klass"></param>
		/// <returns></returns>
		public int GetGostForTypeAndKlass(int type, int klass) 
		{
			int n = 0;
			switch (type) 
			{
				case 0 : n = klass != 0 ? 0 : 1;	break; // 22032 : 22033
				case 1 : n = klass != 0 ? 2 : 3;	break; // 22034 : 22035
				case 2 : n = klass != 0 ? 4 : 5;	break; // 22036 : 22037
				case 3 : n = klass != 0 ? 6 : 7;	break; // 22038 : 22039
				case 4 : n = klass != 0 ? 8 : 9;	break; // 22040 : 22041
				case 5 : n = klass != 0 ? 10 : 11;	break; // 22042 : 22043
			}
			return n;
		}


		public string NumberStr(int gost) 
		{
			string n = string.Empty;
			switch (gost) 
			{
				case 22032 : n = "STR59"; break;
				case 22033 : n = "STR60"; break;
				case 22034 : n = "STR61"; break;
				case 22035 : n = "STR62"; break;
				case 22036 : n = "STR63"; break;
				case 22037 : n = "STR64"; break;
				case 22038 : n = "STR65"; break;
				case 22039 : n = "STR66"; break;
				case 22040 : n = "STR67"; break;
				case 22041 : n = "STR68"; break;
				case 22042 : n = "STR69"; break;
				case 22043 : n = "STR70"; break;
			}
			return n;
		}

		public short DirectionElem
		{
			get
			{
				return 1;
			}
		}

		public ksPart Part
		{
			get
			{
				return part;
			}
		}

		public bool CollectChanged
		{
			get
			{
				return collectChanged;
			}
		}

		public int ObjCount
		{
			get
			{
				return (int)countObj;
			}
		}

		public PIN PinParam
		{
			get
			{
				return tmp;
			}
		}
		
		public PinBase Base
		{
			get
			{
				return base_;
			}
		}

		public ksDataBaseObject Data
		{
			get
			{
				return data;
			}
		}

		public bool BaseOpened
		{
			get
			{
				return openBase;
			}
		}

		public BaseMakroParam MacroParam
		{
			get
			{
				return par;
			}
		}

		public ushort FlagAttr
		{
			get
			{
				return par.flagAttr;
			}
		}
	}
}