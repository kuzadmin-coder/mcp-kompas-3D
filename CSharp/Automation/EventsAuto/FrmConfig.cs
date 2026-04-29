using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Reflection;
using System.Xml;
using System.IO;

namespace Steps.NET
{
	public class FrmConfig : System.Windows.Forms.Form
	{
		#region Designer declarations
		private System.ComponentModel.Container components = null;
		private System.Windows.Forms.Button btnOK;
		private System.Windows.Forms.Button btnCancel;
		public System.Windows.Forms.CheckBox chbAutoAdvise;
		private System.Windows.Forms.GroupBox group;
		public System.Windows.Forms.CheckBox chbAppEvents;
		public System.Windows.Forms.CheckBox chbDocEvents;
		public System.Windows.Forms.CheckBox chb2DObjEvents;
		public System.Windows.Forms.CheckBox chb3DObjEvents;
		public System.Windows.Forms.CheckBox chbStampEvents;
		public System.Windows.Forms.CheckBox chbSelectEvents;
		public System.Windows.Forms.CheckBox chbSpecObjEvents;
		public System.Windows.Forms.CheckBox chbSpecDocEvents;
		public System.Windows.Forms.CheckBox chb3DDocEvents;
		public System.Windows.Forms.CheckBox chb2DDocEvents;
		public System.Windows.Forms.CheckBox chbSpecEvents;
		public System.Windows.Forms.CheckBox chbDocFrameEvents;
		#endregion

		#region Custom declarations
		private XmlDocument xmlDoc = new XmlDocument();
		#endregion

		#region Instance etc...
		private FrmConfig()
		{
			InitializeComponent();

			// Читаем настройки из файла
			FileInfo fInfo = new FileInfo(Assembly.GetExecutingAssembly().Location + ".config");
			if (fInfo.Exists)
			{
				xmlDoc.Load(fInfo.FullName);
				XmlElement rootElem = xmlDoc.DocumentElement;

				foreach (Control c in this.Controls)
					ProcessControl(c, rootElem);

				foreach (Control c in group.Controls)
					ProcessControl(c, rootElem);
			}
		}

		private static FrmConfig instance;
		public static FrmConfig Instance
		{
			get
			{
				if (instance == null)
					instance = new FrmConfig();
				return instance;

			}
		}

		protected override void Dispose(bool disposing)
		{
			if (disposing)
			{
				if (components != null)
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
			this.chbAutoAdvise = new System.Windows.Forms.CheckBox();
			this.group = new System.Windows.Forms.GroupBox();
			this.chbSpecObjEvents = new System.Windows.Forms.CheckBox();
			this.chbSpecDocEvents = new System.Windows.Forms.CheckBox();
			this.chb3DDocEvents = new System.Windows.Forms.CheckBox();
			this.chb2DDocEvents = new System.Windows.Forms.CheckBox();
			this.chbSpecEvents = new System.Windows.Forms.CheckBox();
			this.chb3DObjEvents = new System.Windows.Forms.CheckBox();
			this.chbStampEvents = new System.Windows.Forms.CheckBox();
			this.chbSelectEvents = new System.Windows.Forms.CheckBox();
			this.chb2DObjEvents = new System.Windows.Forms.CheckBox();
			this.chbDocEvents = new System.Windows.Forms.CheckBox();
			this.chbAppEvents = new System.Windows.Forms.CheckBox();
			this.btnOK = new System.Windows.Forms.Button();
			this.btnCancel = new System.Windows.Forms.Button();
			this.chbDocFrameEvents = new System.Windows.Forms.CheckBox();
			this.group.SuspendLayout();
			this.SuspendLayout();
			// 
			// chbAutoAdvise
			// 
			this.chbAutoAdvise.Location = new System.Drawing.Point(16, 3);
			this.chbAutoAdvise.Name = "chbAutoAdvise";
			this.chbAutoAdvise.Size = new System.Drawing.Size(328, 24);
			this.chbAutoAdvise.TabIndex = 0;
			this.chbAutoAdvise.Text = "Автоподписка при загрузке";
			// 
			// group
			// 
			this.group.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.group.Controls.Add(this.chbDocFrameEvents);
			this.group.Controls.Add(this.chbSpecObjEvents);
			this.group.Controls.Add(this.chbSpecDocEvents);
			this.group.Controls.Add(this.chb3DDocEvents);
			this.group.Controls.Add(this.chb2DDocEvents);
			this.group.Controls.Add(this.chbSpecEvents);
			this.group.Controls.Add(this.chb3DObjEvents);
			this.group.Controls.Add(this.chbStampEvents);
			this.group.Controls.Add(this.chbSelectEvents);
			this.group.Controls.Add(this.chb2DObjEvents);
			this.group.Controls.Add(this.chbDocEvents);
			this.group.Controls.Add(this.chbAppEvents);
			this.group.Location = new System.Drawing.Point(8, 24);
			this.group.Name = "group";
			this.group.Size = new System.Drawing.Size(370, 298);
			this.group.TabIndex = 1;
			this.group.TabStop = false;
			// 
			// chbSpecObjEvents
			// 
			this.chbSpecObjEvents.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.chbSpecObjEvents.Location = new System.Drawing.Point(8, 248);
			this.chbSpecObjEvents.Name = "chbSpecObjEvents";
			this.chbSpecObjEvents.Size = new System.Drawing.Size(354, 24);
			this.chbSpecObjEvents.TabIndex = 11;
			this.chbSpecObjEvents.Text = "Выдавать сообщения для событий объекта спецификации ";
			// 
			// chbSpecDocEvents
			// 
			this.chbSpecDocEvents.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.chbSpecDocEvents.Location = new System.Drawing.Point(8, 224);
			this.chbSpecDocEvents.Name = "chbSpecDocEvents";
			this.chbSpecDocEvents.Size = new System.Drawing.Size(354, 24);
			this.chbSpecDocEvents.TabIndex = 10;
			this.chbSpecDocEvents.Text = "Выдавать сообщения для событий документа спецификации";
			// 
			// chb3DDocEvents
			// 
			this.chb3DDocEvents.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.chb3DDocEvents.Location = new System.Drawing.Point(8, 200);
			this.chb3DDocEvents.Name = "chb3DDocEvents";
			this.chb3DDocEvents.Size = new System.Drawing.Size(354, 24);
			this.chb3DDocEvents.TabIndex = 9;
			this.chb3DDocEvents.Text = "Выдавать сообщения для событий 3D документа";
			// 
			// chb2DDocEvents
			// 
			this.chb2DDocEvents.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.chb2DDocEvents.Location = new System.Drawing.Point(8, 176);
			this.chb2DDocEvents.Name = "chb2DDocEvents";
			this.chb2DDocEvents.Size = new System.Drawing.Size(354, 24);
			this.chb2DDocEvents.TabIndex = 8;
			this.chb2DDocEvents.Text = "Выдавать сообщения для событий 2D документа";
			// 
			// chbSpecEvents
			// 
			this.chbSpecEvents.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.chbSpecEvents.Location = new System.Drawing.Point(8, 152);
			this.chbSpecEvents.Name = "chbSpecEvents";
			this.chbSpecEvents.Size = new System.Drawing.Size(354, 24);
			this.chbSpecEvents.TabIndex = 7;
			this.chbSpecEvents.Text = "Выдавать сообщения для событий спецификации";
			// 
			// chb3DObjEvents
			// 
			this.chb3DObjEvents.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.chb3DObjEvents.Location = new System.Drawing.Point(8, 128);
			this.chb3DObjEvents.Name = "chb3DObjEvents";
			this.chb3DObjEvents.Size = new System.Drawing.Size(354, 24);
			this.chb3DObjEvents.TabIndex = 6;
			this.chb3DObjEvents.Text = "Выдавать сообщения для событий объекта 3D документа";
			// 
			// chbStampEvents
			// 
			this.chbStampEvents.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.chbStampEvents.Location = new System.Drawing.Point(8, 104);
			this.chbStampEvents.Name = "chbStampEvents";
			this.chbStampEvents.Size = new System.Drawing.Size(354, 24);
			this.chbStampEvents.TabIndex = 5;
			this.chbStampEvents.Text = "Выдавать сообщения для событий штампа";
			// 
			// chbSelectEvents
			// 
			this.chbSelectEvents.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.chbSelectEvents.Location = new System.Drawing.Point(8, 80);
			this.chbSelectEvents.Name = "chbSelectEvents";
			this.chbSelectEvents.Size = new System.Drawing.Size(354, 24);
			this.chbSelectEvents.TabIndex = 4;
			this.chbSelectEvents.Text = "Выдавать сообщения для событий селектирования";
			// 
			// chb2DObjEvents
			// 
			this.chb2DObjEvents.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.chb2DObjEvents.Location = new System.Drawing.Point(8, 56);
			this.chb2DObjEvents.Name = "chb2DObjEvents";
			this.chb2DObjEvents.Size = new System.Drawing.Size(354, 24);
			this.chb2DObjEvents.TabIndex = 3;
			this.chb2DObjEvents.Text = "Выдавать сообщения для событий объекта 2D документа";
			// 
			// chbDocEvents
			// 
			this.chbDocEvents.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.chbDocEvents.Location = new System.Drawing.Point(8, 32);
			this.chbDocEvents.Name = "chbDocEvents";
			this.chbDocEvents.Size = new System.Drawing.Size(354, 24);
			this.chbDocEvents.TabIndex = 2;
			this.chbDocEvents.Text = "Выдавать сообщения для событий документов";
			// 
			// chbAppEvents
			// 
			this.chbAppEvents.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.chbAppEvents.Location = new System.Drawing.Point(8, 8);
			this.chbAppEvents.Name = "chbAppEvents";
			this.chbAppEvents.Size = new System.Drawing.Size(354, 24);
			this.chbAppEvents.TabIndex = 1;
			this.chbAppEvents.Text = "Выдавать сообщения для событий Компаса";
			// 
			// btnOK
			// 
			this.btnOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
			this.btnOK.Location = new System.Drawing.Point(314, 330);
			this.btnOK.Name = "btnOK";
			this.btnOK.Size = new System.Drawing.Size(64, 23);
			this.btnOK.TabIndex = 2;
			this.btnOK.Text = "ОК";
			this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
			// 
			// btnCancel
			// 
			this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btnCancel.Location = new System.Drawing.Point(242, 330);
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.Size = new System.Drawing.Size(64, 23);
			this.btnCancel.TabIndex = 3;
			this.btnCancel.Text = "Отмена";
			// 
			// chbDocFrameEvents
			// 
			this.chbDocFrameEvents.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.chbDocFrameEvents.Location = new System.Drawing.Point(8, 272);
			this.chbDocFrameEvents.Name = "chbDocFrameEvents";
			this.chbDocFrameEvents.Size = new System.Drawing.Size(354, 24);
			this.chbDocFrameEvents.TabIndex = 12;
			this.chbDocFrameEvents.Text = "Выдавать сообщения для событий объекта спецификации ";
			// 
			// FrmConfig
			// 
			this.AcceptButton = this.btnOK;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.CancelButton = this.btnCancel;
			this.ClientSize = new System.Drawing.Size(386, 359);
			this.Controls.Add(this.btnCancel);
			this.Controls.Add(this.btnOK);
			this.Controls.Add(this.group);
			this.Controls.Add(this.chbAutoAdvise);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "FrmConfig";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Конфигурация";
			this.Load += new System.EventHandler(this.FrmConfig_Load);
			this.group.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion
		#endregion

		private void AddNode(Control c, XmlNode xmlroot)
		{
			if (c is CheckBox)
			{
				CheckBox chb = c as CheckBox;
				XmlNode elem = xmlDoc.CreateNode(XmlNodeType.Element, chb.Name, string.Empty);
				XmlNode elemText = xmlDoc.CreateNode(XmlNodeType.Text, string.Empty, string.Empty);
				elemText.Value = chb.Checked.ToString();
				elem.AppendChild(elemText);
				xmlroot.AppendChild(elem);
			}
		}

		private void btnOK_Click(object sender, System.EventArgs e)
		{
			// Сохраняем настройки в файл
			xmlDoc = new XmlDocument();
			XmlNode xmlroot = xmlDoc.CreateNode(XmlNodeType.Element, "EventsAutoConfig", string.Empty);
			xmlDoc.AppendChild(xmlroot);

			foreach (Control c in this.Controls)
				AddNode(c, xmlroot);

			foreach (Control c in group.Controls)
				AddNode(c, xmlroot);

			xmlDoc.Save(Assembly.GetExecutingAssembly().Location + ".config");
		}

		private void ProcessControl(Control c, XmlElement elem)
		{
			if (c is CheckBox)
			{
				CheckBox chb = c as CheckBox;
				string nodeValue = elem.SelectSingleNode(c.Name).InnerText;
				if (nodeValue != null && nodeValue != string.Empty)
					chb.Checked = bool.Parse(nodeValue);
			}
		}

		private void FrmConfig_Load(object sender, System.EventArgs e)
		{
		
		}
	}
}
