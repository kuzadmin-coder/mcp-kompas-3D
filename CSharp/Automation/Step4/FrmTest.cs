using Kompas6API5;
using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

using reference = System.Int32;

namespace Steps.NET
{
	public class FrmTest : System.Windows.Forms.Form
	{
		#region Designer declarations
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		#endregion

		#region Instance etc...
		private FrmTest()
		{
			InitializeComponent();

		}

		private System.Windows.Forms.Button btnCancel;
		private System.Windows.Forms.Button btnOK;
		private System.Windows.Forms.Panel panel;

		private static FrmTest instance;
		public static FrmTest Instance
		{
			get
			{
				if (instance == null)
					instance = new FrmTest();
				return instance;

			}
		}

		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.btnCancel = new System.Windows.Forms.Button();
			this.btnOK = new System.Windows.Forms.Button();
			this.panel = new System.Windows.Forms.Panel();
			this.SuspendLayout();
			// 
			// btnCancel
			// 
			this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btnCancel.Location = new System.Drawing.Point(222, 240);
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.Size = new System.Drawing.Size(64, 23);
			this.btnCancel.TabIndex = 0;
			this.btnCancel.Text = "Отмена";
			// 
			// btnOK
			// 
			this.btnOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
			this.btnOK.Location = new System.Drawing.Point(150, 240);
			this.btnOK.Name = "btnOK";
			this.btnOK.Size = new System.Drawing.Size(64, 23);
			this.btnOK.TabIndex = 1;
			this.btnOK.Text = "ОК";
			// 
			// panel
			// 
			this.panel.Location = new System.Drawing.Point(8, 8);
			this.panel.Name = "panel";
			this.panel.Size = new System.Drawing.Size(280, 224);
			this.panel.TabIndex = 2;
			// 
			// FrmTest
			// 
			this.AcceptButton = this.btnOK;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.CancelButton = this.btnCancel;
			this.ClientSize = new System.Drawing.Size(298, 273);
			this.Controls.Add(this.panel);
			this.Controls.Add(this.btnOK);
			this.Controls.Add(this.btnCancel);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "FrmTest";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Отрисовка слайда";
			this.Paint += new System.Windows.Forms.PaintEventHandler(this.FrmTest_Paint);
			this.ResumeLayout(false);

		}
		#endregion
		#endregion

		private ksDocument2D doc;
		public ksDocument2D Doc
		{
			get
			{
				return doc;
			}
			set
			{
				doc = value;
			}
		}


		private void FrmTest_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{
			if (doc != null && doc.reference != 0)
			{
				reference gr = doc.ksNewGroup(1);
				doc.ksMtr(20, 15, 45, 1, 1);
				doc.ksLineSeg(-10, 0, 10, 0, 1);
				doc.ksLineSeg(10, 0, 10, 20, 1);
				doc.ksLineSeg(10, 20, -10, 20, 1);
				doc.ksLineSeg(-10, 20, -10, 0, 1);
				doc.ksDeleteMtr();
				doc.ksEndGroup();
				doc.ksDrawKompasGroup(panel.Handle.ToInt32(), gr);
        doc.ksDeleteObj( gr );
			}
		}
	}
}
