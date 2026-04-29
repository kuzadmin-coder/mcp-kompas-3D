using Kompas6API5;

using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Steps.NET
{
	public class FrmTest : System.Windows.Forms.Form
	{
		#region Designer declarations
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private System.Windows.Forms.Button btnCancel;
		private System.Windows.Forms.Button btnOK;
		#endregion

		private string slideFileName;
		private System.Windows.Forms.Panel panel;
	
		public string SlideFileName
		{
			set
			{
				slideFileName = value;
			}
		}


		private KompasObject kompas;
		public KompasObject Kompas
		{
			set
			{
				kompas = value;
			}
		}


		#region Instance etc...
		private FrmTest()
		{
			InitializeComponent();

			this.Paint += new PaintEventHandler(FrmTest_Paint);
		}


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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(FrmTest));
			this.btnCancel = new System.Windows.Forms.Button();
			this.btnOK = new System.Windows.Forms.Button();
			this.panel = new System.Windows.Forms.Panel();
			this.SuspendLayout();
			// 
			// btnCancel
			// 
			this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btnCancel.Location = new System.Drawing.Point(216, 240);
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.Size = new System.Drawing.Size(64, 23);
			this.btnCancel.TabIndex = 0;
			this.btnCancel.Text = "Отмена";
			// 
			// btnOK
			// 
			this.btnOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
			this.btnOK.Location = new System.Drawing.Point(144, 240);
			this.btnOK.Name = "btnOK";
			this.btnOK.Size = new System.Drawing.Size(64, 23);
			this.btnOK.TabIndex = 1;
			this.btnOK.Text = "ОК";
			// 
			// panel
			// 
			this.panel.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.panel.Location = new System.Drawing.Point(8, 8);
			this.panel.Name = "panel";
			this.panel.Size = new System.Drawing.Size(272, 224);
			this.panel.TabIndex = 3;
			// 
			// FrmTest
			// 
			this.AcceptButton = this.btnOK;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.CancelButton = this.btnCancel;
			this.ClientSize = new System.Drawing.Size(290, 273);
			this.Controls.Add(this.panel);
			this.Controls.Add(this.btnOK);
			this.Controls.Add(this.btnCancel);
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "FrmTest";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "Тест слайд";
			this.ResumeLayout(false);

		}
		#endregion
		#endregion

		private void FrmTest_Paint(object sender, PaintEventArgs e)
		{
			if (kompas != null)
			{
				if (slideFileName != null && slideFileName != string.Empty)
				{
					kompas.ksDrawSlideFromFile(this.panel.Handle.ToInt32(), slideFileName);				
				}
			}
		}
	}
}
