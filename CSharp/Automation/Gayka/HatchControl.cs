using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace Steps.NET
{
	public delegate void AngleChangedHandler();
	public delegate void StepChangedHandler();

	public class HatchControl : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.TextBox tbAngle;
		private System.Windows.Forms.TextBox tbStep;
		private System.ComponentModel.Container components = null;

		public event AngleChangedHandler AngleChanged;
		public event StepChangedHandler StepChanged;

		public string AngleEditValue
		{
			get
			{
				return tbAngle.Text;
			}
			set
			{
				tbAngle.Text = value;
			}
		}

		public string StepEditValue
		{
			get
			{
				return tbStep.Text;
			}
			set
			{
				tbStep.Text = value;
			}
		}


		public HatchControl()
		{
			InitializeComponent();
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

		#region Component Designer generated code
		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.tbAngle = new System.Windows.Forms.TextBox();
			this.tbStep = new System.Windows.Forms.TextBox();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(8, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(88, 19);
			this.label1.TabIndex = 0;
			this.label1.Text = "”гол штриховки";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(8, 56);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(88, 19);
			this.label2.TabIndex = 1;
			this.label2.Text = "Ўаг штриховки";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// tbAngle
			// 
			this.tbAngle.Location = new System.Drawing.Point(8, 32);
			this.tbAngle.Name = "tbAngle";
			this.tbAngle.Size = new System.Drawing.Size(88, 20);
			this.tbAngle.TabIndex = 2;
			this.tbAngle.Text = "45";
			this.tbAngle.TextChanged += new System.EventHandler(this.tbAngle_TextChanged);
			// 
			// tbStep
			// 
			this.tbStep.Location = new System.Drawing.Point(8, 80);
			this.tbStep.Name = "tbStep";
			this.tbStep.Size = new System.Drawing.Size(88, 20);
			this.tbStep.TabIndex = 3;
			this.tbStep.Text = "1";
			this.tbStep.TextChanged += new System.EventHandler(this.tbStep_TextChanged);
			// 
			// HatchControl
			// 
			this.Controls.Add(this.tbStep);
			this.Controls.Add(this.tbAngle);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Name = "HatchControl";
			this.Size = new System.Drawing.Size(104, 112);
			this.ResumeLayout(false);

		}
		#endregion

		private void tbAngle_TextChanged(object sender, System.EventArgs e)
		{
			if (AngleChanged != null)
				AngleChanged();
		}

		private void tbStep_TextChanged(object sender, System.EventArgs e)
		{
			if (StepChanged != null)
				StepChanged();
		}
	}
}
