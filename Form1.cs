using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;

namespace Is2CreateIcon
{
	/// <summary>
	/// Form1 の概要の説明です。
	/// </summary>
	public class Form1 : System.Windows.Forms.Form
	{
		private System.ComponentModel.IContainer components;

		public Form1()
		{
			//
			// Windows フォーム デザイナ サポートに必要です。
			//
			InitializeComponent();

			//
			// TODO: InitializeComponent 呼び出しの後に、コンストラクタ コードを追加してください。
			//
		}

		/// <summary>
		/// 使用されているリソースに後処理を実行します。
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows フォーム デザイナで生成されたコード 
		/// <summary>
		/// デザイナ サポートに必要なメソッドです。このメソッドの内容を
		/// コード エディタで変更しないでください。
		/// </summary>
		private void InitializeComponent()
		{
			// 
			// Form1
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 12);
			this.ClientSize = new System.Drawing.Size(292, 266);
			this.Name = "Form1";
			this.Text = "Form1";

		}
		#endregion

		/// <summary>
		/// アプリケーションのメイン エントリ ポイントです。
		/// </summary>
		[STAThread]
		static void Main() 
		{
//			Application.Run(new Form1());
			// カレントディレクトリの取得
			string sカレントディレクトリ = System.IO.Directory.GetCurrentDirectory();
			string sデスクトップ = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

			// ショートカットを作成
			ShellLink shortcut;

			shortcut = new ShellLink();
//			shortcut.Description = "";
			shortcut.TargetPath = Path.Combine(sカレントディレクトリ,"AutoUpGrade.exe");
			shortcut.WorkingDirectory = sカレントディレクトリ;
//			shortcut.Arguments = "";
			shortcut.IconFile = Path.Combine(sカレントディレクトリ,"App\\IS2Client.exe");
//			shortcut.IconFile = Path.Combine(sカレントディレクトリ,"Is2CreateIcon.dll");
//			shortcut.IconIndex = 1;
			shortcut.Save(Path.Combine(sデスクトップ,"is-2.lnk"));
			shortcut.Dispose();
			shortcut = null;

			shortcut = new ShellLink();
//			shortcut.Description = "";
			shortcut.TargetPath = Path.Combine(sカレントディレクトリ,"AutoUpGrade.exe");
			shortcut.WorkingDirectory = sカレントディレクトリ;
			shortcut.Arguments = "hinagata";
//			shortcut.IconFile = Path.Combine(sカレントディレクトリ,"Is2CreateIcon.exe");
			shortcut.IconFile = Path.Combine(sカレントディレクトリ,"Is2CreateIcon.dll");
//			shortcut.IconIndex = 0;
			shortcut.Save(Path.Combine(sデスクトップ,"クイックエントリー.lnk"));
			shortcut.Dispose();
			shortcut = null;

		}
	}
}
