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
	/// Form1 �̊T�v�̐����ł��B
	/// </summary>
	public class Form1 : System.Windows.Forms.Form
	{
		private System.ComponentModel.IContainer components;

		public Form1()
		{
			//
			// Windows �t�H�[�� �f�U�C�i �T�|�[�g�ɕK�v�ł��B
			//
			InitializeComponent();

			//
			// TODO: InitializeComponent �Ăяo���̌�ɁA�R���X�g���N�^ �R�[�h��ǉ����Ă��������B
			//
		}

		/// <summary>
		/// �g�p����Ă��郊�\�[�X�Ɍ㏈�������s���܂��B
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

		#region Windows �t�H�[�� �f�U�C�i�Ő������ꂽ�R�[�h 
		/// <summary>
		/// �f�U�C�i �T�|�[�g�ɕK�v�ȃ��\�b�h�ł��B���̃��\�b�h�̓��e��
		/// �R�[�h �G�f�B�^�ŕύX���Ȃ��ł��������B
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
		/// �A�v���P�[�V�����̃��C�� �G���g�� �|�C���g�ł��B
		/// </summary>
		[STAThread]
		static void Main() 
		{
//			Application.Run(new Form1());
			// �J�����g�f�B���N�g���̎擾
			string s�J�����g�f�B���N�g�� = System.IO.Directory.GetCurrentDirectory();
			string s�f�X�N�g�b�v = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

			// �V���[�g�J�b�g���쐬
			ShellLink shortcut;

			shortcut = new ShellLink();
//			shortcut.Description = "";
			shortcut.TargetPath = Path.Combine(s�J�����g�f�B���N�g��,"AutoUpGrade.exe");
			shortcut.WorkingDirectory = s�J�����g�f�B���N�g��;
//			shortcut.Arguments = "";
			shortcut.IconFile = Path.Combine(s�J�����g�f�B���N�g��,"App\\IS2Client.exe");
//			shortcut.IconFile = Path.Combine(s�J�����g�f�B���N�g��,"Is2CreateIcon.dll");
//			shortcut.IconIndex = 1;
			shortcut.Save(Path.Combine(s�f�X�N�g�b�v,"is-2.lnk"));
			shortcut.Dispose();
			shortcut = null;

			shortcut = new ShellLink();
//			shortcut.Description = "";
			shortcut.TargetPath = Path.Combine(s�J�����g�f�B���N�g��,"AutoUpGrade.exe");
			shortcut.WorkingDirectory = s�J�����g�f�B���N�g��;
			shortcut.Arguments = "hinagata";
//			shortcut.IconFile = Path.Combine(s�J�����g�f�B���N�g��,"Is2CreateIcon.exe");
			shortcut.IconFile = Path.Combine(s�J�����g�f�B���N�g��,"Is2CreateIcon.dll");
//			shortcut.IconIndex = 0;
			shortcut.Save(Path.Combine(s�f�X�N�g�b�v,"�N�C�b�N�G���g���[.lnk"));
			shortcut.Dispose();
			shortcut = null;

		}
	}
}
