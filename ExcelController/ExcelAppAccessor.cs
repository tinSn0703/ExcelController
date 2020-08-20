using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace ExcelController
{
	/// <summary>
	/// Excelのアプリケーションをポンポン開くとおいしくないから、一度だけにしたい。
	/// とりあえず作った。まだよくわからないので試行していく。
	/// このオブジェクトがExcelのアプリへのアクセスを担当する.
	/// </summary>
	class ExcelAppAccessor : IDisposable
	{
		public ExcelAppAccessor()
		{
			_ReferenceCounter += 1;
		}

		/// <summary>
		/// 既存のBookを開く
		/// </summary>
		/// <param name="_FilePath">開きたいファイルのパス</param>
		/// <returns>開いたブック</returns>
		public ExcelBookAccessor Open(string _FilePath)
		{
			if ( ! File.Exists(_FilePath)) throw new FileNotFoundException(_FilePath);

			try
			{
				if (this.IsFileOpened(_FilePath)) return this.BindWorkbook(_FilePath);
			}
			catch (Exception e)
			{
				if (e.Message.IndexOf(_FilePath) < 0) throw e;
				Console.WriteLine(e);
			}

			this.SecureApp();

			return new ExcelBookAccessor(_ExcelBooks.Open(_FilePath));
		}
		
		/// <summary>
		/// 新しいBookを開く
		/// </summary>
		/// <param name="_FilePath">新しいファイルの保存先のパス。既に存在するならそちらが開かれる</param>
		/// <returns>追加したブック</returns>
		public ExcelBookAccessor Add(string _FilePath)
		{
//			if (File.Exists(_FilePath)) return this.Open(_FilePath);

			Excel.Workbook _ExcelBook = null;
			
			try
			{
				this.SecureApp();

				_ExcelBook = _ExcelBooks.Add();
				_ExcelBook.SaveAs(_FilePath);

				return new ExcelBookAccessor(_ExcelBook);
			}
			catch (Exception e)
			{
				this.ReleaseObject(_ExcelBook);
				throw e;
			}
		}

		/// <summary>
		/// アプリを閉じる
		/// </summary>
		public void Close()
		{
			if (_ExcelApp != null) _ExcelApp.Quit();
		}

		/// <summary>
		/// 表示する
		/// </summary>
		/// <param name="_IsVisuble"></param>
		public void Visible(bool _IsVisuble)
		{
			_ExcelApp.Visible = _IsVisuble;
		}
		
		public Excel.Range Union(Range Arg1, Range Arg2)
		{
			if ((Arg1 == null) || (Arg2 == null)) throw new ArgumentNullException(nameof(Arg1) + " or " + nameof(Arg2));

			return _ExcelApp.Union(Arg1, Arg2);
		}

		/// <summary>
		/// 指定したファイルは、既に開かれていますか?
		/// </summary>
		/// <param name="_FilePath">調べるファイルのパス</param>
		/// <returns></returns>
		private bool IsFileOpened(string _FilePath)
		{
			string _FileName = System.IO.Path.GetFileName(_FilePath); //ファイル名を取り出す
			foreach (Process _Process in Process.GetProcesses())
			{
				//関係ないプロセス。スキップ
				if (_Process.MainWindowTitle.Length == 0) continue;

				//現在開かれているプロセス名と比較し、ファイルが開かれているか確認する
				if (_Process.MainWindowTitle.IndexOf(_FileName) >= 0) return true;
			}

			return false;
		}
		
		/// <summary>
		/// 正しい拡張子ですか?
		/// </summary>
		/// <param name="_FilePath">調べるファイルのパス</param>
		/// <returns></returns>
		private bool IsExtensionCorrect(string _FilePath)
		{
			return false;
		}

		/// <summary>
		/// 実行中のExcelブックを取得する
		/// </summary>
		/// <param name="_FilePath">実行中ブックのパス</param>
		/// <returns>実行中のブック</returns>
		private ExcelBookAccessor BindWorkbook(string _FilePath)
		{
			Excel.Workbook _ExcelBook = null;

			try
			{
				_ExcelBook = Marshal.BindToMoniker(_FilePath) as Excel.Workbook;

				if (_ExcelBook == null) throw new System.Exception(_FilePath + "\n確保に失敗");

				this.SecureApp(_ExcelBook.Application);

				return new ExcelBookAccessor(_ExcelBook);
			}
			catch (System.Exception e)
			{
				this.ReleaseObject(_ExcelBook);
				this.ReleaseApplication();
				throw e;
			}
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="_App"></param>
		private void SecureApp(Excel.Application _App = null)
		{
			if (_ExcelApp != null) return;

			this.ReleaseApplication();

			if (_App == null)	_ExcelApp = new Excel.Application();
			else				_ExcelApp = _App;

			_ExcelBooks = _ExcelApp.Workbooks;
		}

		/// <summary>
		/// Objectを開放する
		/// </summary>
		private void ReleaseObject(object _Obj)
		{
			if (_Obj != null)
			{
				while (Marshal.ReleaseComObject(_Obj) > 0) ;
				_Obj = null;
			}
		}

		/// <summary>
		/// Application Objectを開放する
		/// </summary>
		private void ReleaseApplication()
		{
			// Books解放
			if (_ExcelBooks != null)
			{
				while (System.Runtime.InteropServices.Marshal.ReleaseComObject(_ExcelBooks) > 0);
				_ExcelBooks = null;
			}

			// Excelアプリケーションを解放
			if (_ExcelApp != null)
			{
				while (System.Runtime.InteropServices.Marshal.ReleaseComObject(_ExcelApp) > 0);
				_ExcelApp = null;
			}
		}

		/// <summary>
		/// 確保されたリソースの解放
		/// </summary>
		/// <param name="_Disposing">GCが解放してくれるリソースを開放するかしないか</param>
		protected virtual void Dispose(bool _Disposing)
		{
			if (!_DisposeValue)
			{
				if (_Disposing)	{}

				if (_ReferenceCounter < 2)
				{
					this.ReleaseApplication();
				}
				else
				{
					_ReferenceCounter -= 1;
				}

				_DisposeValue = true;
			}
		}

		/// <summary>
		/// 確保されたリソースの解放
		/// </summary>
		public void Dispose()
		{
			this.Dispose(true);
			GC.SuppressFinalize(this);
		}

		/// <summary>
		/// デストラクタではなくファイナライザ。
		/// C++と構文が同じだから勘違いしてたぞ
		/// </summary>
		~ExcelAppAccessor()
		{
			this.Dispose(false);
		}

		private bool _DisposeValue = false;

		static private int _ReferenceCounter = 0;
		static private Excel.Application _ExcelApp = null;
		static private Excel.Workbooks _ExcelBooks = null;
	}
}
