using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelController
{
	/// <summary>
	/// Excel book へのアクセスを担当する。要するにBookのWrappar
	/// </summary>
	class ExcelBookAccessor : IDisposable
    {
		public ExcelBookAccessor()
		{
			this._ExcelBook = null;
			this._ExcelSheets = null;
		}

		public ExcelBookAccessor(Excel.Workbook _ExcelBook)
		{
			this.SetBook(_ExcelBook);
		}

		/// <summary>
		/// アクセス先のExcelブックを設定する。
		/// </summary>
		/// <param name="_ExcelBook">アクセス先のブック</param>
		/// <exception cref="ArgumentNullException">アクセス先のブックがnullだった場合</exception>
		public void SetBook(Excel.Workbook _ExcelBook)
		{
			this.ReleaseBook();

			this._ExcelBook = _ExcelBook ?? throw new ArgumentNullException(nameof(_ExcelBook));
			this._ExcelSheets = this._ExcelBook.Sheets;
		}

		/// <summary>
		/// Sheetを開く
		/// </summary>
		/// <param name="_SheetIndex"></param>
		/// <returns></returns>
		public ExcelSheetAccessor Open(object _SheetIndex)
		{
			try
			{
				return new ExcelSheetAccessor(_ExcelSheets[_SheetIndex]);
			}
			catch (Exception e)
			{
				throw new Exception("[" + _SheetIndex + "] の実行に失敗しました", e);
			}
		}

		/// <summary>
		/// Sheetを追加する
		/// </summary>
		/// <param name="_SheetName"></param>
		/// <returns></returns>
		public ExcelSheetAccessor Add(string _SheetName)
		{
			Excel.Worksheet _ExcelSheet = null;

			try
			{
				_ExcelSheet = _ExcelSheets.Add();
				_ExcelSheet.Name = _SheetName;

				return new ExcelSheetAccessor(_ExcelSheet);
			}
			catch (Exception e)
			{
				this.ReleaseObject(_ExcelSheet);

				throw new Exception("[" + _SheetName + "] の追加に失敗しました", e);
			}
		}

		/// <summary>
		/// ブックを保存する
		/// </summary>
		/// <param name="_FilePath">保存先のパス。新規保存したいときのみ必要</param>
		public void Save(string _FilePath = "")
		{
			if (_FilePath != "")
			{
				_ExcelBook.SaveAs(_FilePath);
				return;
			}

			if (_ExcelBook.Path != "")
			{
				_ExcelBook.Save();
				return;
			}
		}

		public void Close()
		{
			if (_ExcelBook != null)	_ExcelBook.Close();
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

		private void ReleaseBook()
		{
			// Sheets解放
			this.ReleaseObject(_ExcelSheets);

			// Book解放
			this.ReleaseObject(_ExcelBook);
		}

		protected virtual void Dispose(bool _Disposing)
		{
			if (!_DisposeValue)
			{
				if (_Disposing)	{}

				this.ReleaseBook();
			}
		}

		public void Dispose()
		{
			this.Dispose(true);
			GC.SuppressFinalize(this);
		}

		~ExcelBookAccessor()
		{
			this.Dispose(false);
		}

		private bool _DisposeValue = false;
		
		private Excel.Workbook _ExcelBook = null;
		private Excel.Sheets _ExcelSheets = null;
	}
}
