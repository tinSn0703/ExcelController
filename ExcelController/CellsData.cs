using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp1
{
	/// <summary>
	/// Cellのデータを表すオブジェクト。
	/// </summary>
	class CellsData : IDisposable
	{
		public CellsData()
		{
		}

		public CellsData(Excel.Range _ExcelRange)
		{
			this.ReleaseRange();
			this._ExcelRange = _ExcelRange ?? throw new ArgumentNullException(nameof(_ExcelRange));
		}

		public CellsData(Excel.Worksheet _ExcelSheet, string _Cell1 = "", string _Cell2 = "")
		{
			this.ReleaseRange();
			this._ExcelRange = _ExcelSheet.Range[_Cell1, _Cell2] 
				?? throw new ArgumentNullException(nameof(_ExcelSheet) + " " + _Cell1 + ", " + _Cell2);
		}

		public void GetValue()
		{

		}

		public void Fet()
		{

		}

		public void Clear()
		{

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

		private void ReleaseRange()
		{
			//Range解放
			this.ReleaseObject(_ExcelRange);
		}

		private Excel.Range _ExcelRange = null;

		#region IDisposable Support
		private bool disposedValue = false; // 重複する呼び出しを検出するには

		protected virtual void Dispose(bool disposing)
		{
			if (!disposedValue)
			{
				if (disposing)
				{
					// TODO: マネージ状態を破棄します (マネージ オブジェクト)。
				}

				// TODO: アンマネージ リソース (アンマネージ オブジェクト) を解放し、下のファイナライザーをオーバーライドします。
				// TODO: 大きなフィールドを null に設定します。
				this.ReleaseRange();

				disposedValue = true;
			}
		}

		// TODO: 上の Dispose(bool disposing) にアンマネージ リソースを解放するコードが含まれる場合にのみ、ファイナライザーをオーバーライドします。
		~CellsData()
		{
			// このコードを変更しないでください。クリーンアップ コードを上の Dispose(bool disposing) に記述します。
			Dispose(false);
		}

		// このコードは、破棄可能なパターンを正しく実装できるように追加されました。
		public void Dispose()
		{
			// このコードを変更しないでください。クリーンアップ コードを上の Dispose(bool disposing) に記述します。
			Dispose(true);
			// TODO: 上のファイナライザーがオーバーライドされる場合は、次の行のコメントを解除してください。
			GC.SuppressFinalize(this);
		}
		#endregion
	}
}
