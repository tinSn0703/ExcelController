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
	class ExcelSheetAccessor : IDisposable
	{
		public ExcelSheetAccessor()
		{
			_ExcelSheet = null;
		}

		public ExcelSheetAccessor(Excel.Worksheet _ExcelSheet)
		{
			this.ReleaseSheet();

			this._ExcelSheet = _ExcelSheet ?? throw new ArgumentNullException(nameof(_ExcelSheet));
		}

		public Excel.Range GetRange(string _Cell1)
		{
			return _ExcelSheet.Range[_Cell1];
		}

		public Excel.Range GetRange(string _Cell1, string _Cell2)
		{
			return _ExcelSheet.Range[_Cell1, _Cell2];
		}

		/// <summary>
		/// Objectを開放する
		/// </summary>
		private void ReleaseObject(object _Obj)
		{
			if (_Obj != null)
			{
				while (Marshal.ReleaseComObject(_Obj) > 0);
				_Obj = null;
			}
		}

		private void ReleaseSheet()
		{
			//Worksheet解放
			this.ReleaseObject(_ExcelSheet);
		}

		protected virtual void Dispose(bool _Disposing)
		{
			if (!_DisposeValue)
			{
				if (_Disposing)	{}

				this.ReleaseSheet();
				
				_DisposeValue = true;
			}
		}

		public void Dispose()
		{
			this.Dispose(true);
			GC.SuppressFinalize(this);
		}

		~ExcelSheetAccessor()
		{
			this.Dispose(false);
		}

		private bool _DisposeValue = false;
		private Excel.Worksheet _ExcelSheet = null;
	}
}
