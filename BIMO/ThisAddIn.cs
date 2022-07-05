using System;
using Visio = Microsoft.Office.Interop.Visio;
using Microsoft.Office.Interop.Visio;

namespace BIMO
{
	public partial class ThisAddIn
	{
		private Visio.Application _visio;

		private void ThisAddIn_Startup(object sender, EventArgs e)
		{
			_visio = Globals.ThisAddIn.Application;

			// Register shapeadded event for all newly added shape
			_visio.ShapeAdded += OnShapeAdded;
			_visio.BeforeShapeDelete += OnBeforeShapeDeleted;
		}

		private void OnBeforeShapeDeleted(Shape shape)
		{
			// delete bkgndShape  if exist
			if (Convert.ToBoolean(shape.CellExistsU["User.bkgndShape", 1]))
			{
				var uniqueId = shape.Cells["User.bkgndShape"].ResultStr[""];
				var item = shape.ContainingPage.BackPage.Shapes.ItemU[uniqueId];
				if (item is Shape bkgndShape)
					bkgndShape.Delete();
			}
		}

		private void OnShapeAdded(Shape shape)
		{
			// ignore if current page isn't a foreground page
			if (Convert.ToBoolean(shape.ContainingPage.Background)) return;

			// add background page if not exist
			if (shape.ContainingPage.BackPage == null) shape.ContainingPage.AppendBackground();

			// add the related shape into background
			shape.CopyToBackground();

			// Switch back to frontpage
			_visio.ActiveWindow.Page = shape.ContainingPage;
		}


		private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
		{
			_visio.ShapeAdded -= OnShapeAdded;
		}

		#region VSTO generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InternalStartup()
		{
			this.Startup += new System.EventHandler(ThisAddIn_Startup);
			this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
		}

		#endregion
	}
}
