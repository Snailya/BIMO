using Microsoft.Office.Interop.Visio;
using System;

namespace BIMO
{
	internal static class Extensions
	{
		internal static void AppendBackground(this Page page, string name = "")
		{
			var app = page.Application;

			var scopeId = app.BeginUndoScope("Insert Page");

			var background = app.ActiveDocument.Pages.Add();
			background.Background = 1;
			background.Name = string.IsNullOrEmpty(name) ? $"{page.Name}_bg" : name;
			page.BackPage = background;

			app.EndUndoScope(scopeId, true);
		}

		internal static Shape CopyToBackground(this Shape foregndShape, Master master = null)
		{
			var app = foregndShape.Application;

			var scopeId = app.BeginUndoScope("Copy To Background");

			// prepare page, master and positon
			var backpage = foregndShape.ContainingPage.BackPage;
			if (master == null)
			{
				// use the same shape if user not specify the master to use
				master = foregndShape.Master;
			}
			var (positionX, positionY) = (foregndShape.Cells["PinX"].Result[""], foregndShape.Cells["PinY"].Result[""]);

			// place related shape into background page
			var bkgndShape = backpage.Drop(master, positionX, positionY);

#if DEBUG
			// this is only for demo reason to highlight the  background shape
			bkgndShape.Cells["FillForegnd"].FormulaU = "THEMEGUARD(RGB(255, 255, 0))";
#endif

			// persist background shape id to foreground shape's sheet
			var rowName = "bkgndShape";
			var uniqueId = bkgndShape.UniqueID[(short)VisUniqueIDArgs.visGetOrMakeGUID];

			foregndShape.AddNamedRow((short)VisSectionIndices.visSectionUser, rowName, (short)VisRowTags.visTagDefault);
			foregndShape.Cells[$"User.{rowName}"].Formula = $"\"{uniqueId}\"";

			foregndShape.CellChanged += OnPositionChanged;

			app.EndUndoScope(scopeId, true);

			return bkgndShape;
		}

		private static void OnPositionChanged(Cell Cell)
		{
			if (!Convert.ToBoolean(Cell.Shape.CellExistsU["User.bkgndShape", 1])) return;

			var uniqueId = Cell.Shape.Cells["User.bkgndShape"].ResultStr[""];
			var bkgndShape = Cell.Shape.ContainingPage.BackPage.Shapes.ItemU[uniqueId];

			if (Cell.Name == "PinX")
			{
				bkgndShape.Cells["PinX"].Result[""] = Cell.Result[""];
			}
			if (Cell.Name == "PinY")
			{
				bkgndShape.Cells["PinY"].Result[""] = Cell.Result[""];
			}
		}
	}
}
