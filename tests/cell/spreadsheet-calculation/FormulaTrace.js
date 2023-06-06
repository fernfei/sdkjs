/*
 * (c) Copyright Ascensio System SIA 2010-2023
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA at 20A-6 Ernesta Birznieka-Upish
 * street, Riga, Latvia, EU, LV-1050.
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * Pursuant to Section 7(b) of the License you must retain the original Product
 * logo when distributing the program. Pursuant to Section 7(e) we decline to
 * grant you any rights under trademark law for use of our trademarks.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */
QUnit.config.autostart = false;
$(function() {

	Asc.spreadsheet_api.prototype._init = function() {
		this._loadModules();
	};
	Asc.spreadsheet_api.prototype._loadFonts = function(fonts, callback) {
		callback();
	};
	Asc.spreadsheet_api.prototype.onEndLoadFile = function(fonts, callback) {
		openDocument();
	};
	AscCommonExcel.WorkbookView.prototype._calcMaxDigitWidth = function() {
	};
	AscCommonExcel.WorkbookView.prototype._init = function() {
	};
	AscCommonExcel.WorkbookView.prototype._onWSSelectionChanged = function() {
	};
	AscCommonExcel.WorkbookView.prototype.showWorksheet = function() {
	};
	AscCommonExcel.WorksheetView.prototype._init = function() {
	};
	AscCommonExcel.WorksheetView.prototype.updateRanges = function() {
	};
	AscCommonExcel.WorksheetView.prototype._autoFitColumnsWidth = function() {
	};
	AscCommonExcel.WorksheetView.prototype.setSelection = function() {
	};
	AscCommonExcel.WorksheetView.prototype.draw = function() {
	};
	AscCommonExcel.WorksheetView.prototype._prepareDrawingObjects = function() {
	};

	AscCommon.baseEditorsApi.prototype._onEndLoadSdk = function() {
	};

	let g_oIdCounter = AscCommon.g_oIdCounter;

	let wb, ws, sData = AscCommon.getEmpty(), api;
	if (AscCommon.c_oSerFormat.Signature === sData.substring(0, AscCommon.c_oSerFormat.Signature.length)) {
		Asc.spreadsheet_api.prototype._init = function() {
		};
		
		api = new Asc.spreadsheet_api({
			'id-view': 'editor_sdk'
		});

		api.FontLoader = {
			LoadDocumentFonts: function() {
				setTimeout(startTests, 0)
			}
		};

		let docInfo = new Asc.asc_CDocInfo();
		docInfo.asc_putTitle("TeSt.xlsx");
		api.DocInfo = docInfo;

		window["Asc"]["editor"] = api;
		AscCommon.g_oTableId.init();
		if (this.User) {
			g_oIdCounter.Set_UserId(this.User.asc_getId());
		}
		api._onEndLoadSdk();
		api.isOpenOOXInBrowser = false;
		api._openDocument(AscCommon.getEmpty());	// this func set api.wbModel
		// api._openOnClient();
		api.collaborativeEditing = new AscCommonExcel.CCollaborativeEditing({});
		api.wb = new AscCommonExcel.WorkbookView(api.wbModel, api.controller, api.handlers, api.HtmlElement,
			api.topLineEditorElement, api, api.collaborativeEditing, api.fontRenderingMode);
		wb = api.wbModel;

		AscCommonExcel.g_oUndoRedoCell = new AscCommonExcel.UndoRedoCell(wb);
		AscCommonExcel.g_oUndoRedoWorksheet = new AscCommonExcel.UndoRedoWoorksheet(wb);
		AscCommonExcel.g_oUndoRedoWorkbook = new AscCommonExcel.UndoRedoWorkbook(wb);
		AscCommonExcel.g_oUndoRedoCol = new AscCommonExcel.UndoRedoRowCol(wb, false);
		AscCommonExcel.g_oUndoRedoRow = new AscCommonExcel.UndoRedoRowCol(wb, true);
		AscCommonExcel.g_oUndoRedoComment = new AscCommonExcel.UndoRedoComment(wb);
		AscCommonExcel.g_oUndoRedoAutoFilters = new AscCommonExcel.UndoRedoAutoFilters(wb);
		AscCommonExcel.g_DefNameWorksheet = new AscCommonExcel.Worksheet(wb, -1);
		g_oIdCounter.Set_Load(false);

		let oBinaryFileReader = new AscCommonExcel.BinaryFileReader();
		oBinaryFileReader.Read(sData, wb);
		// ws = wb.getWorksheet(wb.getActive());
		ws = api.wbModel.aWorksheets[0];
		AscCommonExcel.getFormulasInfo();

		api.collaborativeEditing = new AscCommonExcel.CCollaborativeEditing({});
		api.wb = new AscCommonExcel.WorkbookView(api.wbModel, api.controller, api.handlers, api.HtmlElement,
			api.topLineEditorElement, api, api.collaborativeEditing, api.fontRenderingMode);

	}

	function traceTests() {
		QUnit.test("Test: \"test\"", function (assert) {
			// create cells with dependencies
			ws.getRange2("A1").setValue("1");
			ws.getRange2("B101").setValue("=A1");
			ws.getRange2("C101").setValue("=B101");
	
			// click" on the button show dependences/precendence
			api.asc_setCellBold();
			api.asc_setCellBold();

			// check the object with dependency cell numbers for compliance
			let wsView = api.wb.getWorksheet();
			let traceManager = wsView.traceDependentsManager;

			let A1Index = AscCommonExcel.getCellIndex(ws.getRange2("A1").bbox.r1, ws.getRange2("A1").bbox.c1),
				B101Index = AscCommonExcel.getCellIndex(ws.getRange2("B101").bbox.r1, ws.getRange2("B101").bbox.c1),
				C101Index = AscCommonExcel.getCellIndex(ws.getRange2("C101").bbox.r1, ws.getRange2("C101").bbox.c1);
			
			// check A1 -> B101
			assert.strictEqual(traceManager._getDependents(A1Index, B101Index), 1);

			// check B101 -> C101
			assert.strictEqual(traceManager._getDependents(B101Index, C101Index), 1);

		});
	}

	QUnit.module("FormulaTrace");

	function startTests() {
		QUnit.start();
		traceTests();
	}

	startTests();
});
