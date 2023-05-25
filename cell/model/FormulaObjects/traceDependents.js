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

"use strict";
(/**
 * @param {Window} window
 * @param {undefined} undefined
 */
function (window, undefined) {

	/*
	 * Import
	 * -----------------------------------------------------------------------------
	 */

	function TraceDependentsManager(ws) {
		this.ws = ws;
		this.precedents = null;
		this.currentPrecedents = null;
		this.dependents = null;
		this.inLoop = null;
		this.isPrecedentsCall = null;
	}
	TraceDependentsManager.prototype.setPrecedentsCall = function () {
		this.isPrecedentsCall = true;
	};
	TraceDependentsManager.prototype.setCurrentPrecedents = function (currentPrecedents) {
		this.currentPrecedents = {...currentPrecedents};
	};
	TraceDependentsManager.prototype.calculateDependents = function (row, col) {
		//depend from row/col cell
		let ws = this.ws && this.ws.model;
		if (!ws) {
			return;
		}
		if (row == null || col == null) {
			let selection = ws.getSelection();
			let activeCell = selection.activeCell;
			row = activeCell.row;
			col = activeCell.col;
		}

		let depFormulas = ws.workbook.dependencyFormulas;
		if (depFormulas && depFormulas.sheetListeners) {
			if (!this.dependents) {
				this.dependents = {};
			}

			let sheetListeners = depFormulas.sheetListeners;
			let curListener = sheetListeners[ws.Id];
			let cellIndex = AscCommonExcel.getCellIndex(row, col);
			this._calculateDependents(cellIndex, curListener);
		}
	};
	TraceDependentsManager.prototype._calculateDependents = function (cellIndex, curListener) {
		if (!this.dependents) {
			this.dependents = {};
		}

		let t = this;
		let wb = this.ws.model.workbook;
		let dependencyFormulas = wb.dependencyFormulas;
		let cellAddress = AscCommonExcel.getFromCellIndex(cellIndex, true);
		let findCellListeners = function () {
			const listeners = {};
			if (curListener && curListener.areaMap) {
				for (let j in curListener.areaMap) {
					if (curListener.areaMap.hasOwnProperty(j)) {
						if (curListener.areaMap[j] && curListener.areaMap[j].bbox.contains(cellAddress.col, cellAddress.row)) {
							Object.assign(listeners, curListener.areaMap[j].listeners);
						}
					}
				}
			}
			if (curListener && curListener.cellMap && curListener.cellMap[cellIndex]) {
				if (Object.keys(curListener.cellMap[cellIndex]).length > 0) {
					Object.assign(listeners, curListener.cellMap[cellIndex].listeners);
				}
			}
			return listeners;
		};

		let getAllAreaIndexes = function (cell) {
			const indexes = [], range = cell.ref;
			for (let i = range.c1; i <= range.c2; i++) {
				for (let j = range.r1; j <= range.r2; j++) {
					let index = AscCommonExcel.getCellIndex(j, i);
					indexes.push(index);
				}
			}

			return indexes;
		};

		let getParentIndex = function (_parent) {
			let _parentCellIndex = AscCommonExcel.getCellIndex(_parent.nRow, _parent.nCol);
			//parent -> cell/defname
			if (_parent.parsedRef/*parent instanceof AscCommonExcel.DefName*/) {
				_parentCellIndex = null;
			} else if (_parent.ws !== t.ws.model) {
				_parentCellIndex += ";" + _parent.ws.index;
			}
			return _parentCellIndex;
		};

		let cellListeners = findCellListeners();
		if (cellListeners) {
			if (!this.dependents[cellIndex]) {
				this.dependents[cellIndex] = {};
				for (let i in cellListeners) {
					if (cellListeners.hasOwnProperty(i)) {
						let parent = cellListeners[i].parent;
						let parentCellIndex = getParentIndex(parent);
						let formula = cellListeners[i].Formula;
						if (parentCellIndex === null) {
							continue;
						}
						if (formula.includes(":") && !cellListeners[i].is3D) {
							let areaIndexes;
							// call splitAreaListeners which return cellIndexes of each element(this will be parentCellIndex)
							// go through the values and _setDependents/Precedents for each
							areaIndexes = getAllAreaIndexes(cellListeners[i]);
							for (let index of areaIndexes) {
								this._setDependents(cellIndex, index);
								this._setPrecedents(index, cellIndex);
							}
							continue;
						}
						this._setDependents(cellIndex, parentCellIndex);
						this._setPrecedents(parentCellIndex, cellIndex);
					}
				}
			} else {
				//if change formulas and add new sheetListeners
				//check current tree
				let isUpdated = false;
				for (let i in cellListeners) {
					if (cellListeners.hasOwnProperty(i)) {
						let parent = cellListeners[i].parent;
						let parentCellIndex = getParentIndex(parent);
						let formula = cellListeners[i].Formula;
						if (parentCellIndex === null) {
							continue;
						}
						if (formula.includes(":") && !cellListeners[i].is3D) {
							let areaIndexes;
							// call splitAreaListeners which return cellIndexes of each element(this will be parentCellIndex)
							// go through the values and _setDependents/Precedents for each
							areaIndexes = getAllAreaIndexes(cellListeners[i]);
							for (let index of areaIndexes) {
								this._setDependents(cellIndex, index);
								this._setPrecedents(index, cellIndex);
							}
						}
						if (!this._getDependents(cellIndex, parentCellIndex)) {
							this._setDependents(cellIndex, parentCellIndex);
							this._setPrecedents(parentCellIndex, cellIndex);
							isUpdated = true;
						}
					}
				}
				// TODO bug (ссылка на себя в ячейке O3: =O2, O2: =O1, O1: =cEmpty) происходит переполнение стека(невозможно найти cellIndex)
				if (!isUpdated) {
					for (let i in this.dependents[cellIndex]) {
						if (this.dependents[cellIndex].hasOwnProperty(i)) {
							this._calculateDependents(i, curListener);
						}
					}
				}
			}
		}
	};
	TraceDependentsManager.prototype._getDependents = function (from, to) {
		return this.dependents[from] && this.dependents[from][to];
	};
	TraceDependentsManager.prototype._setDependents = function (from, to) {
		if (!this.dependents) {
			this.dependents = {};
		}
		if (!this.dependents[from]) {
			this.dependents[from] = {};
		}
		this.dependents[from][to] = 1;
	};
	TraceDependentsManager.prototype.calculatePrecedents = function (row, col) {
		//depend from row/col cell
		let ws = this.ws && this.ws.model;
		if (!ws) {
			return;
		}
		if (row == null || col == null) {
			let selection = ws.getSelection();
			let activeCell = selection.activeCell;
			row = activeCell.row;
			col = activeCell.col;
		}

		let formulaParsed;
		ws.getCell3(row, col)._foreachNoEmpty(function (cell) {
			formulaParsed = cell.formulaParsed;
		});

		if (formulaParsed) {
			this._calculatePrecedents(formulaParsed, row, col);
			this.setPrecedentsCall();
		}
	};
	TraceDependentsManager.prototype._calculatePrecedents = function (formulaParsed, row, col) {
		if (!this.precedents) {
			this.precedents = {};
		}
		// if (!this.precedentsAreas) {
		// 	this.precedentsAreas = {};
		// }

		const getElemIndex = function (_row, _col, is3D, elem) {
			let _cellIndex = AscCommonExcel.getCellIndex(_row, _col);
			if (is3D) {
				_cellIndex += ";" + elem.wsTo.index;
			}
			return _cellIndex;
		};

		let t = this;
		let currentCellIndex = getElemIndex(row, col);

		if (!this.precedents[currentCellIndex]) {
			if (formulaParsed.outStack) {
				let currentWsIndex = formulaParsed.ws.index;
				// iterate and find all reference
				// if reference already in the map, skip it
				// write dependencies too
				// two-way recording - if a cell depends on another, the other one affects the primary cell
				for (const elem of formulaParsed.outStack) {
					let elemType = elem.type ? elem.type : null;
					// 6 - ref
					// 5 - cellsRange
					// 12 - ref3D
					// 13 - cellsRange3D
					if (elemType === 6 || elemType === 5 || elemType === 12 || elemType === 13) {
						let is3D = elemType === 12 || elemType === 13;
						let elemRange = elem.range.bbox ? elem.range.bbox : elem.bbox;
						let elemCellIndex = getElemIndex(elemRange.r1, elemRange.c1);

						let elemDependentsIndex = elemCellIndex;
						let currentDependentsCellIndex = currentCellIndex;
	
						if (is3D) {
							elemCellIndex += ";" + elem.ws.index;	
							currentDependentsCellIndex += ";" + currentWsIndex
						}
						this._setPrecedents(currentCellIndex, elemCellIndex);
						// change links to another sheet
						this._setDependents(elemDependentsIndex, currentDependentsCellIndex);
					}
				}
			}
		} else if (!this.inLoop){
			const currentPrecedent = {...this.precedents};
			let isUpdated = false;
			this.inLoop = true;
			if (!isUpdated) {
				for (let i in currentPrecedent) {
					if (currentPrecedent.hasOwnProperty(i)) {
						for (let j in currentPrecedent[i]) {
							if (currentPrecedent[i].hasOwnProperty(j)) {
								// get row col from cell index
								let coords = AscCommonExcel.getFromCellIndex(j, true);
								this.calculatePrecedents(coords.row, coords.col);
							}
						}
					}
				}
				isUpdated = true;
				this.inLoop = false;
			}
		}
	};
	TraceDependentsManager.prototype._getPrecedents = function (from, to) {
		return this.precedents[from] && this.precedents[from][to];
	};
	TraceDependentsManager.prototype._setPrecedents = function (from, to) {
		if (!this.precedents) {
			this.precedents = {};
		}
		if (!this.precedents[from]) {
			this.precedents[from] = {};
		}
		this.precedents[from][to] = 1;
	};
	TraceDependentsManager.prototype.isHaveData = function () {
		return this.isHaveDependents() || this.isHavePrecedents();
	};
	TraceDependentsManager.prototype.isHaveDependents = function () {
		return !!this.dependents;
	};
	TraceDependentsManager.prototype.isHavePrecedents = function () {
		return !!this.precedents;
	};
	TraceDependentsManager.prototype.forEachDependents = function (callback) {
		for (let i in this.dependents) {
			callback(i, this.dependents[i]);
		}
	};
	TraceDependentsManager.prototype.forEachPrecedents = function (callback) {
		for (let i in this.precedents) {
			callback(this.precedents[i], i);
		}
	};
	TraceDependentsManager.prototype.clear = function (type) {
		if (Asc.c_oAscRemoveArrowsType.all === type || Asc.c_oAscRemoveArrowsType.precedent === type) {
			this.precedents = null;
		}
		if (Asc.c_oAscRemoveArrowsType.all === type || Asc.c_oAscRemoveArrowsType.dependent === type) {
			this.dependents = null;
		}
	};







	//------------------------------------------------------------export---------------------------------------------------
	window['AscCommonExcel'] = window['AscCommonExcel'] || {};

	window["AscCommonExcel"].TraceDependentsManager = TraceDependentsManager;


})(window);