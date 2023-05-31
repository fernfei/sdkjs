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
		this.precedentsExternal = null;
		this.currentPrecedents = null;
		this.dependents = null;
		this.inLoop = null;
		this.isPrecedentsCall = null;
		this.precedentsAreas = null;
	}

	TraceDependentsManager.prototype.setPrecedentsCall = function () {
		this.isPrecedentsCall = true;
	};
	TraceDependentsManager.prototype.setDependentsCall = function () {
		this.isPrecedentsCall = null;
	};
	TraceDependentsManager.prototype.setPrecedentExternal = function (cellIndex) {
		if (!this.precedentsExternal) {
			this.precedentsExternal = new Set();
		}
		this.precedentsExternal.add(cellIndex);
	};
	TraceDependentsManager.prototype.checkPrecedentExternal = function (cellIndex) {
		if (!this.precedentsExternal) {
			return false;
		}
		return this.precedentsExternal.has(cellIndex);
	};
	TraceDependentsManager.prototype.calculateDependents = function (row, col) {
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
			this.setDependentsCall();
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
		const findCellListeners = function () {
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

		const getAllAreaIndexes = function (cell) {
			const indexes = [], range = cell.ref;
			if (!range) {
				return;
			}
			for (let i = range.c1; i <= range.c2; i++) {
				for (let j = range.r1; j <= range.r2; j++) {
					let index = AscCommonExcel.getCellIndex(j, i);
					indexes.push(index);
				}
			}

			return indexes;
		};

		const getParentIndex = function (_parent) {
			let _parentCellIndex = AscCommonExcel.getCellIndex(_parent.nRow, _parent.nCol);
			//parent -> cell/defname
			if (_parent.parsedRef/*parent instanceof AscCommonExcel.DefName*/) {
				_parentCellIndex = null;
			} else if (_parent.ws !== t.ws.model) {
				_parentCellIndex += ";" + _parent.ws.index;
			}
			return _parentCellIndex;
		};

		const cellListeners = findCellListeners();
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
						// TODO in one formula there can be both references to the external and internal area
						if (formula.includes(":") && !cellListeners[i].is3D) {
							let areaIndexes;
							// call splitAreaListeners which return cellIndexes of each element(this will be parentCellIndex)
							// go through the values and _setDependents/Precedents for each
							areaIndexes = getAllAreaIndexes(cellListeners[i]);
							if (areaIndexes) {
								for (let index of areaIndexes) {
									this._setDependents(cellIndex, index);
									// this._setPrecedents(index, cellIndex);
								}
								continue;
							}
						}
						this._setDependents(cellIndex, parentCellIndex);
						// this._setPrecedents(parentCellIndex, cellIndex);
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
						} else if (parentCellIndex == cellIndex) {
							// TODO bug the cell refers to itself and the call stack overflows
							// temporary exception
							isUpdated = true;
							break;
						}
						// TODO in one formula there can be both references to the external and internal area
						if (formula.includes(":") && !cellListeners[i].is3D) {
							// call getAllAreaIndexes which return cellIndexes of each element(this will be parentCellIndex)
							let areaIndexes = getAllAreaIndexes(cellListeners[i]);
							if (areaIndexes) {
								// go through the values and _setDependents/Precedents for each
								for (let index of areaIndexes) {
									this._setDependents(cellIndex, index);
									// this._setPrecedents(index, cellIndex);
								}
								continue;
							}
						}
						if (!this._getDependents(cellIndex, parentCellIndex)) {
							this._setDependents(cellIndex, parentCellIndex);
							// this._setPrecedents(parentCellIndex, cellIndex);
							isUpdated = true;
						}
					}
				}
				
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
		if (!this.precedentsAreas) {
			this.precedentsAreas = {};
		}

		let t = this;
		let currentCellIndex = AscCommonExcel.getCellIndex(row, col);

		if (!this.precedents[currentCellIndex]) {
			if (formulaParsed.outStack) {
				let currentWsIndex = formulaParsed.ws.index;
				// iterate and find all reference
				for (const elem of formulaParsed.outStack) {
					let elemType = elem.type ? elem.type : null;
					// 6 - ref
					// 5 - cellsRange
					// 12 - ref3D
					// 13 - cellsRange3D
					if (elemType === 6 || elemType === 5 || elemType === 12 || elemType === 13) {
						const areaRange = {};
						let is3D = elemType === 12 || elemType === 13;
						let isArea = elemType === 5;
						let elemRange = elem.range.bbox ? elem.range.bbox : elem.bbox;
						let elemCellIndex = AscCommonExcel.getCellIndex(elemRange.r1, elemRange.c1);

						let elemDependentsIndex = elemCellIndex;
						let currentDependentsCellIndex = currentCellIndex;
	
						if (is3D) {
							elemCellIndex += ";" + (elem.wsTo ? elem.wsTo.index : elem.ws.index);	
							currentDependentsCellIndex += ";" + currentWsIndex
							this._setPrecedents(currentCellIndex, elemCellIndex);
							this.setPrecedentExternal(currentCellIndex);
						} else {
							this._setPrecedents(currentCellIndex, elemCellIndex);
							// change links to another sheet
							this._setDependents(elemDependentsIndex, currentDependentsCellIndex);
						}

						if (isArea) {
							// write 4 index to an object - top left cell, top right cell, bottom right cell, bottom left cell - area stroke coordinates
							// TODO minify or refactor
							const tempObj = {};
							const areaName = elem.value;	// areaName - unique key for areaRange
							tempObj.topLeftIndex = AscCommonExcel.getCellIndex(elemRange.r1, elemRange.c1);
							tempObj.topRightIndex = AscCommonExcel.getCellIndex(elemRange.r1, elemRange.c2);
							tempObj.bottomRightIndex = AscCommonExcel.getCellIndex(elemRange.r2, elemRange.c2);
							tempObj.bottomLeftIndex = AscCommonExcel.getCellIndex(elemRange.r2, elemRange.c1);
							areaRange[areaName] = tempObj;

							this._setPrecedentsAreas(areaRange);
						}
					}
				}
			}
		} else if (!this.getPrecedentsLoop()) {
			const currentPrecedent = {...this.precedents};
			this.setPrecedentsLoop(true);

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

			this.setPrecedentsLoop(false);
		}
	};
	TraceDependentsManager.prototype.setPrecedentsLoop = function (inLoop) {
		this.inLoop = inLoop;
	};
	TraceDependentsManager.prototype.getPrecedentsLoop = function () {
		return this.inLoop;
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
	TraceDependentsManager.prototype._setPrecedentsAreas = function (area) {
		if (!this.precedentsAreas) {
			this.precedentsAreas = {};
		}
		Object.assign(this.precedentsAreas, area);
	};
	TraceDependentsManager.prototype._getPrecedentsAreas = function () {
		return this.precedentsAreas;
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
			callback(i, this.dependents[i], this.isPrecedentsCall);
		}
	};
	TraceDependentsManager.prototype.forEachPrecedents = function (callback) {
		for (let i in this.precedents) {
			callback(i);
		}
	};
	TraceDependentsManager.prototype.clear = function (type) {
		if (Asc.c_oAscRemoveArrowsType.all === type || Asc.c_oAscRemoveArrowsType.precedent === type) {
			this.precedents = null;
		}
		if (Asc.c_oAscRemoveArrowsType.all === type || Asc.c_oAscRemoveArrowsType.dependent === type) {
			this.dependents = null;
		}
		if (Asc.c_oAscRemoveArrowsType.all === type || Asc.c_oAscRemoveArrowsType.dependent === type) {
			this.precedentsExternal = null;
		}
		if (Asc.c_oAscRemoveArrowsType.all === type || Asc.c_oAscRemoveArrowsType.dependent === type) {
			this.precedentsAreas = null;
		}
	};







	//------------------------------------------------------------export---------------------------------------------------
	window['AscCommonExcel'] = window['AscCommonExcel'] || {};

	window["AscCommonExcel"].TraceDependentsManager = TraceDependentsManager;


})(window);
