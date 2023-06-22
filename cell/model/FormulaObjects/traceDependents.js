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
			let allDefNamesListeners = depFormulas.defNameListeners;
			let cellIndex = AscCommonExcel.getCellIndex(row, col);
			this._calculateDependents(cellIndex, curListener, allDefNamesListeners);
			this.setDependentsCall();
		}
	};
	TraceDependentsManager.prototype._calculateDependents = function (cellIndex, curListener, allDefNamesListeners) {
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
			// TODO need to find all ranges referring to defName
			if (curListener && curListener.defName3d) {
				Object.assign(listeners, curListener.defName3d);
			}
			return listeners;
		};
		const setDefNameIndexes = function (defName) {
			for (const i in allDefNamesListeners) {
				if (allDefNamesListeners.hasOwnProperty(i) && i === defName) {
					for (const listener in allDefNamesListeners[i].listeners) {
						// TODO возможно стоить добавить все слушатели сразу в curListener
						let isArea = allDefNamesListeners[i].listeners[listener].ref ? !allDefNamesListeners[i].listeners[listener].ref.isOneCell() : false;
						if (isArea) {
							// decompose all elements into dependencies
							let areaIndexes = getAllAreaIndexes(allDefNamesListeners[i].listeners[listener]);
							if (areaIndexes) {
								for (let index of areaIndexes) {
									t._setDependents(cellIndex, index);
								}
								continue;
							}
						}
						let parentCellIndex = getParentIndex(allDefNamesListeners[i].listeners[listener].parent);
						t._setDependents(cellIndex, parentCellIndex);
					}
				}
			}
		};
		const getAllAreaIndexes = function (parserFormula) {
			const indexes = [], range = parserFormula.ref;
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

		const getParentIndex = function (_parent, shared) {
			let _parentCellIndex = AscCommonExcel.getCellIndex(_parent.nRow, _parent.nCol);
			
			if (shared) {
				// ? temporary solution
				let base = shared.base;
				let isRowMode = (shared.ref.r2 - shared.ref.r1) !== 0 ? false : true,
					isColumnMode = (shared.ref.c2 - shared.ref.c1) !== 0 ? false : true;

				// let isBaseElement = (cellAddress.col - (cellAddress.col - base.nCol) === base.nCol) ? true : false;
				// cellAddress.col - (cellAddress.col - base.nCol - 1)
				// cellAddress.col - (Math.abs(cellAddress.col - base.nCol - (isBaseElement ? 0 : 1))) 
				// _parentCellIndex = AscCommonExcel.getCellIndex(_parent.nRow + (cellAddress.row - base.nRow), _parent.nCol + (cellAddress.col - base.nCol));
				// _parentCellIndex = AscCommonExcel.getCellIndex(_parent.nRow, _parent.nCol + (cellAddress.col + (cellAddress.row === _parent.nRow ? 1 : -1) - base.nCol));
				// _parentCellIndex = AscCommonExcel.getCellIndex(_parent.nRow, _parent.nCol + (cellAddress.col + (cellAddress.row === _parent.nRow ? 1 : -1) - base.nCol));
				if (isRowMode) {
					let newCol = cellAddress.row === _parent.nRow ? _parent.nCol + (cellAddress.col + 1 - base.nCol) : _parent.nCol;
					_parentCellIndex = AscCommonExcel.getCellIndex(_parent.nRow, newCol);
				} else {
					let newRow = cellAddress.col === _parent.nCol ? _parent.nRow + (cellAddress.row + 1 - base.nRow) : _parent.nRow;
					_parentCellIndex = AscCommonExcel.getCellIndex(newRow, _parent.nCol);
				}
			}
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
				// if dependents by cellIndex didn't exist, create it
				this.dependents[cellIndex] = {};
				for (let i in cellListeners) {
					if (cellListeners.hasOwnProperty(i)) {
						let parent = cellListeners[i].parent;
						let formula = cellListeners[i].Formula;

						let is3D = false;
						if (parent.name) {
							is3D = false;
							setDefNameIndexes(parent.name);
							continue;
						} else if (cellListeners[i].is3D) {
							is3D = true;
						}
						let parentCellIndex = getParentIndex(parent);
						
						if (parentCellIndex === null) {
							continue;
						}

						if (cellListeners[i].shared !== null) {
							let shared = cellListeners[i].getShared();
							parentCellIndex = getParentIndex(parent, shared);
						}

						if (formula.includes(":") && !is3D) {
							let areaIndexes;
							// call splitAreaListeners which return cellIndexes of each element(this will be parentCellIndex)
							// go through the values and set dependents for each
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
					}
				}
			} else {
				// if dependents by cellIndex aldready exist, check current tree
				let currentIndex = Object.keys(this.dependents[cellIndex])[0];
				let isUpdated = false;
				for (let i in cellListeners) {
					if (cellListeners.hasOwnProperty(i)) {
						let parent = cellListeners[i].parent;
						let elemCellIndex = cellListeners[i].shared !== null ? currentIndex : getParentIndex(parent);
						let formula = cellListeners[i].Formula;

						if (parent.name) {
							continue;
						}

						if (formula.includes(":") && !cellListeners[i].is3D) {
							// call getAllAreaIndexes which return cellIndexes of each element(this will be parentCellIndex)
							let areaIndexes = getAllAreaIndexes(cellListeners[i]);
							if (areaIndexes) {
								// go through the values and set dependents for each
								for (let index of areaIndexes) {
									this._setDependents(cellIndex, index);
								}
								continue;
							}
						}

						// if the child cell does not yet have a dependency with listeners, create it
						if (!this._getDependents(cellIndex, elemCellIndex)) {
							this._setDependents(cellIndex, elemCellIndex);
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
		let isFirstCall;
		if (!this.precedents) {
			this.precedents = {};
			isFirstCall = true;
		}
		if (!this.precedentsAreas) {
			this.precedentsAreas = {};
		}

		let t = this;
		let currentCellIndex = AscCommonExcel.getCellIndex(row, col);
		if (!this.precedents[currentCellIndex]) {
			let isFirstInSharedArea = true;
			let shared, base;
			if (formulaParsed.shared !== null) {
				// base cellIndex + diff between base and current
				shared = formulaParsed.getShared();
				base = shared.base;
			}
			
			if (formulaParsed.outStack && isFirstInSharedArea) {
				let currentWsIndex = formulaParsed.ws.index;
				// iterate and find all reference
				for (const elem of formulaParsed.outStack) {
					let elemType = elem.type ? elem.type : null;
					if (elemType === AscCommonExcel.cElementType.cell || elemType === AscCommonExcel.cElementType.cellsRange || 
						elemType === AscCommonExcel.cElementType.cell3D || elemType === AscCommonExcel.cElementType.cellsRange3D ||
						elemType === AscCommonExcel.cElementType.name || elemType === AscCommonExcel.cElementType.name3D) {
						const areaRange = {};
						let is3D = elemType === AscCommonExcel.cElementType.cell3D || elemType === AscCommonExcel.cElementType.cellsRange3D || elemType === AscCommonExcel.cElementType.name3D;
						let isArea = elemType === AscCommonExcel.cElementType.cellsRange || elemType === AscCommonExcel.cElementType.name;
						let elemRange, elemCellIndex;
						if (elemType === AscCommonExcel.cElementType.name || elemType === AscCommonExcel.cElementType.name3D) {
							let elemValue = elem.getValue();
							elemRange = elemValue.range.bbox ? elemValue.range.bbox : elemValue.bbox;
							if (elemRange.isOneCell()) {
								isArea = false;
							}
						} else {
							elemRange = elem.range.bbox ? elem.range.bbox : elem.bbox;
						}

						if (base) {
							elemCellIndex = AscCommonExcel.getCellIndex(elemRange.r1 + (row - base.nRow), elemRange.c1 + (col - base.nCol));
						} else {
							elemCellIndex = AscCommonExcel.getCellIndex(elemRange.r1, elemRange.c1);
						}
						
						let elemDependentsIndex = elemCellIndex;
						let currentDependentsCellIndex = currentCellIndex;
	

						if (isArea) {
							// write top left and bottom right cell indexes to draw a rectangle around the range
							const tempObj = {};
							const areaName = elem.value;	// areaName - unique key for areaRange
							tempObj.topLeftIndex = AscCommonExcel.getCellIndex(elemRange.r1, elemRange.c1);
							tempObj.bottomRightIndex = AscCommonExcel.getCellIndex(elemRange.r2, elemRange.c2);
							areaRange[areaName] = tempObj;

							this._setPrecedentsAreas(areaRange);
						}

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
					}
				}
			}
		} else if (!this.getPrecedentsLoop()) {
			const currentPrecedent = Object.assign({}, this.precedents);
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
