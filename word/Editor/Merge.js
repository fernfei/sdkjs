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

(function (undefined) {
    const CDocumentComparison = AscCommonWord.CDocumentComparison;
    const CNode = AscCommonWord.CNode;
    const CTextElement = AscCommonWord.CTextElement;

	function insertBookmarkAndContinue(oBookmarkChange, oTextIterator, oBookmarkIterator)
	{
		const bRet = oTextIterator.skipTo(oBookmarkChange.elementIndex, oBookmarkChange.innerElementIndex - oBookmarkChange.addingValue);
		if (bRet)
		{
			const oRun = oTextIterator.splitCurrentRun(oTextIterator.runElementIndex + oBookmarkChange.addingValue);
			oTextIterator.addToCollectBack(oRun);
			if (oBookmarkChange.elementIndex === 0 && oBookmarkChange.innerElementIndex === 0)
			{
				changeFirstTextElement(oTextIterator, oRun);
			}
			for (let i = 0; i < oBookmarkChange.bookmarks.length; i += 1)
			{
				oTextIterator.parent.AddToContent(oTextIterator.runIndex + 1, oBookmarkChange.bookmarks[i]);
			}
			if (oBookmarkIterator.check())
			{
				oBookmarkIterator.next();
				return oBookmarkIterator.value();
			}
		}
	}
	function applyEndChangeReview(oReviewChange, oTextIterator)
	{
		const bRet = oTextIterator.skipTo(oReviewChange.endElementIndex, oReviewChange.endInnerElementIndex - 1);
		if (bRet)
		{
			oTextIterator.startCollectRuns();
			oTextIterator.addToCollectCurrentRun();
			oTextIterator.splitCurrentRun(oTextIterator.runElementIndex + 1);
		}
	}
	function changeFirstTextElement(oTextIterator, oRun)
	{
		const oElement = oTextIterator.getCurrentElement();
		oElement.elements[0] = oRun.Content[0];
		oElement.firstRun = oRun;
	}
	function applyStartChangeReview(oReviewChange, oTextIterator, oChangesIterator, comparison, oNeedReviewWithUser)
	{

		const bRet = oTextIterator.skipTo(oReviewChange.startElementIndex, oReviewChange.startInnerElementIndex);
		if (bRet)
		{
			const oRun = oTextIterator.splitCurrentRun();
			oTextIterator.dropLastCollect();
			oTextIterator.addToCollectBack(oRun);
			if (oReviewChange.startElementIndex === 0 && oReviewChange.startInnerElementIndex === 0)
			{
				changeFirstTextElement(oTextIterator, oRun);
			}


			const arrRuns = oTextIterator.endCollectRuns();

			const nPriorityReviewType = oReviewChange.reviewType;
			const nPriorityMoveReviewType = oReviewChange.moveReviewType;
			const sMoveReviewMarkName = oReviewChange.moveReviewMarkName;
			const oReviewInfo = oReviewChange.reviewInfo;
			const sReviewUserName = oReviewInfo.GetUserName();
			const sReviewDate = oReviewInfo.GetDateTime();


			if (!oNeedReviewWithUser[sReviewDate]) {
				oNeedReviewWithUser[sReviewDate] = {};
			}

			if (!oNeedReviewWithUser[sReviewDate][sReviewUserName]) {
				const oNeedReview = {reviewTypes: {}, moveReviewTypes: {}};
				oNeedReview.reviewTypes[reviewtype_Add] = [];
				oNeedReview.reviewTypes[reviewtype_Remove] = [];
				oNeedReview.moveReviewTypes[Asc.c_oAscRevisionsMove.MoveTo] = [];
				oNeedReview.moveReviewTypes[Asc.c_oAscRevisionsMove.MoveFrom] = [];
				oNeedReviewWithUser[sReviewDate][sReviewUserName] = oNeedReview;
			}

			if (AscFormat.isRealNumber(nPriorityReviewType)) {
				for (let i = 0; i < arrRuns.length; i += 1)
				{
					oNeedReviewWithUser[sReviewDate][sReviewUserName].reviewTypes[nPriorityReviewType].push({element: arrRuns[i], reviewInfo: oReviewInfo});
				}
			}
			else if (AscFormat.isRealNumber(nPriorityMoveReviewType))
			{
				for (let i = 0; i < arrRuns.length; i += 1)
				{
					comparison.oComparisonMoveMarkManager.addMoveMarkNameRunRelation(sMoveReviewMarkName, arrRuns[i]);
					comparison.oComparisonMoveMarkManager.addRunMoveMarkNameRelation(sMoveReviewMarkName, arrRuns[i]);
					oNeedReviewWithUser[sReviewDate][sReviewUserName].moveReviewTypes[nPriorityMoveReviewType].push({element: arrRuns[i], reviewInfo: oReviewInfo});												}
			}
			if (oChangesIterator.check())
			{
				oChangesIterator.next();
				return oChangesIterator.value();
			}
		}
	}

    function getPriorityReviewType(arrOfTypes) {
        const bRemove = arrOfTypes.some(function (reviewType) {
            return reviewType === reviewtype_Remove;
        });
        if (bRemove) return reviewtype_Remove;

        const bAdd = arrOfTypes.some(function (reviewType) {
            return reviewType === reviewtype_Add;
        });
        if (bAdd) return reviewtype_Add;
        return reviewtype_Common;
    }

    function checkArrayForReviewType(arr) {
        for (let i = 0; i < 0; i += 1) {
            if (arr[i].GetReviewType && (arr[i].GetReviewType() !== reviewtype_Common)) {
                return true;
            }
        }
        return false;
    }
	function CReviewChange()
	{
		this.startElementIndex = -1;
		this.startInnerElementIndex = -1;

		this.endElementIndex = -1;
		this.endInnerElementIndex = -1;

		this.reviewInfo = null;

		this.reviewType = null;

		this.moveReviewType = null
		this.moveReviewMarkName = null;
	}
	CReviewChange.prototype.setStart = function (nElementIndex, nInnerElementIndex)
	{
		this.startElementIndex = nElementIndex;
		this.startInnerElementIndex = nInnerElementIndex;
	}
	CReviewChange.prototype.setEnd = function (nElementIndex, nInnerElementIndex)
	{
		this.endElementIndex = nElementIndex;
		this.endInnerElementIndex = nInnerElementIndex;
	}
	CReviewChange.prototype.setMoveReviewType = function (nMoveReviewType, sMoveReviewMarkName)
	{
		this.moveReviewType = nMoveReviewType;
		this.moveReviewMarkName = sMoveReviewMarkName;
	}
	CReviewChange.prototype.setReviewInfo = function (oReviewInfo)
	{
		this.reviewInfo = oReviewInfo;
	}
	CReviewChange.prototype.setReviewType = function (nReviewType)
	{
		this.reviewType = nReviewType;
	}

	function CBookmarkChange(arrBookmarks, nElementIndex, nInnerElementIndex, nAddingValue)
	{
		this.bookmarks = arrBookmarks;
		this.elementIndex = nElementIndex;
		this.innerElementIndex = nInnerElementIndex;
		this.addingValue = nAddingValue;
	}
	function CBookmarkChangesIterator(arrElements)
	{
		this.elements = arrElements;
		this.currentBookmark = null;
		this.nextBookmark = null;
		this.elementIndex = arrElements.length - 1;
		this.elementBookmarkIndexes = this.getElement().getBookmarkInsertIndexes();
		this.innerBookmarkIndex = -1;
		this.next();
		this.next();
	}
	CBookmarkChangesIterator.prototype.check = function ()
	{
		return !!this.nextBookmark;
	}
	CBookmarkChangesIterator.prototype.value = function ()
	{
		return this.currentBookmark;
	}
	CBookmarkChangesIterator.prototype.getInsertIndex = function ()
	{
		return this.elementBookmarkIndexes[this.innerBookmarkIndex];
	}
	CBookmarkChangesIterator.prototype.getAddingValue = function ()
	{
		const nInsertIndex = this.getInsertIndex();
		return nInsertIndex === 0 ? 0 : 1;
	}
	CBookmarkChangesIterator.prototype.next = function ()
	{
		this.currentBookmark = this.nextBookmark;
		this.nextBookmark = null;
		this.innerBookmarkIndex += 1;
		while (this.innerBookmarkIndex === this.elementBookmarkIndexes.length && this.elementIndex > 0)
		{
			this.elementIndex -= 1;
			this.elementBookmarkIndexes = this.getElement().getBookmarkInsertIndexes();
			this.innerBookmarkIndex = 0;
		}
		const oElement = this.getElement();
		if (oElement && this.innerBookmarkIndex < this.elementBookmarkIndexes.length)
		{
			const nInsertIndex = this.getInsertIndex();
			this.nextBookmark = new CBookmarkChange(oElement.bookmarks[nInsertIndex], this.elementIndex, nInsertIndex, this.getAddingValue());
		}
	}

	CBookmarkChangesIterator.prototype.getElement = function ()
	{
		return this.elements[this.elementIndex];
	}

	function CReviewChangesIterator(arrMainElements, arrRevisedElements)
	{
		this.mainElements = arrMainElements;
		this.revisedElements = arrRevisedElements;

		this.currentChange = null;
		this.nextChange = null;
		this.elementIndex = null;
		this.innerElementIndex = null;

		this.canNextElement = true;
		this.init();
	}

	CReviewChangesIterator.prototype.init = function ()
	{
		this.elementIndex = this.mainElements.length - 1;
		this.innerElementIndex = this.getMainElement().elements.length;
		this.findFirstChange();
	};
	CReviewChangesIterator.prototype.check = function ()
	{
		return !!this.currentChange;
	};

	CReviewChangesIterator.prototype.nextElement = function ()
	{
		this.innerElementIndex -= 1;
		if (this.innerElementIndex < 0)
		{
			this.elementIndex -= 1;
			this.innerElementIndex = 0;
		}
		if (this.elementIndex === 0 && this.innerElementIndex === 0)
		{
			this.canNextElement = false;
		}
	}

	CReviewChangesIterator.prototype.value = function ()
	{
		return this.currentChange;
	};
	CReviewChangesIterator.prototype.next = function ()
	{
		this.currentChange = this.nextChange;
		this.nextChange = null;
		while (this.canNextElement)
		{
			this.nextElement();
			if (this.endCurrentAndFindNextChange())
			{
				break;
			}
		}
	};
	CReviewChangesIterator.prototype.getMainElement = function ()
	{
		return this.mainElements[this.elementIndex];
	}
	CReviewChangesIterator.prototype.getRevisedElement = function ()
	{
		return this.revisedElements[this.elementIndex];
	}
	CReviewChangesIterator.prototype.findFirstChange = function ()
	{
		this.isFindFirstChange = true;
		this.next();
		this.isFindFirstChange = false;
		this.next();
	}
	CReviewChangesIterator.prototype.getPriorityChange = function ()
	{
		const oRevisedElement = this.getRevisedElement();
		const oMainElement = this.getMainElement();
		const oMainReviewElement = oMainElement.reviewElementTypes[this.innerElementIndex];
		const oRevisedReviewElement = oRevisedElement.reviewElementTypes[this.innerElementIndex];

		const nRevisedReviewType = oRevisedReviewElement.reviewType;
		const nRevisedMoveReviewType = oRevisedReviewElement.moveReviewType;
		const sRevisedNameMoveMark = oRevisedReviewElement.moveName;
		const oRevisedReviewInfo = oRevisedReviewElement.reviewInfo;
		const oRevisedPrevAdded = oRevisedReviewElement.prevAdded;

		const nMainReviewType = oMainReviewElement.reviewType;
		const oMainPrevAdded = oMainReviewElement.prevAdded;
		const nMainMoveReviewType = oMainReviewElement.moveReviewType;

		let nPriorityReviewType;
		let oPriorityReviewInfo;
		let nPriorityMoveReviewType;
		let bIsMove = false;
		if (nRevisedReviewType !== reviewtype_Common && !(oMainPrevAdded && nMainReviewType === reviewtype_Remove)) {
			if (nMainReviewType !== nRevisedReviewType) {
				nPriorityReviewType = nRevisedReviewType;
				oPriorityReviewInfo = oRevisedReviewInfo;
			} else if (nMainReviewType === reviewtype_Remove && oRevisedPrevAdded) {
				nPriorityReviewType = reviewtype_Add;
				oPriorityReviewInfo = oRevisedPrevAdded;
			}
		}
		if (nMainMoveReviewType === Asc.c_oAscRevisionsMove.NoMove && nRevisedMoveReviewType !== Asc.c_oAscRevisionsMove.NoMove) {
			nPriorityMoveReviewType = nRevisedMoveReviewType;
			oPriorityReviewInfo = oRevisedReviewInfo;
			bIsMove = true;
		}

		if ((AscFormat.isRealNumber(nPriorityReviewType) || AscFormat.isRealNumber(nPriorityMoveReviewType)) && oPriorityReviewInfo)
		{
			const oChange = new CReviewChange();
			oChange.setReviewInfo(oPriorityReviewInfo);
			if (bIsMove)
			{
				oChange.setMoveReviewType(nPriorityMoveReviewType, sRevisedNameMoveMark);
			}
			else
			{
				oChange.setReviewType(nPriorityReviewType);
			}
			return oChange;
		}
	}
	CReviewChangesIterator.prototype.endCurrentAndFindNextChange = function ()
	{
		const oPriorityChange = this.getPriorityChange();
		if (oPriorityChange)
		{
			if (this.updateCurrentChange(oPriorityChange))
			{
				return false;
			}
			else
			{
				this.nextChange = oPriorityChange;
				oPriorityChange.setStart(this.elementIndex, this.innerElementIndex);
				oPriorityChange.setEnd(this.elementIndex, this.innerElementIndex + 1);
				return true;
			}
		}
		return false;
	}
	CReviewChangesIterator.prototype.updateCurrentChange = function (oPriorityChange)
	{
		if (this.isFindFirstChange || !this.currentChange)
		{
			return false;
		}

		const bIsMove = oPriorityChange.moveReviewType !== null;

		const bCheckReviewType = bIsMove ? oPriorityChange.moveReviewType === this.currentChange.moveReviewType : oPriorityChange.reviewType === this.currentChange.reviewType;

		let bIsNextCheckElement = false;
		if (this.innerElementIndex === this.getMainElement().elements.length - 1)
		{
			if ((this.currentChange.startElementIndex === this.elementIndex + 1) &&
				(this.currentChange.startInnerElementIndex === 0))
			{
				bIsNextCheckElement = true;
			}
		}
		else
		{
			if ((this.currentChange.startElementIndex === this.elementIndex) &&
				(this.currentChange.startInnerElementIndex === this.innerElementIndex + 1))
			{
				bIsNextCheckElement = true;
			}
		}
		const oPriorityReviewInfo = oPriorityChange.reviewInfo;
		const sReviewUserName = oPriorityReviewInfo.GetUserName();
		const nReviewDate = oPriorityReviewInfo.GetDateTime();
		if (bIsNextCheckElement
			&& bCheckReviewType
			&& this.currentChange.reviewInfo
			&& (this.currentChange.reviewInfo.GetUserName() === sReviewUserName )
			&& this.currentChange.reviewInfo.GetDateTime() === nReviewDate)
		{
			this.currentChange.setStart(this.elementIndex, this.innerElementIndex);
			return true;
		}
		return false;
	}

    function CMergeComparisonNode(oElement, oParent) {
        CNode.call(this, oElement, oParent);
        this.bHaveMoveMarks = false;
    }

    CMergeComparisonNode.prototype = Object.create(CNode.prototype);
    CMergeComparisonNode.prototype.constructor = CMergeComparisonNode;

    CMergeComparisonNode.prototype.privateCompareElements = function (oNode, bCheckNeighbors) {
        const oElement1 = this.element;
        const oElement2 = oNode.element;
        if (oElement1.isReviewWord !== oElement2.isReviewWord) {
            return false;
        }
        return CNode.prototype.privateCompareElements.call(this, oNode, bCheckNeighbors);
    }

    CMergeComparisonNode.prototype.copyRunWithMockParagraph = function (oRun, mockParagraph, comparison) {
        const oRet = CNode.prototype.copyRunWithMockParagraph.call(this, oRun, mockParagraph, comparison);
        return oRet;
    };

    CMergeComparisonNode.prototype.setCommonReviewTypeWithInfo = function (element, info) {
        element.SetReviewTypeWithInfo((element.GetReviewType && element.GetReviewType()) || reviewtype_Common, info);
    };

    CMergeComparisonNode.prototype.applyInsert = function (arrToInsert, arrToRemove, nInsertPosition, comparison, opts) {
        opts = opts || {};
        if (arrToInsert.length === 0) {
            for (let i = 0; i < arrToRemove.length; i += 1) {
                comparison.setRemoveReviewType(arrToRemove[i]);
            }
        } else if (arrToRemove.length === 0) {
            this.insertContentAfterRemoveChanges(arrToInsert, nInsertPosition, comparison);
        } else {
            arrToInsert = arrToInsert.reverse();
            if (opts.needReverse) {
                arrToRemove = arrToRemove.reverse();
            }
            nInsertPosition = arrToRemove[0].GetPosInParent();
            comparison.resolveConflicts(arrToInsert, arrToRemove, arrToRemove[0].Paragraph, nInsertPosition);
        }
    }

    CMergeComparisonNode.prototype.insertContentAfterRemoveChanges = function (aContentToInsert, nInsertPosition, comparison) {
        const oElement = this.getApplyParagraph(comparison);
        if(nInsertPosition > -1)
        {
            for (let t = 0; t < aContentToInsert.length; t += 1) {
                if(comparison.isElementForAdd(aContentToInsert[t]))
                {
                    if (aContentToInsert[t] instanceof AscCommon.CParaRevisionMove) {
                        const oInsertBefore = oElement.Content[nInsertPosition];
                        if (oInsertBefore) {
                            comparison.oComparisonMoveMarkManager.addRevisedMoveMarkToInserts(aContentToInsert[t], oInsertBefore, oElement, false);
                        }
                    } else {
                        const sMoveName = comparison.oComparisonMoveMarkManager.getMoveMarkNameByRun(aContentToInsert[t]);
                        comparison.oComparisonMoveMarkManager.addMoveMarkNameRunRelation(sMoveName, aContentToInsert[t]);
                        oElement.AddToContent(nInsertPosition, aContentToInsert[t]);
                    }

                }
            }
        }
    };

    function CMergeComparisonTextElement() {
        CTextElement.call(this);
        this.isReviewWord = false;
        this.reviewElementTypes = [];
    }

    CMergeComparisonTextElement.prototype = Object.create(CTextElement.prototype);
    CMergeComparisonTextElement.prototype.constructor = CMergeComparisonTextElement;

    CMergeComparisonTextElement.prototype.addToElements = function (element, reviewType) {
        CTextElement.prototype.addToElements.call(this, element);
        this.reviewElementTypes.push(reviewType);
        if (reviewType.reviewType !== reviewtype_Common || reviewType.moveReviewType !== Asc.c_oAscRevisionsMove.NoMove) {
            this.isReviewWord = true;
        }
    };

	function compareReviewElements(reviewElement1, reviewElement2)
	{
		return (
			reviewElement1.reviewType === reviewtype_Common && reviewElement1.reviewType === reviewElement2.reviewType &&
			reviewElement1.moveReviewType === Asc.c_oAscRevisionsMove.NoMove && reviewElement1.moveReviewType === reviewElement2.moveReviewType ||
			reviewElement1.moveReviewType === reviewElement2.moveReviewType &&
			reviewElement1.reviewType === reviewElement2.reviewType &&
			reviewElement1.reviewInfo && reviewElement2.reviewInfo && reviewElement1.reviewInfo.IsEqual(reviewElement2.reviewInfo, true)
		);
	}
	CMergeComparisonTextElement.prototype.compareReviewElements = function (oAnotherElement)
	{
		if (this.reviewElementTypes.length === oAnotherElement.reviewElementTypes.length) {
			for (let i = 0; i < this.reviewElementTypes.length; i += 1) {
				const bNotEqualsReviewTypes = !compareReviewElements(this.reviewElementTypes[i], oAnotherElement.reviewElementTypes[i]);
				if (bNotEqualsReviewTypes) {
					return false;
				}
			}
		}
		else
		{
			return false;
		}
		return true;
	};
    CMergeComparisonTextElement.prototype.equals = function (oOtherElement, bNeedCheckTypes) {
        const bEquals = CTextElement.prototype.equals.call(this, oOtherElement);
        if (!bEquals) {
            return false;
        }
				if (bNeedCheckTypes)
				{
					const bCheck = this.compareReviewElements(oOtherElement);
					if (!bCheck)
					{
						return false;
					}
				}

        return true;
    }

    function CResolveConflictTextElement() {
        CTextElement.call(this);
        this.reviewElementTypes = [];
    }

    CResolveConflictTextElement.prototype = Object.create(CTextElement.prototype);
    CResolveConflictTextElement.prototype.constructor = CResolveConflictTextElement;
    CResolveConflictTextElement.prototype.addToElements = CMergeComparisonTextElement.prototype.addToElements;

    CResolveConflictTextElement.prototype.checkRemoveReviewType = function (nIndex)
    {
        const oReviewInformation = this.reviewElementTypes[nIndex];
        if (oReviewInformation)
        {
            return oReviewInformation.reviewType === reviewtype_Remove;
        }
        return false;
    };

    CResolveConflictTextElement.prototype.isWordBeginWith = function (oOther)
    {
        if (this.elements.length < oOther.elements.length) {
            return false;
        }

        for (let i = 0; i < oOther.elements.length; i += 1)
        {
            const oMainElement = this.elements[i];
            const oSecondaryElement = oOther.elements[i];
            if (typeof oMainElement.Value !== 'number' || typeof oSecondaryElement.Value !== 'number')
            {
                return false;
            } else if (oMainElement.Value !== oSecondaryElement.Value) {
                return false;
            }
        }
        return true;
    };
    CResolveConflictTextElement.prototype.isWordEndWith = function (oOther)
    {
        if (this.elements.length < oOther.elements.length) {
            return false;
        }

        for (let i = 0; i < oOther.elements.length; i += 1)
        {
            const oMainElement = this.elements[this.elements.length - 1 - i];
            const oSecondaryElement = oOther.elements[oOther.elements.length - 1 - i];
            if (typeof oMainElement.Value !== 'number' || typeof oSecondaryElement.Value !== 'number')
            {
                return false;
            } else if (oMainElement.Value !== oSecondaryElement.Value) {
                return false;
            }
        }
        return true;
    };

    CResolveConflictTextElement.prototype.equals = function (other, bNeedCheckReview)
    {
        const bResult = CTextElement.prototype.equals.call(this, other);
        if (bResult || this.elements.length === other.elements.length) {
            return bResult;
        }
        let oMainTextElement;
        let oSecondaryTextElement;
        if (this.elements.length > other.elements.length)
        {
            oMainTextElement = this;
            oSecondaryTextElement = other;
        } else {
            oMainTextElement = other;
            oSecondaryTextElement = this;
        }
        let bCheckStart = false;
        let bCheckEnd = false;
        if (oMainTextElement.checkRemoveReviewType(oSecondaryTextElement.elements.length - 1)) {
            bCheckStart = oMainTextElement.isWordBeginWith(oSecondaryTextElement);
        }
        if (oMainTextElement.checkRemoveReviewType(oMainTextElement.elements.length - oSecondaryTextElement.elements.length)) {
            bCheckEnd = oMainTextElement.isWordEndWith(oSecondaryTextElement);
        }

        return bCheckStart || bCheckEnd;
    };
    

    function CDocumentResolveConflictComparison(oOriginalDocument, oRevisedDocument, oOptions) {
        CDocumentComparison.call(this, oOriginalDocument, oRevisedDocument, oOptions);
        this.parentParagraph = null;
        this.startPosition = 0;
        this.bSkipChangeMoveType = true;
        this.copyPr = {
            CopyReviewPr: false,
            Comparison: this,
        };
        this.bSaveCustomReviewType = true;
    }
    CDocumentResolveConflictComparison.prototype = Object.create(CDocumentComparison.prototype);
    CDocumentResolveConflictComparison.prototype.constructor = CDocumentResolveConflictComparison;

    CDocumentResolveConflictComparison.prototype.getNodeConstructor = function () {
        return CConflictResolveNode;
    }
    CDocumentResolveConflictComparison.prototype.checkOriginalAndSplitRun = function (oOriginalRun, oSplitRun) {
        const sName = this.oComparisonMoveMarkManager.getMoveMarkNameByRun(oOriginalRun);
        this.oComparisonMoveMarkManager.addMoveMarkNameRunRelation(sName, oSplitRun);
        this.oComparisonMoveMarkManager.addRunMoveMarkNameRelation(sName, oSplitRun);
    };

    CDocumentResolveConflictComparison.prototype.setReviewInfoForArray = function (arrNeedReviewObjects, nType) {
        for (let i = 0; i < arrNeedReviewObjects.length; i += 1) {
            const oNeedReviewObject = arrNeedReviewObjects[i];
            if (oNeedReviewObject.SetReviewTypeWithInfo) {
                let oReviewInfo = oNeedReviewObject.ReviewInfo.Copy();
                this.setReviewInfo(oReviewInfo);
                if (this.bSaveCustomReviewType) {
                    const reviewType = oNeedReviewObject.GetReviewType && oNeedReviewObject.GetReviewType();
                    if (reviewType === reviewtype_Add || reviewType === reviewtype_Remove) {
                        if (nType === reviewtype_Add && reviewType === reviewtype_Remove) {
                            oReviewInfo = oNeedReviewObject.ReviewInfo.Copy();
                            oReviewInfo.SavePrev(reviewtype_Add);
                            nType = reviewtype_Remove;
                        } else if (reviewType === reviewtype_Add && nType === reviewtype_Remove) {
                            oReviewInfo = oNeedReviewObject.ReviewInfo.Copy();
                            oReviewInfo.SavePrev(reviewtype_Add);
                        }
                    }
                }
                oNeedReviewObject.SetReviewTypeWithInfo(nType, oReviewInfo, false);
            }
        }
    }

    CDocumentResolveConflictComparison.prototype.getTextElementConstructor = function () {
        return CResolveConflictTextElement;
    }

    CDocumentResolveConflictComparison.prototype.getCompareReviewInfo = function (oRun) {
        const oReviewInfo = oRun.GetReviewInfo && oRun.GetReviewInfo();
        const prevAdded = oReviewInfo.GetPrevAdded();
        const reviewType = oRun.GetReviewType && oRun.GetReviewType();
        const moveReviewType = oRun.GetReviewMoveType && oRun.GetReviewMoveType();
        const bNotRunIdInit = !this.oComparisonMoveMarkManager.getMoveMarkNameByRun(oRun);
        if (moveReviewType !== Asc.c_oAscRevisionsMove.NoMove && bNotRunIdInit && this.oComparisonMoveMarkManager.moveMarksStack.length) {
            const oCurrentMoveMark = this.oComparisonMoveMarkManager.moveMarksStack[this.oComparisonMoveMarkManager.moveMarksStack.length - 1];
            this.oComparisonMoveMarkManager.addRunMoveMarkNameRelation(oCurrentMoveMark.Name, oRun);
        }
        return {
            reviewType: reviewType,
            moveReviewType: moveReviewType,
            moveName: this.oComparisonMoveMarkManager.getMoveMarkNameByRun(oRun),
            reviewInfo: oReviewInfo,
            prevAdded: prevAdded
        };
    }

    CDocumentResolveConflictComparison.prototype.applyChangesToParagraph = function (oNode) {
        oNode.changes.sort(function (c1, c2) {
            return c2.anchor.index - c1.anchor.index;
        });
        let currentChangeId = 0;
        for (let i = oNode.children.length - 1; i >= 0; i -= 1) {
            const oChildNode = oNode.children[i];
            if (i !== oNode.children.length - 1) {
                oChildNode.tryUpdateNode(this);
                oChildNode.resolveTypesWithPartner(this);
            }
            if (currentChangeId < oNode.changes.length && oNode.changes[currentChangeId].anchor.index === i) {
                const aContentToInsert = oNode.getArrOfInsertsFromChanges(currentChangeId, this);
                //handle removed elements
                oNode.applyInsertsToParagraph(this, aContentToInsert, currentChangeId);
                currentChangeId += 1
            }
            if (currentChangeId < oNode.changes.length && oNode.changes[currentChangeId].anchor.index > i) {
                currentChangeId += 1;
            }
        }

        this.applyChangesToChildrenOfParagraphNode(oNode);
        this.applyChangesToSectPr(oNode);
    }

    CDocumentResolveConflictComparison.prototype.getLCSEqualsMethod = function () {
        return function () {
            return true;
        }
    }

    CDocumentResolveConflictComparison.prototype.setRemoveReviewType = function (element) {
        if (!(element.IsParaEndRun && element.IsParaEndRun())) {
            if (!element.GetReviewType || element.GetReviewType && element.GetReviewType() === reviewtype_Common) {
                this.setReviewInfoRecursive(element, this.nRemoveChangesType);
            }
        }
    }

    CDocumentResolveConflictComparison.prototype.resolveCustomReviewTypesBetweenElements = function (oMainElement, nRevisedReviewType, oRevisedReviewInfo) {
        const nMainReviewType = oMainElement.GetReviewType();
        if (nRevisedReviewType !== reviewtype_Common && nRevisedReviewType !== nMainReviewType) {
            const oMainReviewInfo = oMainElement.GetReviewInfo().Copy();
            oRevisedReviewInfo = oRevisedReviewInfo.Copy();
            if (nMainReviewType === reviewtype_Common) {
                oMainElement.SetReviewTypeWithInfo(nRevisedReviewType, oRevisedReviewInfo);
            } else if (nMainReviewType === reviewtype_Add) {
                oRevisedReviewInfo.SetPrevReviewTypeWithInfoRecursively(reviewtype_Add, oMainReviewInfo);
                oMainElement.SetReviewTypeWithInfo(reviewtype_Remove, oRevisedReviewInfo);
            } else if (nMainReviewType === reviewtype_Remove) {
                oMainReviewInfo.SetPrevReviewTypeWithInfoRecursively(reviewtype_Add, oRevisedReviewInfo);
                oMainElement.SetReviewTypeWithInfo(reviewtype_Remove, oMainReviewInfo);
            }
        }
    }
	CDocumentResolveConflictComparison.prototype.applyResolveTypes = function (oNeedReviewWithUser)
	{
		for (let sReviewDate in oNeedReviewWithUser)
		{
			for (let sUserName in oNeedReviewWithUser[sReviewDate])
			{
				for (let i = 0; i < oNeedReviewWithUser[sReviewDate][sUserName].reviewTypes[reviewtype_Add].length; i += 1)
				{
					const info = oNeedReviewWithUser[sReviewDate][sUserName].reviewTypes[reviewtype_Add][i];
					const element = info.element;
					const reviewInfo = info.reviewInfo;
					this.resolveCustomReviewTypesBetweenElements(element, reviewtype_Add, reviewInfo);

				}
				for (let i = 0; i < oNeedReviewWithUser[sReviewDate][sUserName].reviewTypes[reviewtype_Remove].length; i += 1)
				{
					const info = oNeedReviewWithUser[sReviewDate][sUserName].reviewTypes[reviewtype_Remove][i];
					const element = info.element;
					const reviewInfo = info.reviewInfo;
					this.resolveCustomReviewTypesBetweenElements(element, reviewtype_Remove, reviewInfo);
				}
				for (let i = 0; i < oNeedReviewWithUser[sReviewDate][sUserName].moveReviewTypes[Asc.c_oAscRevisionsMove.MoveFrom].length; i += 1)
				{
					const info = oNeedReviewWithUser[sReviewDate][sUserName].moveReviewTypes[Asc.c_oAscRevisionsMove.MoveFrom][i];
					const element = info.element;
					const reviewInfo = info.reviewInfo;
					const nOldReviewType = element.GetReviewType();
					if (nOldReviewType !== reviewtype_Common)
					{
						const oOldReviewInfo = element.GetReviewInfo().Copy();
						reviewInfo.SetPrevReviewTypeWithInfoRecursively(nOldReviewType, oOldReviewInfo);
					}
					element.SetReviewTypeWithInfo(reviewtype_Remove, reviewInfo);
				}
				for (let i = 0; i < oNeedReviewWithUser[sReviewDate][sUserName].moveReviewTypes[Asc.c_oAscRevisionsMove.MoveTo].length; i += 1)
				{
					const info = oNeedReviewWithUser[sReviewDate][sUserName].moveReviewTypes[Asc.c_oAscRevisionsMove.MoveTo][i];
					const element = info.element;
					const reviewInfo = info.reviewInfo;
					const nOldReviewType = element.GetReviewType();
					if (nOldReviewType !== reviewtype_Common)
					{
						const oOldReviewInfo = element.GetReviewInfo().Copy();
						reviewInfo.SetPrevReviewTypeWithInfoRecursively(nOldReviewType, oOldReviewInfo);
					}
					element.SetReviewTypeWithInfo(reviewtype_Add, reviewInfo);
				}
			}
		}
	}

    function CConflictResolveNode(oElement, oParent) {
        CNode.call(this, oElement, oParent);
    }

    CConflictResolveNode.prototype = Object.create(CNode.prototype);
    CConflictResolveNode.prototype.constructor = CConflictResolveNode;

    CConflictResolveNode.prototype.applyInsertsToParagraphsWithRemove = function (comparison, aContentToInsert, idxOfChange) {
        const arrSetRemoveReviewType = [];
        const infoAboutEndOfRemoveChange = this.prepareEndOfRemoveChange(idxOfChange, comparison, arrSetRemoveReviewType);
        const posLastRunInContent = infoAboutEndOfRemoveChange.posLastRunInContent;

        const nInsertPosition = infoAboutEndOfRemoveChange.nInsertPosition;
        this.setReviewTypeForRemoveChanges(comparison, idxOfChange, posLastRunInContent, nInsertPosition, arrSetRemoveReviewType);

        const nInsertPosition2 = arrSetRemoveReviewType[arrSetRemoveReviewType.length - 1].GetPosInParent();
        this.applyInsert(aContentToInsert, arrSetRemoveReviewType, nInsertPosition2, comparison, {needReverse: true});
    };

    // обновим ноды в любом случае, для дальнейшего разрешения типов
    CConflictResolveNode.prototype.tryUpdateNode = function (comparison) {
        const oPartnerNode = this.partner;
        if (oPartnerNode)
        {
            const oOriginalTextElement = this.element;
            const oPartnerTextElement = oPartnerNode.element;
            if (oPartnerTextElement.elements.length > oOriginalTextElement.elements.length) {
                const oNewOriginalTextElement = new CResolveConflictTextElement();
                oNewOriginalTextElement.firstRun = oOriginalTextElement.firstRun;
                oNewOriginalTextElement.lastRun = oOriginalTextElement.lastRun;

                const bIsWordBeginWithText = oPartnerTextElement.isWordBeginWith(oOriginalTextElement);
                const bIsWordEndWithText = oPartnerTextElement.isWordEndWith(oOriginalTextElement);

                const oParagraph = oOriginalTextElement.lastRun.Paragraph;
                if (bIsWordBeginWithText) {
                    for (let i = 0; i < oOriginalTextElement.elements.length; i += 1) {
                        oNewOriginalTextElement.addToElements(oOriginalTextElement.elements[i], oOriginalTextElement.reviewElementTypes[i]);
                    }
                }
                let nPreviousRunPosition;
                if (bIsWordBeginWithText || bIsWordEndWithText) {
                    this.element = oNewOriginalTextElement;
                    const oMockParagraph = oPartnerNode.par.element;
                    let nAmountOfAddingElements = oPartnerTextElement.elements.length - oOriginalTextElement.elements.length;
                    let nCurrentRunPosition = oPartnerTextElement.lastRun.GetPosInParent(oMockParagraph);
                    let oCurrentRun = oMockParagraph.Content[nCurrentRunPosition];
                    let nLastPartnerElementPosition = oCurrentRun.GetElementPosition(oPartnerTextElement.elements[oPartnerTextElement.elements.length - 1]);

                    if (bIsWordEndWithText) {
                        let nOffset = oOriginalTextElement.elements.length;
                        while (nOffset) {
                            if (nOffset - oCurrentRun.Content.length <= 0) {
                                nLastPartnerElementPosition = oCurrentRun.Content.length - nOffset - 1;
                                break;
                            }
                            nOffset -= oCurrentRun.Content.length;
                            nCurrentRunPosition -= 1;
                            oCurrentRun = oMockParagraph.Content[nCurrentRunPosition];
                        }
                    } else {
                        nLastPartnerElementPosition = oCurrentRun.GetElementPosition(oPartnerTextElement.elements[oPartnerTextElement.elements.length - 1]);
                    }
                    oCurrentRun.Split2(nLastPartnerElementPosition + 1);
                    const arrContentForInsert = [];
                    while (nAmountOfAddingElements) {
                        const oReviewInfo = comparison.getCompareReviewInfo(oCurrentRun);
                        for (let i = oCurrentRun.Content.length - 1; i >= 0; i -= 1) {
                            nAmountOfAddingElements -= 1;
                            if (nAmountOfAddingElements === 0) {
                                oCurrentRun = oCurrentRun.Split2(i);
                                break;
                            }
                        }
                        for (let i = 0; i < oCurrentRun.Content.length; i += 1) {
                            oNewOriginalTextElement.addToElements(oCurrentRun.Content[i], oReviewInfo);
                        }
                        arrContentForInsert.push(oCurrentRun.Copy(false, {CopyReviewPr: true}));
                        nCurrentRunPosition -= 1;
                        oCurrentRun = oMockParagraph.Content[nCurrentRunPosition];
                    }
                    let nLastOriginalElementPosition;
                    let nLastRunPosition;
                    if (bIsWordBeginWithText) {
                        nLastRunPosition = oOriginalTextElement.lastRun.GetPosInParent();
                        oNewOriginalTextElement.lastRun = arrContentForInsert[0];
                        nLastOriginalElementPosition = oParagraph.Content[nLastRunPosition].GetElementPosition(oOriginalTextElement.elements[oOriginalTextElement.elements.length - 1]);
                        oParagraph.Content[nLastRunPosition].Split2(nLastOriginalElementPosition + 1, oParagraph, nLastRunPosition)
                    } else {
                        nLastRunPosition = oOriginalTextElement.firstRun.GetPosInParent();
                        nPreviousRunPosition = nLastRunPosition + arrContentForInsert.length;
                        nLastOriginalElementPosition = oParagraph.Content[nLastRunPosition].GetElementPosition(oOriginalTextElement.elements[0]);
                        oParagraph.Content[nLastRunPosition].Split2(nLastOriginalElementPosition/* + 1*/, oParagraph, nLastRunPosition);
                        oNewOriginalTextElement.firstRun = arrContentForInsert[0];
                    }

                    //oParagraph.Content[nLastRunPosition].Split2(nLastOriginalElementPosition + 1, oParagraph, nLastRunPosition)
                    for (let i = 0; i < arrContentForInsert.length; i += 1) {
                        oParagraph.Add_ToContent(nLastRunPosition + 1, arrContentForInsert[i]);
                    }
                }

                if (bIsWordEndWithText && !bIsWordBeginWithText) {
                    let nElementsAmount = oOriginalTextElement.elements.length;
                    let nCurrentRunPosition = nPreviousRunPosition + 1;
                    let oCurrentRun = oParagraph.Content[nCurrentRunPosition];
                    while (nElementsAmount) {
                        const oReviewInfo = comparison.getCompareReviewInfo(oCurrentRun);
                        oNewOriginalTextElement.lastRun = oCurrentRun;
                        for (let i = 0; i < oCurrentRun.Content.length; i += 1) {
                            oNewOriginalTextElement.addToElements(oCurrentRun.Content[i], oReviewInfo);
                            nElementsAmount -= 1;
                            if (nElementsAmount === 0)
                            {
                                break;
                            }
                        }
                        nCurrentRunPosition += 1;
                        oCurrentRun = oParagraph.Content[nCurrentRunPosition];
                    }
                }
            } else if (oPartnerTextElement.elements.length < oOriginalTextElement.elements.length) {
                // здесь мы просто выравниваем количество элементов в ноде, чтобы разрешить остатки типов
                const bIsWordBeginWithText = oOriginalTextElement.isWordBeginWith(oPartnerTextElement);
                const bIsWordEndWithText = oOriginalTextElement.isWordEndWith(oPartnerTextElement);
                const oNewPartnerTextElement = new CResolveConflictTextElement();
                oNewPartnerTextElement.lastRun = oPartnerTextElement.lastRun;
                oNewPartnerTextElement.firstRun = oPartnerTextElement.firstRun;
                oPartnerNode.element = oNewPartnerTextElement;
                if (bIsWordBeginWithText) {
                    for (let i = 0; i < oPartnerTextElement.elements.length; i += 1) {
                        oNewPartnerTextElement.addToElements(oPartnerTextElement.elements[i], oPartnerTextElement.reviewElementTypes[[i]]);
                    }
                    for (let i = oPartnerTextElement.elements.length; i < oOriginalTextElement.elements.length; i += 1) {
                        oNewPartnerTextElement.addToElements(oOriginalTextElement.elements[i], oOriginalTextElement.reviewElementTypes[[i]]);
                    }
                } else if (bIsWordEndWithText) {
                    for (let i = 0; i < (oOriginalTextElement.elements.length - oPartnerTextElement.elements.length); i += 1) {
                        oNewPartnerTextElement.addToElements(oOriginalTextElement.elements[i], oOriginalTextElement.reviewElementTypes[[i]]);
                    }
                    for (let i = 0; i < oPartnerTextElement.elements.length; i += 1) {
                        oNewPartnerTextElement.addToElements(oPartnerTextElement.elements[i], oPartnerTextElement.reviewElementTypes[[i]]);
                    }
                }
            }
        }
    };

	CConflictResolveNode.prototype.resolveTypesWithPartner = function (comparison)
	{
		this.forEachRight(function (oNode)
		{
			const oPartnerNode = oNode.partner;
			const oOriginalTextElement = oNode.element;
			if (oOriginalTextElement instanceof CTextElement && oPartnerNode)
			{
				const oRevisedTextElement = oPartnerNode.element;
				const oNeedReviewWithUser = {};
				const oChangesIterator = new CReviewChangesIterator([oOriginalTextElement], [oRevisedTextElement]);
				const oBookmarkIterator = new CBookmarkChangesIterator([oRevisedTextElement]);
				const oTextIterator = new AscCommonWord.CTextElementRunIterator([oOriginalTextElement]);
				let oBookmarkChange = oBookmarkIterator.value();
				let oReviewChange = oChangesIterator.value();
				while (oBookmarkChange && oReviewChange)
				{
					if ((oReviewChange.endElementIndex >= oBookmarkChange.elementIndex) ||
						(oReviewChange.endElementIndex === oBookmarkChange.elementIndex && oReviewChange.endInnerElementIndex >= oBookmarkChange.innerElementIndex))
					{
						applyEndChangeReview(oReviewChange, oTextIterator);

						while (oBookmarkChange && ((oReviewChange.startElementIndex <= oBookmarkChange.elementIndex) ||
						(oReviewChange.startElementIndex === oBookmarkChange.elementIndex && oReviewChange.startInnerElementIndex <= oBookmarkChange.innerElementIndex))){
							oBookmarkChange = insertBookmarkAndContinue(oBookmarkChange, oTextIterator, oBookmarkIterator);
						}
						oReviewChange = applyStartChangeReview(oReviewChange, oTextIterator, oChangesIterator, comparison, oNeedReviewWithUser);
					}
					else
					{
						while (oBookmarkChange && (oBookmarkChange.elementIndex > oReviewChange.endElementIndex || oBookmarkChange.elementIndex === oReviewChange.endElementIndex && oBookmarkChange.innerElementIndex > oReviewChange.endInnerElementIndex))
						{
							oBookmarkChange = insertBookmarkAndContinue(oBookmarkChange, oTextIterator, oBookmarkIterator);
						}
					}
				}

				while (oBookmarkChange)
				{
					oBookmarkChange = insertBookmarkAndContinue(oBookmarkChange, oTextIterator, oBookmarkIterator);
				}
				while (oReviewChange)
				{
					applyEndChangeReview(oReviewChange, oTextIterator);
					oReviewChange = applyStartChangeReview(oReviewChange, oTextIterator, oChangesIterator, comparison, oNeedReviewWithUser);
				}
				comparison.applyResolveTypes(oNeedReviewWithUser);
			}
		});
	}

    CConflictResolveNode.prototype.applyInsertsToParagraphsWithoutRemove = function (comparison, aContentToInsert, idxOfChange) {
        const bRet = CNode.prototype.applyInsertsToParagraphsWithoutRemove.call(this, comparison, aContentToInsert, idxOfChange);
        if (!bRet) {
            const oChange = this.changes[idxOfChange];
            const applyingParagraph = this.getApplyParagraph(comparison);
            const index = oChange.anchor.index;
            if (index === this.children.length - 1) {

                const oLastConflictElement = this.children[this.children.length - 2].element;
                const nInsertIndex = oLastConflictElement.lastRun.GetPosInParent(applyingParagraph);
                const nLastSymbolPosition = oLastConflictElement.lastRun.GetElementPosition(oLastConflictElement.elements[oLastConflictElement.elements.length - 1]);
                if (nLastSymbolPosition !== -1) {
                    const oNewRun = oLastConflictElement.lastRun.Split2(nLastSymbolPosition + 1, applyingParagraph, nInsertIndex);
                    comparison.checkOriginalAndSplitRun(oNewRun, oLastConflictElement.lastRun);
                    this.applyInsert(aContentToInsert, [], nInsertIndex + 1, comparison);
                }
            }
        }
    }
    CConflictResolveNode.prototype.insertContentAfterRemoveChanges = CMergeComparisonNode.prototype.insertContentAfterRemoveChanges;

    CConflictResolveNode.prototype.getApplyParagraph = function (comparison) {
        return comparison.parentParagraph;
    }

    CConflictResolveNode.prototype.copyRunWithMockParagraph = function (oRun, mockParagraph, comparison) {
        comparison.copyPr.bSaveCustomReviewType = true;
        const oRet = CNode.prototype.copyRunWithMockParagraph.call(this, oRun, mockParagraph, comparison);
        delete comparison.copyPr.bSaveCustomReviewType;
        return oRet;
    };

    CConflictResolveNode.prototype.pushToArrInsertContentWithCopy = function (aContentToInsert, elem, comparison) {
            comparison.copyPr.bSaveCustomReviewType = true;
            CNode.prototype.pushToArrInsertContentWithCopy.call(this, aContentToInsert, elem, comparison);
            delete comparison.copyPr.bSaveCustomReviewType;
    }

    CConflictResolveNode.prototype.setCommonReviewTypeWithInfo = function (element, info) {
        element.SetReviewTypeWithInfo((element.GetReviewType && element.GetReviewType()) || reviewtype_Common, info);
    }
    
    CConflictResolveNode.prototype.getStartPosition = function (comparison) {
        return comparison.startPosition;
    }

    function CMockDocument() {
        this.Content = [];
    }

    function CMockParagraph() {
        this.Content = [];
    }

    function CMockMinHash() {
        this.count = 0;
        this.countLetters = 0;
    }

    CMockMinHash.prototype.jaccard = function () {
        return 0.8;
    }

    CMockMinHash.prototype.update = function () {
        this.count += 1;
    }

    function CDocumentMergeComparison(oOriginalDocument, oRevisedDocument, oOptions) {
        CDocumentComparison.call(this, oOriginalDocument, oRevisedDocument, oOptions);
        this.bSaveCustomReviewType = true;
        this.copyPr = {
            CopyReviewPr: false,
            Comparison: this,
            SkipUpdateInfo: true,
            CheckComparisonMoveMarks: true
        };
    }

    CDocumentMergeComparison.prototype = Object.create(CDocumentComparison.prototype);
    CDocumentMergeComparison.prototype.constructor = CDocumentMergeComparison;

    CDocumentMergeComparison.prototype.executeWithCheckInsertAndRemove = function (callback, oChange) {
        if (!oChange.remove.length || !oChange.insert.length) {
            const bOldSkipUpdateInfo = this.copyPr.SkipUpdateInfo;
            const bSaveCustomReviewType = this.copyPr.bSaveCustomReviewType;
            this.copyPr.SkipUpdateInfo = false;
            this.copyPr.bSaveCustomReviewType = true;
            callback();
            this.copyPr.SkipUpdateInfo = bOldSkipUpdateInfo;
            this.copyPr.bSaveCustomReviewType = bSaveCustomReviewType;
        } else {
            callback();
        }
    };

    CDocumentMergeComparison.prototype.checkOriginalAndSplitRun = CDocumentResolveConflictComparison.prototype.checkOriginalAndSplitRun;

    CDocumentMergeComparison.prototype.createNodeFromDocContent = function (oElement, oParentNode, oHashWords, isOriginalDocument) {
        this.oComparisonMoveMarkManager.resetMoveMarkStack();
        const oRet = CDocumentComparison.prototype.createNodeFromDocContent.call(this, oElement, oParentNode, oHashWords, isOriginalDocument);
        this.oComparisonMoveMarkManager.checkMoveMarksContentNode(oRet);
        return oRet;
    }
    CDocumentMergeComparison.prototype.checkCopyParagraphElement = function (oOldItem, oNewItem, arrMoveMarks) {
        if (para_RevisionMove === oOldItem.Type) {
            arrMoveMarks.unshift({moveMark: oNewItem, parentElement: this});
            return true;
        } else if (Array.isArray(oNewItem)) {
            const sMoveName = this.oComparisonMoveMarkManager.getMoveMarkNameByRun(oOldItem[0]);
            this.oComparisonMoveMarkManager.addRunMoveMarkNameRelation(sMoveName, oNewItem[0]);
            this.oComparisonMoveMarkManager.oRevisedMoveMarksInserts[oNewItem[0].Id] = arrMoveMarks;
        } else {
            const sMoveName = this.oComparisonMoveMarkManager.getMoveMarkNameByRun(oOldItem);
            this.oComparisonMoveMarkManager.addRunMoveMarkNameRelation(sMoveName, oNewItem);
            this.oComparisonMoveMarkManager.oRevisedMoveMarksInserts[oNewItem.Id] = arrMoveMarks;
        }
        return false;
    }
    CDocumentMergeComparison.prototype.correctMoveMarks = function () {
        const oRemoveMoveTypeNames = {};
        const oTrackRevisionManager = this.api.WordControl.m_oLogicDocument.TrackRevisionsManager;
        const oComparisonMoveMarkManager = this.oComparisonMoveMarkManager;
        const oMoveMarks = oTrackRevisionManager.MoveMarks;
        const oRevisedMoveMarksInserts = oComparisonMoveMarkManager.getRevisedMoveMarkToInserts();
        const oInsertMoveMarkId = oComparisonMoveMarkManager.oInsertMoveMarkId;
        const arrCheckMoveMarksElements = oComparisonMoveMarkManager.getCheckMoveMarkElements();

        for (let i = 0; i < arrCheckMoveMarksElements.length; i += 1) {
            const oRootElement = arrCheckMoveMarksElements[i];
            let nStartEndOriginalCounter = 0;
            let oStartMove = {};
            function checkOriginalMoveMark(oElement) {
                if (oMoveMarks[oElement.Name]) {
                    const oMoveMarkInfo = oMoveMarks[oElement.Name];
                    if (oMoveMarkInfo.From.Start === oElement || oMoveMarkInfo.To.Start === oElement) {
                        nStartEndOriginalCounter += 1;
                    } else if (oMoveMarkInfo.From.End === oElement || oMoveMarkInfo.To.End === oElement) {
                        nStartEndOriginalCounter -= 1;
                    }
                    for (let sName in oStartMove) {
                        oInsertMoveMarkId[sName] = false;
                        oRemoveMoveTypeNames[sName] = true;
                    }
                    oStartMove = {};
                }
            }

            function checkRevisedMoveMark(oElement) {
                if (oRevisedMoveMarksInserts[oElement.Id]) {
                    const arrMoveMarks = oRevisedMoveMarksInserts[oElement.Id];
                    for (let t = 0; t < arrMoveMarks.length; t += 1) {
                        const oRevisedInsertInfo = arrMoveMarks[t];
                        if (oRevisedInsertInfo.moveMark.Start) {
                            oStartMove[oRevisedInsertInfo.moveMark.Name] = true;
                        } else {
                            if (!nStartEndOriginalCounter) {
                                delete oStartMove[oRevisedInsertInfo.moveMark.Name];
                            }
                        }
                    }
                }
            }

            for (let j = 0; j < oRootElement.Content.length; j += 1) {
                const oChildElement = oRootElement.Content[j];
                if (oChildElement instanceof Paragraph) {
                    const arrContent = oChildElement.Content;
                    for (let k = 0; k < arrContent.length - 1; k += 1) {
                        const oElement = arrContent[k];
                        if (oElement instanceof AscCommon.CParaRevisionMove) {
                            checkOriginalMoveMark(oElement);
                        } else {
                            checkRevisedMoveMark(oElement);
                        }
                    }
                    const oParaEnd = oChildElement.GetParaEndRun();
                    checkRevisedMoveMark(oParaEnd);
                    const oLastMoveMark = oParaEnd.GetLastTrackMoveMark();
                    if (oLastMoveMark) {
                        checkOriginalMoveMark(oLastMoveMark);
                    }
                }
            }
        }
        for (let sName in oRemoveMoveTypeNames) {
            const oRevertMoveTypeByName = oComparisonMoveMarkManager.oRevertMoveTypeByName;
            if (oRevertMoveTypeByName[sName]) {
                for (let i = 0; i < oRevertMoveTypeByName[sName].length; i += 1) {
                    oRevertMoveTypeByName[sName][i].RemoveReviewMoveType();
                }
            }
        }
        for (let sId in oRevisedMoveMarksInserts) {
            const arrInsert = oRevisedMoveMarksInserts[sId];
            for (let i = arrInsert.length - 1; i >= 0; i -= 1) {
                const oInsertInfo = arrInsert[i];
                const oInsertParaMove = oInsertInfo.moveMark;
                if (oInsertMoveMarkId[oInsertParaMove.Name]) {
                    const oRun = AscCommon.g_oTableId.Get_ById(sId);
                    if (oInsertInfo.isParaEnd) {
                        oRun.AddAfterParaEnd(oInsertParaMove);
                        } else {
                        const oParagraph = oRun.Paragraph;
                        const nPosition = oRun.GetPosInParent(oParagraph);
                        oParagraph.AddToContent(nPosition, oInsertParaMove);
                    }
                }
            }
        }
    };

    CDocumentMergeComparison.prototype.setRemoveReviewType = function (element) {
        if (!(element.IsParaEndRun && element.IsParaEndRun())) {
            if (!element.GetReviewType || element.GetReviewType && element.GetReviewType() === reviewtype_Common) {
                this.setReviewInfoRecursive(element, this.nRemoveChangesType);
            }
        }
    }
    CDocumentMergeComparison.prototype.resolveCustomReviewTypesBetweenElements = CDocumentResolveConflictComparison.prototype.resolveCustomReviewTypesBetweenElements;

    CDocumentMergeComparison.prototype.checkParaEndReview = function (oNode) {
        if (oNode && oNode.element.GetType && oNode.element.GetType() === type_Paragraph && oNode.partner) {
            const oMainParaEnd = oNode.element.GetParaEndRun();
            const oRevisedParaEnd = oNode.partner.element.GetParaEndRun();
            const nRevisedReviewType = oRevisedParaEnd.GetReviewType();
            const oRevisedReviewInfo = oRevisedParaEnd.GetReviewInfo();
            const nOldMainMoveReviewType = oMainParaEnd.GetReviewMoveType();
            this.resolveCustomReviewTypesBetweenElements(oMainParaEnd, nRevisedReviewType, oRevisedReviewInfo);
            if (nOldMainMoveReviewType === Asc.c_oAscRevisionsMove.NoMove && oMainParaEnd.GetReviewMoveType() !== Asc.c_oAscRevisionsMove.NoMove) {
                const oRunMoveMark = oRevisedParaEnd.GetLastTrackMoveMark();
                if (oRunMoveMark) {
                    const oCopyMoveMark = oRunMoveMark.Copy(this.copyPr);
                    const sChangedMoveMarkName = this.oComparisonMoveMarkManager.getChangedMoveMarkName(oCopyMoveMark);
                    oCopyMoveMark.Name = sChangedMoveMarkName;
                    this.oComparisonMoveMarkManager.addRunMoveMarkNameRelation(sChangedMoveMarkName, oMainParaEnd);
                    this.oComparisonMoveMarkManager.addMoveMarkNameRunRelation(sChangedMoveMarkName, oMainParaEnd);
                    this.oComparisonMoveMarkManager.addRevisedMoveMarkToInserts(oCopyMoveMark, oMainParaEnd, oNode.element, true);

                }
            }
        }
    };

    CDocumentMergeComparison.prototype.applyChangesToTableSize = function(oNode) {
        this.copyPr.SkipUpdateInfo = false;
        this.copyPr.bSaveCustomReviewType = true;
        CDocumentComparison.prototype.applyChangesToTableSize.call(this, oNode);
        delete this.copyPr.bSaveCustomReviewType;
        this.copyPr.SkipUpdateInfo = true;
    }

    CDocumentMergeComparison.prototype.checkRowReview = function(oRowNode) {
        const oPartnerNode = oRowNode.partner;
        if (oPartnerNode) {
            const oMainRow = oRowNode.element;
            const oPartnerRow = oPartnerNode.element;
            const nRevisedReviewType = oPartnerRow.GetReviewType();
            const oRevisedReviewInfo = oPartnerRow.GetReviewInfo();
            this.resolveCustomReviewTypesBetweenElements(oMainRow, nRevisedReviewType, oRevisedReviewInfo);
        }
    };

    CDocumentMergeComparison.prototype.resolveConflicts = function (arrToInserts, arrToRemove, applyParagraph, nInsertPosition) {
        if (arrToInserts.length === 0 || arrToRemove.length === 0) return;
        arrToRemove.push(new AscCommonWord.ParaRun());
        arrToInserts.push(new AscCommonWord.ParaRun());
        arrToRemove[arrToRemove.length - 1].Content.push(new AscWord.CRunParagraphMark());
        arrToInserts[arrToInserts.length - 1].Content.push(new AscWord.CRunParagraphMark());
        const comparison = new CDocumentResolveConflictComparison(this.originalDocument, this.revisedDocument, this.options);
	    comparison.oBookmarkManager = this.oBookmarkManager;
        comparison.oComparisonMoveMarkManager = this.oComparisonMoveMarkManager;
        const originalDocument = new CMockDocument();
        const revisedDocument = new CMockDocument();
        const originalParagraph = new CMockParagraph();
        const revisedParagraph = new CMockParagraph();
        const origParagraph = applyParagraph;
        comparison.startPosition = nInsertPosition;
        comparison.parentParagraph = origParagraph;
        originalParagraph.Content = arrToRemove;
        revisedParagraph.Content = arrToInserts;
        originalDocument.Content.push(originalParagraph);
        revisedDocument.Content.push(revisedParagraph);

        comparison.oComparisonMoveMarkManager.executeResolveConflictMode(function () {
            comparison.compareRoots(originalDocument, revisedDocument);
        });
        return originalParagraph.Content;
    }

    CDocumentMergeComparison.prototype.getCompareReviewInfo = CDocumentResolveConflictComparison.prototype.getCompareReviewInfo;

    CDocumentMergeComparison.prototype.applyParagraphComparison = function (oOrigRoot, oRevisedRoot) {
        this.copyPr.SkipUpdateInfo = false;
        this.copyPr.bSaveCustomReviewType = true;
        CDocumentComparison.prototype.applyParagraphComparison.call(this, oOrigRoot, oRevisedRoot);
        for (let i = oOrigRoot.children.length - 1; i >= 0; i -= 1) {
            this.checkParaEndReview(oOrigRoot.children[i]);
        }
        const oParentContent = oOrigRoot.element.Content;
        const oLastElement = oParentContent[oParentContent.length - 1];
        if (oLastElement && oLastElement.GetReviewType() !== reviewtype_Common) {
            oLastElement.SetReviewTypeWithInfo(reviewtype_Common, new CReviewInfo());
        }

        delete this.copyPr.bSaveCustomReviewType;
        this.copyPr.SkipUpdateInfo = true;
    };

    CDocumentMergeComparison.prototype.getNodeConstructor = function () {
        return CMergeComparisonNode;
    };


    CDocumentMergeComparison.prototype.getTextElementConstructor = function () {
        return CMergeComparisonTextElement;
    };

    CDocumentMergeComparison.prototype.GetReviewTypeFromParaDrawing = function (oParaDrawing) {
        const oRun = oParaDrawing.GetRun();
        if (oRun) {
            return oRun.GetReviewType();
        }
        return reviewtype_Common;
    }

    CDocumentMergeComparison.prototype.compareDrawingObjects = function (oBaseDrawing, oCompareDrawing, bOrig) {
        if (oBaseDrawing && oCompareDrawing) {
            const baseReviewType = this.GetReviewTypeFromParaDrawing(oBaseDrawing);
            const compareReviewType = this.GetReviewTypeFromParaDrawing(oCompareDrawing);
            const arrOfReviewTypes = [];

            if (baseReviewType) arrOfReviewTypes.push(baseReviewType);
            if (compareReviewType) arrOfReviewTypes.push(compareReviewType);

            const priorityReviewType = getPriorityReviewType(arrOfReviewTypes);

            const oBaseRun = bOrig ? oBaseDrawing.GetRun() : oCompareDrawing.GetRun();
            this.setReviewInfoForArray([oBaseRun], priorityReviewType);
        }
        CDocumentComparison.prototype.compareDrawingObjects.call(this, oBaseDrawing, oCompareDrawing);
    }



    CDocumentMergeComparison.prototype.compare = function (callback) {
        const oOriginalDocument = this.originalDocument;
        const oRevisedDocument = this.revisedDocument;
        if (!oOriginalDocument || !oRevisedDocument) {
            return;
        }
	    this.oBookmarkManager.init(oOriginalDocument, oRevisedDocument);
        const oThis = this;
        const aImages = AscCommon.pptx_content_loader.End_UseFullUrl();
        const oObjectsForDownload = AscCommon.GetObjectsForImageDownload(aImages);
        const oApi = oOriginalDocument.GetApi();
        if (!oApi) {
            return;
        }
        const fCallback = function (data) {
            const oImageMap = {};
            AscFormat.ExecuteNoHistory(function () {
                AscCommon.ResetNewUrls(data, oObjectsForDownload.aUrls, oObjectsForDownload.aBuilderImagesByUrl, oImageMap);
            }, oThis, []);

            const NewNumbering = oRevisedDocument.Numbering.CopyAllNums(oOriginalDocument.Numbering);
            oRevisedDocument.CopyNumberingMap = NewNumbering.NumMap;
            oOriginalDocument.Numbering.AppendAbstractNums(NewNumbering.AbstractNum);
            oOriginalDocument.Numbering.AppendNums(NewNumbering.Num);
            for (let key in NewNumbering.NumMap) {
                if (NewNumbering.NumMap.hasOwnProperty(key)) {
                    oThis.checkedNums[NewNumbering.NumMap[key]] = true;
                }
            }
            oThis.compareRoots(oOriginalDocument, oRevisedDocument);
            oThis.compareSectPr(oOriginalDocument, oRevisedDocument);

            const oFonts = oOriginalDocument.Document_Get_AllFontNames();
            const aFonts = [];
            for (let i in oFonts) {
                if (oFonts.hasOwnProperty(i)) {
                    aFonts[aFonts.length] = new AscFonts.CFont(i, 0, "", 0, null);
                }
            }
            oApi.pre_Paste(aFonts, oImageMap, function () {
                callback && callback();
            });
        };
        AscCommon.sendImgUrls(oApi, oObjectsForDownload.aUrls, fCallback, true);
        return null;
    };

    function CDocumentMerge(oOriginalDocument, oRevisedDocument, oOptions) {
        this.originalDocument = oOriginalDocument;
        this.revisedDocument = oRevisedDocument;
        this.options = oOptions;
        this.api = oOriginalDocument.GetApi();
        this.comparison = new CDocumentMergeComparison(oOriginalDocument, oRevisedDocument, oOptions ? oOptions : new AscCommonWord.ComparisonOptions());
        this.oldTrackRevisions = false;
    }

    CDocumentMerge.prototype.resolveConflicts = CDocumentMergeComparison.prototype.resolveConflicts;

    CDocumentMerge.prototype.applyLastMergeCallback = function () {
        const oOriginalDocument = this.originalDocument;
        const oApi = this.api;
        if (!(oApi && oOriginalDocument)) {
            return;
        }
        this.comparison.correctMoveMarks();
        oOriginalDocument.SetTrackRevisions(this.oldTrackRevisions);
        const oTrackRevisionManager = oOriginalDocument.TrackRevisionsManager;
        oTrackRevisionManager.SkipPreDeleteMoveMarks = this.oldSkipPreDeleteMoveMarks;
        oOriginalDocument.End_SilentMode(false);
        oOriginalDocument.Recalculate();
        oOriginalDocument.UpdateInterface();
        oOriginalDocument.FinalizeAction();
        oApi.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.SlowOperation);
    }

    CDocumentMerge.prototype.merge = function () {
        const oOriginalDocument = this.originalDocument;
        const oRevisedDocument = this.revisedDocument;
        if (!oOriginalDocument || !oRevisedDocument) {
            return;
        }
        oOriginalDocument.StopRecalculate();
        oOriginalDocument.StartAction(AscDFH.historydescription_Document_MergeDocuments);
        oOriginalDocument.Start_SilentMode();
        this.oldTrackRevisions = oOriginalDocument.IsTrackRevisions();
        oOriginalDocument.SetTrackRevisions(false);
        const oTrackRevisionManager = oOriginalDocument.TrackRevisionsManager;
        this.oldSkipPreDeleteMoveMarks = oTrackRevisionManager.SkipPreDeleteMoveMarks;
        oTrackRevisionManager.SkipPreDeleteMoveMarks = true;
        this.comparison.compare(this.applyLastMergeCallback.bind(this));
    };


    function mergeBinary(oApi, sBinary2, oOptions) {
        const oDoc1 = oApi.WordControl.m_oLogicDocument;
        if (!window['NATIVE_EDITOR_ENJINE']) {
            const oCollaborativeEditing = oDoc1.CollaborativeEditing;
            if (oCollaborativeEditing && !oCollaborativeEditing.Is_SingleUser()) {
                oApi.sendEvent("asc_onError", Asc.c_oAscError.ID.CannotCompareInCoEditing, c_oAscError.Level.NoCritical);
                return;
            }
        }
        oApi.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.SlowOperation);

        const oDoc2 = AscFormat.ExecuteNoHistory(function () {
            const openParams = {noSendComments: true};
            let oDoc2 = new CDocument(oApi.WordControl.m_oDrawingDocument, true);
            oApi.WordControl.m_oDrawingDocument.m_oLogicDocument = oDoc2;
            oApi.WordControl.m_oLogicDocument = oDoc2;
            const oBinaryFileReader = new AscCommonWord.BinaryFileReader(oDoc2, openParams);
            AscCommon.pptx_content_loader.Start_UseFullUrl(oApi.insertDocumentUrlsData);
            if (!oBinaryFileReader.Read(sBinary2)) {
                oDoc2 = null;
            }
            oApi.WordControl.m_oDrawingDocument.m_oLogicDocument = oDoc1;
            oApi.WordControl.m_oLogicDocument = oDoc1;
            if (oDoc1.History)
                oDoc1.History.Set_LogicDocument(oDoc1);
            if (oDoc1.CollaborativeEditing)
                oDoc1.CollaborativeEditing.m_oLogicDocument = oDoc1;
            return oDoc2;
        }, this, []);

        oDoc1.History.Document = oDoc1;

        if (oDoc2) {
            const oMerge = new AscCommonWord.CDocumentMerge(oDoc1, oDoc2, oOptions ? oOptions : new AscCommonWord.ComparisonOptions());
            oMerge.merge();
        } else {
            AscCommon.pptx_content_loader.End_UseFullUrl();
        }

    }
    
    function mergeDocuments(oApi, oTmpDocument) {
        oApi.insertDocumentUrlsData = {
            imageMap: oTmpDocument["GetImageMap"](), documents: [], convertCallback: function (_api, url) {
            }, endCallback: function (_api) {
            }
        };
        mergeBinary(oApi, oTmpDocument["GetBinary"](), null, true);
        oApi.insertDocumentUrlsData = null;
    }

    window['AscCommonWord'].CDocumentMerge = CDocumentMerge;
    window['AscCommonWord'].mergeBinary = mergeBinary;
    window['AscCommonWord'].CMockMinHash = CMockMinHash;
    window['AscCommonWord'].CMockDocument = CMockDocument;
    window['AscCommonWord'].CMockParagraph = CMockParagraph;
    window['AscCommonWord']["mergeDocuments"] = window['AscCommonWord'].mergeDocuments = mergeDocuments;

})()
