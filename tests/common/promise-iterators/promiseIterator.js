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

$(function ()
{

	QUnit.module("Test sequential execution for promises");


	QUnit.test("Test promise iterator forEach", function (assert)
	{
		assert.timeout(1000)
		const done = assert.async(5);
		const p1 = new Promise(function (resolve)
		{
			setTimeout(function ()
			{
				resolve(1);
			}, 500);
		});

		let p2;

		function createRejectPromise()
		{
			p2 = new Promise(function (resolve, reject)
			{
				setTimeout(function ()
				{
					reject(2);
				}, 300);
			});
			return p2;
		}

		assert.rejects(createRejectPromise())

		const p3 = new Promise(function (resolve)
		{
			setTimeout(function ()
			{
				resolve(3);
			}, 100);
		});

		const p4 = new Promise(function (resolve)
		{
			setTimeout(function ()
			{
				resolve(4);
			}, 0);
		});


		const oPromiseIterator = new AscCommon.CPromiseIterator([p1, p2, p3, p4]);
		let nAnswerIndex = 0;
		const arrAnswers = [1, 2, 3, 4];
		oPromiseIterator.forEachValue(function (oValue)
		{
			assert.strictEqual(oValue, arrAnswers[nAnswerIndex]);
			done();
			nAnswerIndex += 1;
		}, function (oValue)
		{
			assert.strictEqual(oValue, arrAnswers[nAnswerIndex]);
			done();
			nAnswerIndex += 1;
		}, function ()
		{
			assert.strictEqual(nAnswerIndex, 4);
			done();
		});
	});

	QUnit.test("Test promise iterator forAll", function (assert)
	{
		assert.timeout(1000)
		const done = assert.async();
		const p1 = new Promise(function (resolve)
		{
			setTimeout(function ()
			{
				resolve(1);
			}, 500);
		});

		let p2;

		function createRejectPromise()
		{
			p2 = new Promise(function (resolve, reject)
			{
				setTimeout(function ()
				{
					reject(2);
				}, 300);
			});
			return p2;
		}

		assert.rejects(createRejectPromise())

		const p3 = new Promise(function (resolve)
		{
			setTimeout(function ()
			{
				resolve(3);
			}, 100);
		});

		const p4 = new Promise(function (resolve)
		{
			setTimeout(function ()
			{
				resolve(4);
			}, 0);
		});


		const oPromiseIterator = new AscCommon.CPromiseIterator([p1, p2, p3, p4]);

		oPromiseIterator.forAllSuccessValues(function (arrValues)
		{
			assert.deepEqual(arrValues, [1, 3, 4]);
			done();
		});
	});

	QUnit.test("Test promise getter iterator forEach", function (assert)
	{
		assert.timeout(1500)
		const done = assert.async(5);
		const p1 = function ()
		{
			return new Promise(function (resolve)
			{
				setTimeout(function ()
				{
					resolve(1);
				}, 500);
			});
		}


		const p2 = function ()
		{
			return new Promise(function (resolve, reject)
			{
				setTimeout(function ()
				{
					reject(2);
				}, 300);
			});
		};

		const p3 = function ()
		{
			return new Promise(function (resolve)
			{
				setTimeout(function ()
				{
					resolve(3);
				}, 100);
			});
		};

		const p4 = function ()
		{
			return new Promise(function (resolve)
			{
				setTimeout(function ()
				{
					resolve(4);
				}, 0);
			});
		}

		const oPromiseIterator = new AscCommon.CPromiseGetterIterator([p1, p2, p3, p4]);
		let nAnswerIndex = 0;
		const arrAnswers = [1, 2, 3, 4];
		oPromiseIterator.forEachValue(function (oValue)
		{
			assert.strictEqual(oValue, arrAnswers[nAnswerIndex]);
			done();
			nAnswerIndex += 1;
		}, function (oValue)
		{
			assert.strictEqual(oValue, arrAnswers[nAnswerIndex]);
			done();
			nAnswerIndex += 1;
		}, function ()
		{
			assert.strictEqual(nAnswerIndex, 4);
			done();
		});
	});

	QUnit.test("Test promise getter iterator forAll", function (assert)
	{
		assert.timeout(1000)
		const done = assert.async();
		const p1 = function ()
		{
			return new Promise(function (resolve)
			{
				setTimeout(function ()
				{
					resolve(1);
				}, 500);
			});
		}

		const p2 = function ()
		{
			return new Promise(function (resolve, reject)
			{
				setTimeout(function ()
				{
					reject(2);
				}, 300);
			});
		}


		const p3 = function ()
		{
			return new Promise(function (resolve)
			{
				setTimeout(function ()
				{
					resolve(3);
				}, 100);
			});
		}

		const p4 = function ()
		{
			return new Promise(function (resolve)
			{
				setTimeout(function ()
				{
					resolve(4);
				}, 0);
			});
		}


		const oPromiseIterator = new AscCommon.CPromiseGetterIterator([p1, p2, p3, p4]);

		oPromiseIterator.forAllSuccessValues(function (arrValues)
		{
			assert.deepEqual(arrValues, [1, 3, 4]);
			done();
		});
	});
});
