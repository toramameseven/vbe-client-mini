import * as assert from 'assert';
import iconv = require("iconv-lite");
import path = require('path');
import * as fs from 'fs'; 
import { compare, compareSync, Options, Result, DifferenceState } from 'dir-compare';
import { disconnect } from 'process';
import * as vbs from '../../vbsModule';

// You can import and use all API from the 'vscode' module
// as well as import your extension to test it
import * as vscode from 'vscode';
// import * as myExtension from '../../extension';

suite('Extension Test Suite2', () => {
	vscode.window.showInformationMessage('Start all tests.');

  const options: Options = { 
    compareSize: true, 
    compareContent: true,
    excludeFilter: "**/*.frx"
  };


  const rootFolder = path.dirname(__filename);
  const xlsPath = path.resolve(rootFolder, "../../../xlsms");

  const path1 = path.resolve(xlsPath, 'src_macrotest.xlsm');
  const path2 = path.resolve(xlsPath, 'src_macrotest_test.xlsm');

  const uriXlsmFile = vscode.Uri.file(path.resolve(xlsPath, 'macrotest.xlsm'));

  fs.rmSync(path1, { recursive: true, force: true });

  vbs.exportModuleSync(uriXlsmFile);

	test('export with create folder', () => {
    const r = compareSync(path1, path2, options);
		assert.strictEqual( r.same, true, 'test export files.');
	});

  vbs.exportModuleSync(uriXlsmFile);

  test('export to exist folder', () => {
    const r = compareSync(path1, path2, options);
		assert.strictEqual( r.same, true, 'test export files.');
	});

  vbs.importModuleSync(uriXlsmFile);
  vbs.exportModuleSync(uriXlsmFile);
  test('import test', () => {
    const r = compareSync(path1, path2, options);
		assert.strictEqual( r.same, true, 'test export files.');
	});
});


// Synchronous
// export const doDiff = (path1: string, path2: string) =>{
//   const res : Result = compareSync(path1, path2, options);
//   return res;
// };

// function print(result : Result) {
//   console.log('Directories are %s', result.same ? 'identical' : 'different');

//   console.log('Statistics - equal entries: %s, distinct entries: %s, left only entries: %s, right only entries: %s, differences: %s',
//     result.equal, result.distinct, result.left, result.right, result.differences);

//   result.diffSet?.filter((dif) => dif.state === 'distinct' )
//   .forEach((dif) => 
//     console.log('Difference - name1: %s, type1: %s, name2: %s, type2: %s, state: %s',
//                 dif.name1, dif.type1, dif.name2, dif.type2, dif.state));
// };
