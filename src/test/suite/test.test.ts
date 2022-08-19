import * as assert from 'assert';
import iconv = require("iconv-lite");
import path = require('path');
import * as vscode from 'vscode';
import * as fs from 'fs';

suite('Extension Test Suite', () => {
	vscode.window.showInformationMessage('Start all tests.');
  assert.strictEqual(-1, [1, 2, 3].indexOf(0));
	// test('Sample test', async () => {
  //   const ret = dirExists('C:\\projects\\toramame-hub\\vbe-client-mini');
	// 	assert.strictEqual(true, ret);
	// });

});


// async function dirExists(filepath: string) {
//   try {
//     const res = (await fs.promises.lstat(filepath)).isDirectory();
//     return (res);
//   } catch (e) {
//     return false;
//   }
// }
