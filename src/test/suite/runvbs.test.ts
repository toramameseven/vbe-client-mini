import * as assert from 'assert';
import iconv = require('iconv-lite');
import path = require('path');
import * as fse from 'fs-extra';
import * as vbs from '../../vbsModule';
import * as vbecmCommon from '../../vbecmCommon';
import { expect } from 'chai';

const xlsPath = path.resolve(__dirname, '../../../xlsms');
const pathSrc = path.resolve(xlsPath, 'src_macrotest.xlsm');
const pathSrcBase = path.resolve(xlsPath, 'src_macrotest.xlsm', vbs.FOLDER_BASE);
const pathSrcVbe = path.resolve(xlsPath, 'src_macrotest.xlsm', vbs.FOLDER_VBE);
// windows display size,100% or others(125%)
let displaySize = 120;
const expectDir = displaySize === 100 ? 'src_macrotest_test.xlsm' : 'src_macrotest_test_note.xlsm';
const pathSrcExpect = path.resolve(xlsPath, expectDir);
const folderExpectFrx = path.resolve(xlsPath, 'src_macrotest_expect_frx.xlsm');
const bookPath = path.resolve(xlsPath, 'macrotest.xlsm');
const bookPathOriginal = path.resolve(xlsPath, 'macrotest_org.xlsm');
const pathTestModuleBas = path.resolve(pathSrc, 'Module1.bas');
const pathTestModuleCls = path.resolve(pathSrc, 'Class1.cls');
const pathTestModuleFrm = path.resolve(pathSrc, 'UserForm1.frm');
const pathTestModuleShtCls = path.resolve(pathSrc, 'Sheet1.sht.cls');

//const exportTest = true;
const importTest = true;
const exportFrx = true;
const pushTest = true;
const pullTest = true;
const runTest = true;
const commonTest = true;

// export test
describe('#vbsModules.exportModulesAndSyncronize', () => {
  before('#Before exportModulesAndSyncronize', async function () {
    this.timeout(60000);

    vbs.closeBook(bookPath);
    (await vbecmCommon.fileExists(bookPath)) && fse.rmSync(bookPath);
    fse.copyFileSync(bookPathOriginal, bookPath);

    await vbecmCommon.rmDirIfExist(pathSrc, { recursive: true, force: true });

    if (await vbecmCommon.dirExists(pathSrc)) {
      await vbs.deleteModulesInFolder(pathSrc);
      await vbs.deleteModulesInFolder(pathSrcBase);
      await vbs.deleteModulesInFolder(pathSrcVbe);
    }
  });

  // test code
  beforeEach('#Before Each exportModulesAndSyncronize', async function () {
    this.timeout(60000);
    try {
      await vbs.exportModulesAndSynchronize(bookPath);
    } catch (error) {
      console.log(error);
    }
  });

  // check
  describe('##vbs Test main', () => {
    ['create folder and export to the folder', 'export to created folder'].forEach(
      async (message) => {
        it(message + ' src', async () => {
          const r = vbs.comparePath(pathSrc, pathSrcExpect);
          assert.strictEqual(r.same, true, message + ' src');
        });
        it(message + ' base', async () => {
          const r = vbs.comparePath(pathSrcBase, pathSrcExpect);
          assert.strictEqual(r.same, true, message + ' base');
        });
        it(message + ' vbe', async () => {
          const r = vbs.comparePath(pathSrcVbe, pathSrcExpect);
          assert.strictEqual(r.same, true, message + ' vbe');
        });
      }
    );
  });
});

// import modules test
describe('#vbsModules.importModules', () => {
  if (!importTest) {
    return;
  }

  before(async function () {
    this.timeout(60000);
    await vbs.exportModulesAndSynchronize(bookPath);
    await vbs.importModules(bookPath, vbs.STRING_EMPTY);
  });

  // check
  describe('##importModules main', async () => {
    const message = 'import';
    it(message + ' src', async () => {
      const r = vbs.comparePath(pathSrc, pathSrcExpect);
      assert.strictEqual(r.same, true, message + ' src');
    });
    it(message + ' base', async () => {
      const r = vbs.comparePath(pathSrcBase, pathSrcExpect);
      assert.strictEqual(r.same, true, message + ' base');
    });
    it(message + ' vbe', async () => {
      const r = vbs.comparePath(pathSrcVbe, pathSrcExpect);
      assert.strictEqual(r.same, true, message + ' vbe');
    });
  });
});

// export frx modules test
describe('#vbsModules.exportFrxModules', () => {
  if (!exportFrx) {
    return;
  }

  before(async function () {
    await vbs.deleteModulesInFolder(pathSrc);
    await vbs.deleteModulesInFolder(pathSrcBase);
    await vbs.deleteModulesInFolder(pathSrcVbe);

    try {
      await vbs.exportFrxModules(bookPath);
    } catch (error) {
      //
    }
  });

  // check  // todo test if file exists or not
  describe('##exportFrxModules main', async function () {
    this.timeout(60000);
    const message = 'frx export';
    it(message + ' src', async () => {
      const r = vbs.comparePath(pathSrc, folderExpectFrx);
      assert.strictEqual(r.same, true, message + ' src');
    });
    it(message + ' base', async () => {
      const r = vbs.comparePath(pathSrcBase, folderExpectFrx);
      assert.strictEqual(r.same, true, message + ' base');
    });
    it(message + ' vbe', async () => {
      const r = vbs.comparePath(pathSrcVbe, folderExpectFrx);
      assert.strictEqual(r.same, true, message + ' vbe');
    });
  });
});

// push a module test
describe('#vbsModules.commitModule', function () {
  before('#Before push module', async function () {
    this.timeout(60000);
    try {
      await vbs.exportModulesAndSynchronize(bookPath);
    } catch (error) {
      console.log(error);
    }
  });

  if (!pushTest) {
    return;
  }
  // check
  it('## push modules bas', async function () {
    this.timeout(60000);
    let r = await vbs.importModules(bookPath, pathTestModuleBas);
    assert.strictEqual(r, true);
  });
  it('## push modules cls', async function () {
    this.timeout(60000);
    let r = await vbs.importModules(bookPath, pathTestModuleCls);
    assert.strictEqual(r, true);
  });
  it('## push modules frm', async function () {
    this.timeout(60000);
    let r = await vbs.importModules(bookPath, pathTestModuleFrm);
    assert.strictEqual(r, true);
  });
  it('## push modules sht.cls', async function () {
    this.timeout(60000);
    let r = await vbs.importModules(bookPath, pathTestModuleShtCls);
    assert.strictEqual(r, true);
  });
});

// pull a module test
describe('#vbsModules.pullModule', function () {
  if (!pullTest) {
    return;
  }
  // check
  it('## update module bas', async function () {
    this.timeout(60000);

    let r = await vbs.exportModulesAndSynchronize(bookPath, pathTestModuleBas);
    assert.strictEqual(r, true);
  });
  it('## update module cls', async function () {
    this.timeout(60000);

    let r = await vbs.exportModulesAndSynchronize(bookPath, pathTestModuleCls);
    assert.strictEqual(r, true);
  });
  it('## update module frm', async function () {
    this.timeout(60000);

    let r = await vbs.exportModulesAndSynchronize(bookPath, pathTestModuleFrm);
    assert.strictEqual(r, true);
  });
  it('## update module sht.cls', async function () {
    this.timeout(60000);

    let r = await vbs.exportModulesAndSynchronize(bookPath, pathTestModuleShtCls);
    assert.strictEqual(r, true);
  });
});

// run sub function test
describe('#vbsModules.vbaRun', function () {
  if (!runTest) {
    return;
  }
  // check
  it('## rum module', async function () {
    this.timeout(60000);
    let returnValue = await vbs.vbaSubRun(bookPath, 'module1', 'test4');
    assert.strictEqual(returnValue, '10');
  });
});

// run vbecmCommon.ts test
describe('#vbecmCommon.ts', async () => {
  if (!commonTest) {
    return;
  }
  //return;

  const folderExists = __dirname;
  const folderNotExist = path.resolve(__dirname, 'xxxxxxxxxxxxxxxxx');

  const fileExists = __filename;
  const fileNotExits = path.resolve(__filename, 'xxxxxxxxxxxxxxxxx');

  // dir
  it('test exists folder', async () => {
    assert.strictEqual(await vbecmCommon.dirExists(folderExists), true);
  });

  it('test not exist folder', async () => {
    assert.strictEqual(await vbecmCommon.dirExists(folderNotExist), false);
  });

  it('test not exist Folder, but file exists', async () => {
    assert.strictEqual(await vbecmCommon.dirExists(fileExists), false);
  });

  // folder
  it('test exists file', async () => {
    assert.strictEqual(await vbecmCommon.fileExists(fileExists), true);
  });

  it('test not exist file', async () => {
    assert.strictEqual(await vbecmCommon.fileExists(fileNotExits), false);
  });

  it('test not exist file, but exists folder', async () => {
    assert.strictEqual(await vbecmCommon.fileExists(folderExists), false);
  });

  it('test not exist file, but exists folder', async () => {
    const r = await vbecmCommon.fileExists(folderExists);
    expect(r).is.false;
  });
});
