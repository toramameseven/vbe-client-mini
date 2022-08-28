import * as assert from 'assert';
import iconv = require('iconv-lite');
import path = require('path');
import * as fse from 'fs-extra';
import * as vbs from '../../vbsModule';
import * as common  from '../../common';
import { expect } from 'chai';


function getTestPath(){
  const xlsPath = path.resolve(__dirname, '../../../xlsms');
  const pathSrc = path.resolve(xlsPath, 'src_macrotest.xlsm');
  // windows display size,100% or others(125%)
  let displaySize = 120;
  const testDir = displaySize === 100 ? 'src_macrotest_test.xlsm' : 'src_macrotest_test_note.xlsm';
  const pathSrcBase = path.resolve(xlsPath, 'src_macrotest.xlsm/.base');
  const pathSrcExpect = path.resolve(xlsPath, testDir);
  const bookPath = path.resolve(xlsPath, 'macrotest.xlsm');
  const bookPathOriginal = path.resolve(xlsPath, 'macrotest_org.xlsm');
  const pathTestModule = path.resolve(pathSrc, 'Module1.bas');

  return {
    pathSrc,
    pathSrcBase,
    pathSrcExpect,
    bookPath,
    bookPathOriginal,
    pathTestModule
  };
}


describe('#vbsModules.exportModulesToScrAndBase', () => {
  
  const {
    pathSrc,
    pathSrcBase,
    pathSrcExpect,
    bookPath,
    bookPathOriginal,
  } = getTestPath();


  before('#Before exportModulesToScrAndBase', async function(){
    this.timeout(60000);

    vbs.closeBook(bookPath);
    (await common.fileExists (bookPath)) && fse.rmSync(bookPath);
    fse.copyFileSync(bookPathOriginal, bookPath);

    await vbs.rmDirIfExist(pathSrc, { recursive: true, force: true });

    if (await common.dirExists(pathSrc)){
      await vbs.deleteModulesInSrc(pathSrc);
      await vbs.deleteModulesInSrc(pathSrcBase);
    }
  });
  
  // test code
  beforeEach('#Before Each exportModulesToScrAndBase', async function(){
    this.timeout(60000);
    try {
      await vbs.exportModulesToScrAndBase(bookPath);   
    } catch (error) {
      console.log(error);
    }
  });

  // check
  describe('##vbs Test main', () => {
    const {
      pathSrc,
      pathSrcExpect,
      bookPath,
      bookPathOriginal,
    } = getTestPath();

    ['create folder and export to the folder','export to created folder'].forEach(
      async (message) => {
        it(message, async () => {
        const  r = vbs.comparePath(pathSrc, pathSrcExpect);
        assert.strictEqual(r.same, true, message);
      });} 
    );
  });
});


describe('#vbsModules.importModules', () => {
  
  const {
    pathSrc,
    pathSrcExpect,
    bookPath,
    bookPathOriginal,
  } = getTestPath();

  before(async function(){
    this.timeout(60000);
    await vbs.exportModulesToScrAndBase(bookPath); 
    await vbs.importModules(bookPath);
  });

  // check
  describe('##importModules main', async () => {
    it('## import modules', async () => {
      const  r = vbs.comparePath(pathSrc, pathSrcExpect);
      assert.strictEqual(r.same, true, 'import modules');
    });

  });
});

//
describe('#vbsModules.exportFrxModules', () => {
  const {
    pathSrc,
    pathSrcExpect,
    bookPath,
    bookPathOriginal,
  } = getTestPath();

  // check
  describe('##exportFrxModules main', async function(){
    this.timeout(60000);
    it('## export frx modules', async () => {
      let r = false;      
      try {
        await vbs.exportFrxModules(bookPath); 
        r = true;
      } catch (error) {
        //
      }
      assert.strictEqual(r, true);
    });
  });
});

describe('#vbsModules.commitModule', function(){
  const {
    pathSrc,
    pathSrcExpect,
    bookPath,
    bookPathOriginal,
    pathTestModule
  } = getTestPath();

  // check
  it('## commit modules', async function(){
    this.timeout(60000);
    let r = await vbs.commitModule(bookPath, pathTestModule); 
    assert.strictEqual(r, true);
  });
});

describe('#vbsModules.updateModule', function(){
  const {
    pathSrc,
    pathSrcExpect,
    bookPath,
    bookPathOriginal,
    pathTestModule
  } = getTestPath();

  // check
  it('## update module', async function(){
    this.timeout(60000);

    let r = await vbs.updateModule(bookPath, pathTestModule); 
    assert.strictEqual(r, true);
  });
});

describe('#vbsModules.vbaRun', function(){
  const {
    bookPath,
  } = getTestPath();

  // check
  it('## rum module', async function(){
    this.timeout(60000);
    let returnValue = await vbs.vbaSubRun(bookPath, 'module1', 'test4'); 
    assert.strictEqual(returnValue, '10');
  });
});

describe('#common.ts', async () => {

  return;

  const folderExists = __dirname;
  const folderNotExist = path.resolve(__dirname, 'xxxxxxxxxxxxxxxxx');

  const fileExists = __filename;
  const fileNotExits = path.resolve(__filename, 'xxxxxxxxxxxxxxxxx');

  // dir
  it('test exists folder', async () => {
    assert.strictEqual(await common.dirExists(folderExists), true);
  });

  it('test not exist folder', async () => {
    assert.strictEqual(await common.dirExists(folderNotExist), false);
  });

  it('test not exist Folder, but file exists', async () => {
    assert.strictEqual(await common.dirExists(fileExists), false);
  });

  // folder
  it('test exists file', async () => {
    assert.strictEqual(await common.fileExists(fileExists), true);
  });

  it('test not exist file', async () => {
    assert.strictEqual(await common.fileExists(fileNotExits), false);
  });

  it('test not exist file, but exists folder', async () => {
    assert.strictEqual(await common.fileExists(folderExists), false);
  });

  it('test not exist file, but exists folder', async () => {
    const r = await common.fileExists(folderExists);
    expect(r).is.false;
  });
});


