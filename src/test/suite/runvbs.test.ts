import * as assert from 'assert';
import iconv = require('iconv-lite');
import path = require('path');
import * as fs from 'fs'; 
import { compare, compareSync, Options, Result, DifferenceState } from 'dir-compare';
import { disconnect } from 'process';
import * as vbs from '../../vbsModule';
import * as common  from '../../common';
import * as vscode from 'vscode';
import { expect } from 'chai';





function getTestPath(){
  const xlsPath = path.resolve(__dirname, '../../../xlsms');
  const pathSrc = path.resolve(xlsPath, 'src_macrotest.xlsm');
  const pathSrcBase = path.resolve(xlsPath, 'src_macrotest.xlsm/.base');
  const pathSrcExpect = path.resolve(xlsPath, 'src_macrotest_test.xlsm');
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

describe('#vbsModules.exportModules', () => {
  const {
    pathSrc,
    pathSrcBase,
    pathSrcExpect,
    bookPath,
    bookPathOriginal,
  } = getTestPath();


  before(async function(){
    this.timeout(60000);

    vbs.closeBook(bookPath);
    fs.rmSync(bookPath);
    fs.copyFileSync(bookPathOriginal, bookPath);

    console.log('                                           ======= before');
    if (await common.dirExists(pathSrc)){
      // fs.rmSync(pathSrc, { recursive: true, force: true });
      vbs.deleteModulesInSrc(pathSrc);
      vbs.deleteModulesInSrc(pathSrcBase);
    }
  });
  
  // test code
  beforeEach(async function(){
    this.timeout(60000);
    console.log('                                           ======= beforeEach');
    try {
      await vbs.exportModules(bookPath);   
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

    const options: Options = {
      compareSize: true,
      compareContent: true,
      skipSubdirs: true,
      excludeFilter: '**/*.frx,.base,.current'
    };

    ['create folder and export to the folder','export to created folder'].forEach(
      async (message) => {
        it(message, async () => {
        let r : Result;
        try {
          r = await compare(pathSrc, pathSrcExpect, options);
        } catch (error) {
          console.log(error);
          throw(error);
        }
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
    await vbs.exportModules(bookPath); 
    await vbs.importModules(bookPath);
  });

  // check
  describe('##vbs Test main', async () => {
  
    const options: Options = { 
      compareSize: true, 
      compareContent: true,
      excludeFilter: '**/*.frx,.base,.current'
    };
    
    it('## import modules', async () => {
      let r : Result;
      try {
        r = await compare(pathSrc, pathSrcExpect, options);
      } catch (error) {
        console.log(error);
        throw(error);
      }
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
  describe('##vbs Test main', async function(){
    this.timeout(60000);
    it('## export frx modules', async () => {
      let r : boolean = true;
      try {
        await vbs.exportFrxModules(bookPath); 
      } catch (error) {
        console.log(error);
        r = false;
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

    let r : boolean = true;
    try {
      await vbs.commitModule(bookPath, pathTestModule); 
    } catch (error) {
      console.log(error);
      r = false;
    }
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

    let r : boolean = true;
    try {
      await vbs.updateModule(bookPath, pathTestModule); 
    } catch (error) {
      console.log(error);
      r = false;
    }
    assert.strictEqual(r, true);
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
