import tar = require('tar'); // version ^6.0.1
import fs = require('fs');
import * as fse from 'fs-extra';
import zlib = require('zlib');
import * as path from 'path';
import * as common  from './common';


export async function baseToTar(pathBase: string, isDeleteOrg: boolean = true){

  const isExitBaseDir = (await common.dirExists(pathBase));
  if (!isExitBaseDir){
    return;
  }

  try {
    const f = await tar.c(
      {
        gzip: false, // this will perform the compression too
        file: path.resolve(pathBase, '..', '.base.tar'),
        cwd: path.dirname(pathBase)
        
      },
      [path.basename(pathBase)]
    );    
  } catch (error) {
    console.log(error);
    throw(error); 
  }
  isDeleteOrg && await common.deleteModulesInSrc(pathBase);
}

export async function tarToBase(pathBase: string){

  const tarFile = pathBase + '.tar';
  const isExistTar = (await common.fileExists(tarFile));
  if (!isExistTar){
    return;
  }

  try {
    await tar.x(  // or tar.extract(
      {
        file: tarFile,
        cwd:path.dirname(pathBase)
      }
    );    
  } catch (error) {
    console.log(error);
    throw(error);
  }

  fse.rmSync(tarFile);
}

