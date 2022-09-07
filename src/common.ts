import * as fse from 'fs-extra';
import * as path from 'path';


export async function fileExists(filepath: string) {
  try {
    const res = (await fse.promises.lstat(filepath)).isFile();
    return (res);
  } catch (e) {
    return false;
  }
}


export async function dirExists(filepath: string) {
  try {
    const res = (await fse.promises.lstat(filepath)).isDirectory();
    return (res);
  } catch (e) {
    return false;
  }
}


export function myLog(message: string, title: string = '(_empty_)')
{
  //console.log(`vbecm: ${title}: ${message}`);
}


export async function rmDirIfExist(pathFolder: string, option: {}){
  try {

    const isExist = await dirExists(pathFolder);
    if (!isExist)
    {
      // no folder no delete
      return;
    }
    await fse.promises.rm(pathFolder, option);
  } catch (error) {
    throw(error);
  }
}

export async function rmFileIfExist(pathFile: string, option: {}){
  try {

    const isExist = await fileExists(pathFile);
    if (!isExist)
    {
      // no file no delete
      return;
    }
    await fse.promises.rm(pathFile, option);
  } catch (error) {
    throw(error);
  }
}



