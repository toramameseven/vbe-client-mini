import * as fse from 'fs-extra';
import * as path from 'path';


/**
 * test file exits
 * @param filepath
 * @returns 
 */
 export async function fileExists(filepath: string) {
  try {
    const res = (await fse.promises.lstat(filepath)).isFile();
    return (res);
  } catch (e) {
    return false;
  }
}

/**
 * test dir exists
 * @param filepath
 * @returns 
 */
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


export async function deleteModulesInSrc(pathSrc: string){
  try {

    const isExist = await dirExists(pathSrc);
    if (!isExist)
    {
      // フォルダが存在しないので 消さない
      return;
    }

    //ファイルとディレクトリのリストが格納される(配列)
    const files = fse.readdirSync(pathSrc);
    
    //ディレクトリのリストに絞る
    const moduleList = files.forEach((file) => {
        const fullPath = path.resolve(pathSrc, file);
        const isFile = fse.statSync(fullPath).isFile();
        const ext  = path.extname(file).toLowerCase();
        if (isFile && ['.bas','.frm','.cls', '.frx'].includes(ext)){
          fse.rmSync(fullPath);
        }
    });   
  } catch (error) {
    throw(error);
  }
}



