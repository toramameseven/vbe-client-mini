import * as fs from 'fs';
import * as path from 'path';


/**
 * test file exits
 * @param filepath
 * @returns 
 */
 export async function fileExists(filepath: string) {
  try {
    const res = (await fs.promises.lstat(filepath)).isFile();
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
    const res = (await fs.promises.lstat(filepath)).isDirectory();
    return (res);
  } catch (e) {
    return false;
  }
}


export function myLog(message: string, title: string = '(_empty_)')
{
  console.log(`vbecm: ${title}: ${message}`);
}



