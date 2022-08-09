// import { compare, compareSync, Options, Result, DifferenceState } from 'dir-compare';
// import { disconnect } from 'process';

// const options: Options = { 
//   compareSize: true, 
//   compareContent: true,
//   excludeFilter: "**/*.frx"
// };
// // Multiple compare strategy can be used simultaneously - compareSize, compareContent, compareDate, compareSymlink.
// // If one comparison fails for a pair of files, they are considered distinct.
// const path1 = 'C:\\projects\\toramame-hub\\xy\\xlsms\\src_macroTest.xlsm\\.base';
// const path2 = 'C:\\projects\\toramame-hub\\xy\\xlsms\\src_macroTest.xlsm\\.current';

// // Synchronous
// export const doDiff = (path1: string, path2: string) =>{
//   const res : Result = compareSync(path1, path2, options)
//   print(res)
// }

// function print(result : Result) {
//   console.log('Directories are %s', result.same ? 'identical' : 'different')

//   console.log('Statistics - equal entries: %s, distinct entries: %s, left only entries: %s, right only entries: %s, differences: %s',
//     result.equal, result.distinct, result.left, result.right, result.differences)

//   result.diffSet?.filter(dif => dif.state === 'distinct' ).forEach(dif => console.log('Difference - name1: %s, type1: %s, name2: %s, type2: %s, state: %s',
//     dif.name1, dif.type1, dif.name2, dif.type2, dif.state))
// }