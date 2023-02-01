const _ = require('lodash')
const fs = require('fs-extra')
const { gengerate } = require('./wirite')
const { readFile } = require('./read')
function start(fromFile, toFileName) {
  const { peList, petList,otherList, commonList } = readFile(fromFile)
  let Mother = fromFile.substring(4,6)
  Mother = parseInt(Mother)

  const FileName = toFileName

  function iowritePPP(list) {
    _.forEach(list, (item, index) => {
      if(item && item.length > 0) {
        const depName = item[0].dep ? item[0].dep : '(其他)'
        gengerate(item, Mother, FileName, `${depName}-${index + 1}`)
      }
    })
  }

  function iowriteCommon(list) {
    _.forEach(list, (item) => {
      if(item && item.length > 0) {
        gengerate(item, Mother, FileName, `${item[0].name}`)
      }
    })
  }

  fs.emptyDir('./dist', (err) => {
    if (err) {
      console.error(err)
    }
    iowritePPP(peList)
    iowritePPP(petList)
    iowritePPP(otherList)
    iowriteCommon(commonList)
  })
}
// （xxxx01xxx.xlsx,'xxxxxxx产线计时明细表'）
// start('xxxx01xxx.xlsx', 'xxxxxxx产线计时明细表')