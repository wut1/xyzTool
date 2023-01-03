const _ = require('lodash')
const fs = require('fs-extra')
const { gengerate } = require('./wirite')
const { readFile } = require('./read')
function start(fromFile, toFileName) {
  const { peList, petList, commonList } = readFile(fromFile)
  const Mother = fromFile.substring(4,6)

  const FileName = toFileName

  function iowritePPP(list) {
    _.forEach(list, (item, index) => {
      gengerate(item, Mother, FileName, `${item[0].dep}-${index + 1}`)
    })
  }

  function iowriteCommon(list) {
    _.forEach(list, (item) => {
      gengerate(item, Mother, FileName, `${item[0].name}`)
    })
  }

  fs.emptyDir('./dist', (err) => {
    if (err) {
      console.error(err)
    }
    iowritePPP(peList)
    iowritePPP(petList)
    iowriteCommon(commonList)
  })
}

// start(fileFrom, fileName)