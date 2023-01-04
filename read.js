const xlsx = require('xlsx')
const _ = require('lodash')
const chalk = require('chalk')

function getMapJson(json) {
  const obj = {}
  _.forEach(Object.entries(json), ([key, value]) => {
    obj[value] = key
  })
  return obj
}

function forIn(arr) {
  let end = 31

  if (end > arr.length - 1) {
    return [arr]
  } else {
    if (arr[end].dayNum === arr[end - 1].dayNum) {
      end = _.findIndex(arr, (o) => o.dayNum === arr[end].dayNum)
    }
    if(end > 31 || end === 0) {
      end = 31
    }
    const pre = _.slice(arr, 0, end)
    const next = _.slice(arr, end)
    return _.concat([pre], forIn(next))
  }
}

function dtoMap(cur) {
  const newCur = { ...cur }
  // 不统计 工时为0 的人
  // if (newCur.dayList && newCur.dayList.length === 0) {
  //   newCur.dayList = [{ dayNum: 0, time: 0 }]
  // }
  let arr = _.map(newCur.dayList, (item) => {
    return {
      ...item,
      name: newCur.name,
      dep: newCur.dep,
      jobNum: newCur.jobNum,
    }
  })
  return arr
}

function sortMergeList(list) {
  let mergeList = _.reduce(
    list,
    (total, cur) => {
      let arr = dtoMap(cur)
      return _.concat(total, arr)
    },
    [],
  )
  mergeList = _.sortBy(mergeList, ['dayNum', 'jobNum'])

  const newlist = forIn(mergeList)

  const nums = _.reduce(
    newlist,
    (total, curArr) => {
      return total + curArr.length
    },
    0,
  )

  if (mergeList.length !== nums) {
    console.log(chalk.red.bold('数据不对了!!!!!!!'))
  }

  return newlist
}

function readFile(dirFile) {
  const wb = xlsx.readFile(dirFile, { cellStyles: true })
  const sheet = wb.Sheets[wb.SheetNames[0]]

  const sheetJsonA = xlsx.utils.sheet_to_json(sheet, { header: 'A' })

  const mapJsonA = sheetJsonA[0]

  const mapJson = getMapJson(mapJsonA)

  // B：工号 C:部门 D: 姓名

  const B = mapJson['工号'] || 'B'
  const C = mapJson['部门'] || 'C'
  const D = mapJson['姓名'] || 'D'

  let sheetJson = _.filter(
    sheetJsonA,
    (sjObj) => !!sjObj[D] && sjObj[B] !== '工号',
  )

  const AllList = _.map(sheetJson, (sjObj, index) => {
    const depAll = sjObj[C]
    const dep = depAll && depAll.split(/([A-Z]+)/)[1]

    const newObj = { dayList: [], jobNum: sjObj[B], dep, name: sjObj[D] }

    _.forEach(Object.entries(sjObj), ([key, value]) => {
      const dayNum = mapJsonA[key]
      if (1 <= +dayNum && +dayNum <= 31) {
        const content = sheet[`${key}${index + 2}`]
        const listObj = { dayNum, time: value }
        if (
          content.s &&
          content.s.fgColor &&
          content.s.fgColor.rgb &&
          content.s.fgColor.rgb === 'FFFF00'
        ) {
          listObj.white = true
        }
        newObj.dayList.push(listObj)
      }
    })
    return newObj
  })

  const mergePEList = _.filter(
    AllList,
    (item) => item.dayList && item.dayList.length < 5 && item.dep === 'PE',
  )
  const mergePETList = _.filter(
    AllList,
    (item) => item.dayList && item.dayList.length < 5 && item.dep === 'PET',
  )
  let commonList = _.filter(
    AllList,
    (item) => item.dayList && item.dayList.length >= 5,
  )

  const peList = sortMergeList(mergePEList)
  const petList = sortMergeList(mergePETList)

  commonList = _.map(commonList, (o, index, collection) => {
    const hasIndex = _.findIndex(
      collection,
      (item) =>
        _.trim(item.name) === _.trim(o.name) && item.jobNum !== o.jobNum,
    )
    if (hasIndex > -1) {
      o.name = `${o.name}_${o.jobNum}`
    }
    return o
  })

  commonList = _.map(commonList, (o) => {
    const { name, jobNum, dep } = o
    const list = dtoMap(o)
    return list
  })

  return { peList, petList, commonList }
}

module.exports = {
  readFile,
}
