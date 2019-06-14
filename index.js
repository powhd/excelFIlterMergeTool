/**
 * @description excel批量合成工具
 * @author pow
 */


// 公共库
const xlsx = require('node-xlsx')
const fs = require('fs')
// excel文件夹路径（把要合并的文件放在excel文件夹内）
const _file = `${__dirname}/excel/`
const _output = `${__dirname}/result/`
let count = 0
// 合并数据的结果集
let dataList = [{
	name: 'financeAndTax',
	data: []
}]
function unique (arr) {
	return Array.from(new Set(arr))
}
init()
function init () {
	fs.readdir(_file, function(err, files) {
		console.log(files)
		if (err) {
			throw err
		}
		// files是一个数组
		// 每个元素是此目录下的文件或文件夹的名称
		files.forEach((item, index) => {
			try {
				if(item === '.DS_Store') {
					return
				}
				if(item.indexOf('~$') === -1){
					console.log(`开始合并：${item}`)
					let excelData = xlsx.parse(`${_file}${item}`)
					console.log('\x1B[33m%s\x1b[0m', '表头数量' + excelData[0].data[0].length)
					console.log('\x1B[33m%s\x1b[0m', '表格数量' + excelData[0].data.length)
					if (excelData) {
						if (dataList[0].data.length > 0) {
							excelData[0].data.splice(0, 1) // 去除合并表格的第一行 也就是头部
						}
						// excelData[0].data.forEach(item => {
						// 	console.log(item)
						// })
						dataList[0].data = dataList[0].data.concat(excelData[0].data)
					}
					count ++ 
				}
			} catch (e) {
				console.log(e)
				console.log('excel表格内部字段不一致，请检查后再合并。')
			}
		})
		// 写xlsx
		let tableData = dataList[0].data
		// let componyArr = []
		// dataList[0].data.forEach(item => {
		// 	componyArr.push(item[1])
		// })
		for(let i=0; i<tableData.length; i++) {
			let maxIndex = 0
			let maxLen = 0
			if(tableData[i]) {
				maxLen = tableData[i].length
				for(let j=i+1; j<tableData.length; j++) {
					if(tableData[j] && tableData[i]){
						if(tableData[i][1] === tableData[j][1]) {
							let maxOri = true // true为 i比较   false j比较
							let iLen = tableData[i].filter(item => item).length
							let jLen = tableData[j].filter(item => item).length
							// console.log(i,j)
							if(maxOri) {
								if(iLen >= jLen){
									maxIndex = i
									maxLen = iLen
									tableData.splice(j, 1, null)
								}else{
									maxIndex = j
									maxOri = false
									maxLen = jLen
									tableData.splice(i, 1, null)
								}
							}else{
								if(maxLen < jLen){
									tableData.splice(maxIndex, 1, null)
									maxIndex = j
									maxLen = jLen
								}else{
									tableData.splice(j, 1, null)
								}
							}
						}
					}
				}
			}else{
				maxIndex = 0
			}
		}
		// const data = unique(componyArr)
		// console.log(data.length)
		dataList[0].data = tableData.filter(item => item)
		var buffer = xlsx.build(dataList)
		console.log('\x1B[33m%s\x1b[0m', '所有表格数量总和' + dataList[0].data.length)
		console.log('\x1B[33m%s\x1b[0m', '成功合并数量为:' + count + '个表格')

		fs.writeFile(`${_output}financeAndTax.xlsx`, buffer, function (err) {
			if (err) {
				throw err
			}
			console.log('\x1B[33m%s\x1b[0m', `完成合并：${_output}financeAndTax.xlsx`)
		})
	})
}
