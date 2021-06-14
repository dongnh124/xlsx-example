const xlsx = require('xlsx')

const file = xlsx.readFile('./assets/product.xlsx')

const wb = xlsx.utils.book_new()
wb.SheetNames.push("Product")
const ws_data = [ [ 'Handle', 'Title', 'Option1 Value', 'Option2 Value', 'Option3 Value', 'Variant SKU', 'custom_fields["cas"]', 'custom_fields["specs"]', 'custom_fields["reactivity"]'] ]

const START = 9
const END = 256
const keys = Object.keys(file.Sheets.Sheet1)
const data = []
const specs = [
  { key: 'Formula', value: 'N' },
  { key: 'M.W.', value: 'O' },
  { key: 'Purity', value: 'P' },
  { key: 'Solubility', value: 'Q' },
  { key: 'Laser line', value: 'R' },
  { key: 'Common filter set', value: 'S' },
  { key: 'Emission', value: 'T' },
  { key: 'Spectrally similar Dyes', value: 'U' },
  { key: 'Excitation maximum (nm)', value: 'V' },
  { key: 'Extinction coefficient Atexcitation maximum (Lmol-1cm-1)', value: 'W' },
  { key: 'Emission maximum (nm)', value: 'X' },
  { key: 'Fluorescence quantum yield', value: 'Y' },
  { key: 'CF260', value: 'Z' },
  { key: 'CF280', value: 'AA' },
]

const func = [
  { key: 'Functional Group 1', value: 'AC' },
  { key: 'Function 1', value: 'AD' },
  { key: 'Functional Group 2', value: 'AE' },
  { key: 'Function 2', value: 'AF' },
]

let products = []
// GET PRODUCT
for (let i = START; i <= END; i+= 1) {
  products.push((file.Sheets.Sheet1[`B${i}`].v || '').trim())
}
products = [...new Set(products)]
// console.log(products.length)

// GET SHEET DATA
for (let i = START; i <= END; i+= 1) {
  const rowI = keys.filter(k => k.replace(/[a-z]/gi,'')==i)
  // console.log(rowI)
  const dataI = {}
  rowI.forEach(k => {
    dataI[k.replace(/[0-9]/g,'')] = file.Sheets.Sheet1[k].w || ''
  })
  data.push(dataI)
}
// console.log(data)

const createRowData = (data, sizeOption, price) => {
  const arr = []

  arr.push(data['B'].toLowerCase().trim().replace(/ /gi, '-'))
  arr.push(null)
  arr.push(data['E'])
  if (data['D']) {
    arr.push(data['D'])
    arr.push(sizeOption)
  } else {
    arr.push(sizeOption)
    arr.push(null)
  }
  arr.push(data['F'])
  arr.push(data['G'])
  const specsObj = []
  specs.forEach(s => {
    if (data[s.value] && data[s.value] !== '0') {
      specsObj.push({
        key: s.key,
        value: String(data[s.value]),
      })
    }
  })
  arr.push(JSON.stringify(specsObj))
  const funcObj = []
  func.forEach(s => {
    if (data[s.value] && data[s.value] !== '0') {
      funcObj.push({
        key: s.key,
        value: String(data[s.value]),
      })
    }
  })
  arr.push(JSON.stringify(funcObj))
  return arr
}

const createProductRow = (data) => {
  const arr = []

  arr.push(data['B'].toLowerCase().trim().replace(/ /gi, '-'))
  arr.push(data['B'])
  arr.push(null)
  arr.push(null)
  arr.push(null)
  arr.push(null)
  arr.push(null)
  arr.push(null)
  arr.push(null)
  return arr
}

// PARSE DATA
products.forEach(title => {
  const productData = data.filter(d => d['B'].trim() === title)
  ws_data.push(createProductRow(productData[0]))
  productData.forEach((d, i) => {
    const price1 = Number(String(d['I'] || 0).replace('USD', '').replace(',', '').trim())
    const price2 = Number(String(d['K'] || 0).replace('USD', '').replace(',', '').trim())
    const price3 = Number(String(d['M'] || 0).replace('USD', '').replace(',', '').trim())
    price1 && ws_data.push(createRowData(d, d['H'], price1))
    price2 && ws_data.push(createRowData(d, d['J'], price2))
    price3 && ws_data.push(createRowData(d, d['L'], price3))
  })
})

const ws = xlsx.utils.aoa_to_sheet(ws_data);
wb.Sheets[ "Product" ] = ws;
// console.log(worksheetJson)
xlsx.writeFile(wb, 'product_metafields.csv')