const xlsx = require('xlsx')

const file = xlsx.readFile('./assets/product.xlsx')

const wb = xlsx.utils.book_new()
wb.SheetNames.push("Product")
const ws_data = [ [ 'Handle', 'Title', 'Body (HTML)', 'Vendor', 'Type', 'Tags', 'Published', 'Option1 Name', 'Option1 Value', 'Option2 Name', 'Option2 Value', 'Option3 Name', 'Option3 Value', 'Variant SKU', 'Variant Grams', 'Variant Inventory Tracker', ' Variant Inventory Qty', 'Variant Inventory Policy', 'Variant Fulfillment Service', 'Variant Price', 'Variant Compare At Price', 'Variant Requires Shipping', 'Variant Taxable', 'Variant Barcode', 'Image Src', 'Image Position', 'Image Alt Text', 'Gift Card', 'SEO Title', 'Variant Image', 'Status'] ]

const START = 9
const END = 256
const keys = Object.keys(file.Sheets.Sheet1)
const data = []
let products = []
// GET PRODUCT
for (let i = START; i <= END; i+= 1) {
  products.push((file.Sheets.Sheet1[`B${i}`].v || '').trim())
}
products = [...new Set(products)]
console.log(products.length)

// GET SHEET DATA
for (let i = START; i <= END; i+= 1) {
  const rowI = keys.filter(k => k.replace(/[a-z]/gi,'')==i)
  // console.log(rowI)
  const dataI = {}
  rowI.forEach(k => {
    dataI[k.replace(/[0-9]/g,'')] = file.Sheets.Sheet1[k].v || ''
  })
  data.push(dataI)
}
// console.log(data)

const createRowData = (data, sizeOption, price) => {
  const arr = []

  arr.push(data['B'].toLowerCase().trim().replace(/ /gi, '-'))
  arr.push(data['B'])
  arr.push(data['C'])
  arr.push('Affolink')
  arr.push(data['A'])
  arr.push(data['A'])
  arr.push('TRUE')
  arr.push('Title')
  arr.push(data['E'])
  if (data['D']) {
    arr.push('Linker')
    arr.push(data['D'])
    arr.push('Size')
    arr.push(sizeOption)
  } else {
    arr.push('Size')
    arr.push(sizeOption)
    arr.push(null)
    arr.push(null)
  }
  arr.push(data['F'])
  arr.push(null)
  arr.push('shopify')
  arr.push(10000)
  arr.push('deny')
  arr.push('manual')
  arr.push(price)
  arr.push(null)
  arr.push('TRUE')
  arr.push('TRUE')
  arr.push(null)
  arr.push(data['AJ'])
  arr.push(null)
  arr.push(data['AJ'] ? data['E'] : null)
  arr.push('FALSE')
  arr.push(data['AI'] || null)
  arr.push(data['AJ'])
  arr.push('active')
  return arr
}

// PARSE DATA
products.forEach(title => {
  const productData = data.filter(d => d['B'].trim() === title)
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
xlsx.writeFile(wb, 'product.csv')