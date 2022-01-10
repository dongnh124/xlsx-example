import xlsx from 'xlsx'
import fs from 'fs'

const file = fs.readFileSync('./assets/duplicate.json')
const data = JSON.parse(file)
const wb = xlsx.utils.book_new()
wb.SheetNames.push("Product")
const ws_data = []
const url = 'https://luckystar-sneaker.myshopify.com/admin/products/'

Object.keys(data).forEach(k => {
  const uniqueProductId = [ ...new Set(data[k].map(i => i.productId)) ]
    .map(id => `${url}${id}`)
  const row = [k, ...uniqueProductId]
  ws_data.push(row)
})

const ws = xlsx.utils.aoa_to_sheet(ws_data);
Object.keys(ws).forEach(k => {
  if (ws[k].v && ws[k].v.includes('https')) {
    ws[k].l = {
      Target: ws[k].v,
    }
  }
})

wb.Sheets[ "Product" ] = ws;
xlsx.writeFile(wb, 'duplicate.xlsx')