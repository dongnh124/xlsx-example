import xlsx from 'xlsx'
import fetch from 'node-fetch'

const file = xlsx.readFile('./assets/SubscriberDiceWithDups.csv')
const dataRaw = file.Sheets.Sheet1
const data = []
const keys = Object.keys(dataRaw)
const START_ROW = 2
const END_ROW = 1476

for (let i = START_ROW; i <= END_ROW; i += 1) {
  const keyInRow = keys.filter(key => key.replace(/[a-z,A-Z]/gi, '') == i)
  if (!keyInRow.length) continue

  const email = dataRaw[keyInRow[0]].v
  const skus = []
  for (let j = 1; j < keyInRow.length; j += 1) {
    skus.push(dataRaw[keyInRow[j]].v)
  }
  data.push({
    email,
    skus
  })
}

console.log(data.filter(i => i.skus.length).length)

// fetch('https://subapp.d20collective.com/task/order-csv/add', {
//   method: 'POST',
//   headers: {
//     'Accept': 'application/json',
//     'Content-Type': 'application/json',
//     'pass_dev': 'pass_dev',
//   },
//   body: JSON.stringify({ data })
// })
//   .then(res => res.json())
//   .then(res => console.log(res))
//   .catch(er => console.error(er))
