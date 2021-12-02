import fetch from 'node-fetch'
fetch('https://subapp.d20collective.com/task/order-csv/remove-all', {
  method: 'GET',
  headers: {
    'Accept': 'application/json',
    'Content-Type': 'application/json',
    'pass_dev': 'pass_dev',
  }
})
  .then(res => res.json())
  .then(res => console.log(res))
  .catch(er => console.error(er))