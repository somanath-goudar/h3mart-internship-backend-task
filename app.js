const express = require('express')
const upload = require('express-fileupload')

const XLSX = require('xlsx')
const fetch = require('cross-fetch')

const app = express()

app.use(upload())
app.use(express.static('public'))

app.get('/', (req, res) => {
  res.sendFile(__dirname + '/index.html')
})

app.post('/', (req, res) => {
  if (req.files) {
    console.log(req.files)
    let file = req.files.file
    let filename = file.name
    console.log(filename)

    file.mv('./uploads/' + filename, function (err) {
      if (err) {
        res.send(err)
      } else {
        res.send('File Uploaded <a href="product_list.xlsx">Download The Updated Excel File</a>')

        // Once File Uploaded Successfully... Fetch for Prices and Update
        let workbook = XLSX.readFile('./uploads/product_list.xlsx')
        let worksheets = {}

        for (const sheetName of workbook.SheetNames) {
          worksheets[sheetName] = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName])
        }

        let data = worksheets.Sheet1.map((obj) => obj.product_code)

        // get prices and update to excel file
        async function getPrices() {
          fetchedPrices = []
          for (let i = 0; i < data.length; i++) {
            const response = await fetch(`https://api.storerestapi.com/products/${data[i]}`)

            let data2 = await response.json()

            fetchedPrices.push(data2.data['price'])
          }

          let arrayOfObject = data.map(function (value, index) {
            return [value, fetchedPrices[index]]
          })

          headings = [['product_code', 'price']]

          const ws = XLSX.utils.json_to_sheet(arrayOfObject, { origin: 'A2', skipHeader: true })
          XLSX.utils.sheet_add_aoa(ws, headings)
          const wb = XLSX.utils.book_new()
          XLSX.utils.book_append_sheet(wb, ws, 'Sheet1')
          XLSX.writeFile(wb, './public/product_list.xlsx')
        }

        getPrices()
      }
    })
  }
})

app.listen(3000)
