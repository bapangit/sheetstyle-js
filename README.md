# what is this?

Create sheets with styling option using javascript.

# installation

`npm i sheetstyle-js --save`

Then...

```
import Sheet from 'sheetstyle-js'

const width = 26
const x = 0
const y = 2

const sheet = new Sheet(width,x,y)

const row = { width: 9, v: `Branch : Head Office : ${BranchName}`, t: "s" },
          { width: 8, v: `Invoice Category : ${InvoiceCategory}`, t: "s" },
          { width: 9, v: `GST Category :${GNRCategory}`, t: "s" },
        ],
        styles: [
          getBorder(1, 26),
          getBorder(1, 26, "top"),
          { s: 1, e: 1, style: { font: { bold: true } } },
          { s: 10, e: 10, style: { font: { bold: true } } },
          { s: 18, e: 18, style: { font: { bold: true } } },
          {
            s: 10,
            e: 10,
            style: {
              alignment: { horizontal: "center" },
            },
          },
          {
            s: 18,
            e: 18,
            style: {
              alignment: { horizontal: "right" },
            },
          },
        ],
      }

sheet.addRow(row)

sheet .addRows([row,...])

sheet.columnStyles = [
      { cn: [1], styles: [getBorder(1, 17 + invoiceItems.length, "left")] },
      { cn: [9, 17], styles: [getBorder(1, 1, "right"), getBorder(6, 10, "right")] }]

sheet.save()
```
