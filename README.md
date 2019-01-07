Problem: wrapping cell content in Libre Office and Open Office doesn't work. It seems like a bug:
https://bugs.documentfoundation.org/show_bug.cgi?id=62268

Decision from:
https://stackoverflow.com/questions/39437194/jxls-auto-fit-row-height-according-to-the-content
also doesn't work for Libre Office and Open Office.

Doesn't work:

    CellStyle style = cell.getCellStyle()
    style.setWrapText(true)
    cell.setCellStyle(style)

Doesn't work:

    currentRow.setHeight((short)-1)

I found example for calculating row height by it's content:
https://stackoverflow.com/questions/19145628/auto-size-height-for-rows-in-apache-poi
This decision works correct for non-merged cells. If cell is merged, this example doesn't work correctly.

My example determines if cell is merged, and then calculate cell width for merged cells.