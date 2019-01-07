package ru.davtsin.jxls;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.SheetUtil;
import org.jxls.area.Area;
import org.jxls.command.AbstractCommand;
import org.jxls.command.Command;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.Size;
import org.jxls.transform.poi.PoiTransformer;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.font.FontRenderContext;
import java.awt.font.LineBreakMeasurer;
import java.awt.font.TextAttribute;
import java.text.AttributedString;
import java.util.Iterator;
import java.util.List;
import java.util.Optional;

public class AutoRowHeightCommand extends AbstractCommand {

    // TODO set logger
    private static Logger logger = LoggerFactory.getLogger(AutoRowHeightCommand.class);
    private Area area;
    //    private Cell cellWithMaxValue;
//    private Float cellWidth;
    private Integer lineCountForCell;

    @Override
    public String getName() {
        return "autoRowHeight";
    }

    @Override
    public Size applyAt(CellRef cellRef, Context context) {
        Size size = this.area.applyAt(cellRef, context);
        PoiTransformer transformer = (PoiTransformer) area.getTransformer();
        Row row = transformer.getWorkbook().getSheet(cellRef.getSheetName()).getRow(cellRef.getRow());

        Cell cell = getCellWithMaxValueInRow(row, cellRef);

        cell.getCellStyle().setWrapText(true);
        row.setHeight((short) (row.getHeight() * calculateLineCountForCell(cell)));

        return size;
    }

    private Cell getCellWithMaxValueInRow(Row row, CellRef cellRef) {
        logger.debug("Searching cell with max value in a row: {}", row.getRowNum());
        Cell maxCell = row.getCell(cellRef.getCol());
        maxCell.setCellType(CellType.STRING);
        for (Cell cell : row) {
            cell.setCellType(CellType.STRING);
            if (cell.getStringCellValue().length() > maxCell.getStringCellValue().length()) {
                maxCell = cell;
            }
        }
        System.out.println("Cell with max value is: " + maxCell.getAddress());
        System.out.println("Cell value is: " + maxCell.getStringCellValue());
        return maxCell;
    }

    private int calculateLineCountForCell(Cell cell) {
        // Create Font object with Font attribute (e.g. Font family, Font size, etc) for calculation
        java.awt.Font currFont = new java.awt.Font("Calibri", 0, 11);
        String cellValue = cell.getStringCellValue();
        AttributedString attrStr = new AttributedString(cellValue);
        attrStr.addAttribute(TextAttribute.FONT, currFont);

        // Use LineBreakMeasurer to count number of lines needed for the text
        FontRenderContext frc = new FontRenderContext(null, true, true);
        LineBreakMeasurer measurer = new LineBreakMeasurer(attrStr.getIterator(), frc);
        int nextPos = 0;
        int lineCnt = 0;

        while (measurer.getPosition() < cellValue.length()) {
            nextPos = measurer.nextOffset(calculateCellWidth(cell)); // cellWidth is the max width of each line
            lineCnt++;
            measurer.setPosition(nextPos);
        }

        System.out.println("Line count: " + lineCnt);
        return lineCnt;
        // The above solution doesn't handle the newline character, i.e. "\n", and only
        // tested under horizontal merged cells.
    }

    // определение щирины ячейки
    public float calculateCellWidth(Cell cell) {
        Optional<CellRangeAddress> cellRangeAddressOptional = getMergedRegionForCell(cell);
        if (!cellRangeAddressOptional.isPresent()) {
            return cell.getSheet().getColumnWidthInPixels(cell.getColumnIndex());
        } else {
            float result = 0;
            CellRangeAddress cellRangeAddress = cellRangeAddressOptional.get();
            System.out.println("Cell address: " + cellRangeAddress.formatAsString());
            if (cellRangeAddress.getNumberOfCells() > 1) {
                System.out.println("Cell " + cell.getAddress() + " is a merged cell");
                Iterator<CellAddress> cellAddressIt = cellRangeAddress.iterator();
                while (cellAddressIt.hasNext()) {
                    CellAddress cellAddress = cellAddressIt.next();
                    Cell subCell = SheetUtil.getCell(cell.getSheet(),
                            cellAddress.getRow(),
                            cellAddress.getColumn());
                    float subCellWidth = subCell.getSheet().getColumnWidthInPixels(subCell.getColumnIndex());
                    System.out.println("Width of subcell " + subCell.getAddress() + ": " + subCellWidth);
                    result += subCell.getSheet().getColumnWidthInPixels(subCell.getColumnIndex());
                }
                System.out.println("Result width of " + cell.getAddress() + " is: " + result);
            }
            return result;
        }
    }

    // получение объединенного региона, который занимает ячейка
    private Optional<CellRangeAddress> getMergedRegionForCell(Cell cell) {
        System.out.println("Get merged region for cell: " + cell);
        Optional<CellRangeAddress> result = Optional.empty();
        if (!isCellInMergedRegion(cell)) {
            return result;
        } else {
            for (CellRangeAddress cellRangeAddress : cell.getSheet().getMergedRegions()) {
                if (cellRangeAddress.isInRange(cell)) {
                    result = Optional.of(cellRangeAddress);
                }
            }
        }
        return result;
    }

    // является ли ячейка составной
    private boolean isCellInMergedRegion(Cell cell) {
        System.out.println("Is cell: " + cell.getAddress() + " in merged region?");
        List<CellRangeAddress> cellRangeAddresses = cell.getSheet().getMergedRegions();
        System.out.print("All merged regions: ");
        cellRangeAddresses.forEach(cellRangeAddress ->
                System.out.print(cellRangeAddress.formatAsString() + ", "));
        System.out.println();

        for (CellRangeAddress cellRangeAddress : cellRangeAddresses) {
            if (cellRangeAddress.isInRange(cell)) {
                System.out.println("Cell " + cell.getAddress() + " is in merged region: " + cellRangeAddress.formatAsString());
                return true;
            }
        }
        System.out.println("Cell is not in a merged region");
        return false;
    }

    @Override
    public Command addArea(Area area) {
        super.addArea(area);
        this.area = area;
        return this;
    }
}
