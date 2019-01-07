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
import java.util.Optional;

public class AutoRowHeightCommand extends AbstractCommand {
    private static Logger logger = LoggerFactory.getLogger(AutoRowHeightCommand.class);
    private Area area;

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
        logger.debug("Cell with max value: {}. Cell value: {}", maxCell.getAddress(), maxCell.getStringCellValue());
        return maxCell;
    }

    private int calculateLineCountForCell(Cell cell) {
        logger.debug("Calculating line count for cell: {}", cell.getAddress());
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

        float cellWidth = calculateCellWidth(cell);

        while (measurer.getPosition() < cellValue.length()) {
            nextPos = measurer.nextOffset(cellWidth); // cellWidth is the max width of each line
            lineCnt++;
            measurer.setPosition(nextPos);
        }

        logger.debug("Line count: {}", lineCnt);
        return lineCnt;
        // The above solution doesn't handle the newline character, i.e. "\n", and only
        // tested under horizontal merged cells.
    }

    // Определение ширины ячейки. Если ячейка составная, то определется её суммарная ширина.
    public float calculateCellWidth(Cell cell) {
        logger.debug("Calculating cell width for cell {}", cell.getAddress());
        Optional<CellRangeAddress> cellRangeAddressOptional = getMergedRegionForCell(cell);
        if (!cellRangeAddressOptional.isPresent()) {
            return cell.getSheet().getColumnWidthInPixels(cell.getColumnIndex());
        } else {
            float result = 0;
            CellRangeAddress cellRangeAddress = cellRangeAddressOptional.get();
            Iterator<CellAddress> cellAddressIt = cellRangeAddress.iterator();
            while (cellAddressIt.hasNext()) {
                CellAddress cellAddress = cellAddressIt.next();
                Cell subCell = SheetUtil.getCell(cell.getSheet(),
                        cellAddress.getRow(),
                        cellAddress.getColumn());
                float subCellWidth = subCell.getSheet().getColumnWidthInPixels(subCell.getColumnIndex());
                logger.debug("Width of subcell {} is {}", subCell.getAddress(), subCellWidth);
                result += subCell.getSheet().getColumnWidthInPixels(subCell.getColumnIndex());
            }
            logger.debug("Result width of cell {} is {}", cell.getAddress(), result);
            return result;
        }
    }

    // Получение объединенного региона для ячейки.
    private Optional<CellRangeAddress> getMergedRegionForCell(Cell cell) {
        logger.debug("Getting merged region for cell {}", cell.getAddress());
        for (CellRangeAddress cellRangeAddress : cell.getSheet().getMergedRegions()) {
            if (cellRangeAddress.isInRange(cell)) {
                logger.debug("Cell {} is in merged region {}", cell.getAddress(), cellRangeAddress.formatAsString());
                return Optional.of(cellRangeAddress);
            }
        }
        logger.debug("Cell {} is not in a merged region", cell.getAddress());
        return Optional.empty();
    }

    @Override
    public Command addArea(Area area) {
        super.addArea(area);
        this.area = area;
        return this;
    }
}
