import {Worksheet, Column} from "exceljs";

/**
 * Autofit columns by width
 *
 * @param worksheet The worksheet
 * @param minimalWidth The minimal width
 */
export function autoWidth(worksheet: Worksheet, minimalWidth = 10) : void {
    worksheet.columns.forEach((column: Partial<Column>) => {
    
    	// Required by typescript compiler because partial marks every field as optional
        if (!column.eachCell) {
            return;
        }
        
        let maxColumnLength = 0;
        column.eachCell({ includeEmpty: true }, (cell) => {
            maxColumnLength = Math.max(
                maxColumnLength,
                minimalWidth,
                cell.value ? cell.value.toString().length : 0
            );
        });
        column.width = maxColumnLength + 2;
    });
};
