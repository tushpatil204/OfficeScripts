function main(workbook: ExcelScript.Workbook, nonGDOWorksheetName: string, gdoWorksheetName: string) {
    // This logic assumes that the project ID is of length 12. 
    // If the project ID format changes, then this logic will need to be replaced.
    let projectIDLength = 12

    /**
     * Get non-GDO data
     */
    //let nonGDOWorksheetName = 'PCE Dollarized FTEs by skillset'
    let nonGDOSheet = workbook.getWorksheet('PCE Dollarized FTEs by skillset');
    nonGDOSheet.getRange('A:A').unmerge();
    let nonGDORange = nonGDOSheet.getUsedRange();
    let nonGDOValues = nonGDORange.getValues();

    /**
     * Get GDO data
     */
    //let gdoWorksheetName = 'GDO Dollarized';
    let gdoSheet = workbook.getWorksheet('GDO Dollarized');

    // Delete additional years
    gdoSheet.getRange("GG:IN")
        .delete(ExcelScript.DeleteShiftDirection.left);

    const gdoRange = gdoSheet.getUsedRange();
    const gdoValues = gdoRange.getValues();

    /**
     * Set offsets
     */
    const blankRowsBeforeUsedRangeGDO = 2
    const blankRowsBeforeUsedRangeNonGDO = 2
    const startingRow = 5
    const summationCellIndex = 7
    const projectIndex = 0

    let currentproject: string;
    let previousProject: string;


    /**
     * Loop through the non-GDO values
     */
    for (let row = startingRow; row < nonGDOValues.length; row++) {

        /**
        * Get the name of the project
        */
        currentproject = String(nonGDOValues[row][projectIndex]);

        /**
        * If the project has data (non-summarized), process accordingly
        */
        if (currentproject != 'Σ' && currentproject.length == projectIDLength) {

            /**
            * Check if the resource field is blank.
            */
            if (nonGDORange.getCell(row, summationCellIndex).getValue() == '') {

                /**
                 * Get the name of the previous project. We'll insert the GDO data after this row.
                 */
                previousProject = String(nonGDOValues[row - 1][projectIndex]);

                /**
                 * Add GDO data
                 */
                row =
                    addGDOData(previousProject, gdoRange, gdoValues, blankRowsBeforeUsedRangeGDO, blankRowsBeforeUsedRangeNonGDO, nonGDOSheet, gdoSheet, row, false)

                
                
           
            }
        }
        /**
         * Update non-GDO range
         */
        nonGDORange = nonGDOSheet.getUsedRange();
        nonGDOValues = nonGDORange.getValues();
    }
    

    /**
     * Update non-GDO range
     */
    nonGDORange = nonGDOSheet.getUsedRange();
    nonGDOValues = nonGDORange.getValues();


    let lastRow = nonGDOValues.length - blankRowsBeforeUsedRangeNonGDO

    /**
     * Get the last project
     */
    let lastProject: String = String(nonGDOValues[lastRow][projectIndex]);

    /**
    * Add GDO data
    */
    addGDOData(lastProject, gdoRange, gdoValues, blankRowsBeforeUsedRangeGDO, blankRowsBeforeUsedRangeNonGDO, nonGDOSheet, gdoSheet, lastRow, true)

    return;

}


function getGDORange(projectName: String, gdoRange: ExcelScript.Range, gdoValues: (string | number | boolean)[][]) {

    let st = -1, en = 0;

    for (let row = 4; row < gdoValues.length; row++) {

        let project: String = String(gdoValues[row][0]);

        if (project != 'Σ') {

            if (project == projectName) {

                if (gdoRange.getCell(row, 7).getValue() !== '') {

                    if (st == -1) {
                        st = row;
                        en = st
                    }
                    en = row;
                }
            }
        }
    }

    return {
        start: st,
        end: en
    }

}



function addGDOData(project: String, gdoRange: ExcelScript.Range, gdoValues: (string | number | boolean)[][], blankRowsBeforeUsedRangeGDO: number, blankRowsBeforeUsedRangeNonGDO: number, nonGDOSheet: ExcelScript.Worksheet, gdoSheet: ExcelScript.Worksheet, row: number, isLastRow: boolean) {
    let GDOSampleRangeObj = {
        start: 0,
        end: 0
    }

    /**
     * Get the start and end rows of the GDO data for the project.
     */
    GDOSampleRangeObj = getGDORange(project, gdoRange, gdoValues)

    /**
     * If we don't find any matching data in the GDO sheet return the current row 
     */
    if (GDOSampleRangeObj.start == -1) {
        return row;
    }

    /**
     * Get the range for the GDO data
     */
    let GDOSampleRange = GDOSampleRangeObj.start + (blankRowsBeforeUsedRangeGDO) + ':' + (GDOSampleRangeObj.end + blankRowsBeforeUsedRangeGDO)


    /**
     * Get how many rows we would need to add to place the GDO data in
     */
    let GDORangeDifference = GDOSampleRangeObj.end - GDOSampleRangeObj.start


    /**
     * Get the range of rows we would need to add
     */
    let nonGDOSampleRange: string;

    if (!isLastRow) {
        nonGDOSampleRange = (row + blankRowsBeforeUsedRangeNonGDO) + ':' + (row + GDORangeDifference + blankRowsBeforeUsedRangeNonGDO)
    } else {
        nonGDOSampleRange = (row + blankRowsBeforeUsedRangeNonGDO + 2) + ':' + (row + GDORangeDifference + blankRowsBeforeUsedRangeNonGDO + 1)
    }

    /**
     * Insert blank rows
     */
    nonGDOSheet.getRange(nonGDOSampleRange).insert(ExcelScript.InsertShiftDirection.down);

    nonGDOSheet.getRange(nonGDOSampleRange).copyFrom(gdoSheet.getRange(GDOSampleRange));


    /**
     * Previously this code used nonGDOSampleRange in the formula which gave a NaN value for row in some cases
     * i = i + nonGDOSampleRange + blankRowsBeforeUsedRangeNonGDO - 1;
     */
    row = row + GDORangeDifference + blankRowsBeforeUsedRangeNonGDO - 1;

    return row

}
