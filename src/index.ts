import {NewmanRunExecutionAssertion, NewmanRunOptions, NewmanRunSummary} from "newman";
import {EventEmitter} from "events";
import { Workbook, Worksheet, Column } from 'exceljs';
import { mkdirSync } from "fs";
import { dirname } from "path";
import {autoWidth} from "./util/worksheet";

interface XslxOptions {
    export: string;
}

interface NewmanEventEmitter extends EventEmitter {
    summary: NewmanRunSummary
    exports: any[]
}

enum State {
    OK = "OK",
    FAIL = "FAIL"
}

class ColumnDefinition {
    static readonly Test = new ColumnDefinition({
        header: 'Test',
        key: 'test'
    });

    static readonly State = new ColumnDefinition({
        header: 'State',
        key: 'state'
    });

    static readonly Error = new ColumnDefinition({
        header: 'Error',
        key: 'error'
    });

    private static readonly columns: ColumnDefinition[] = [];

    private constructor(public readonly definition: Required<Pick<Column, "header" | "key">> & Partial<Pick<Column, "width">>) {
        ColumnDefinition.columns.push(this);
    }

    static values(): readonly ColumnDefinition[] {
        return ColumnDefinition.columns;
    }
}

function getColumn(worksheet: Worksheet, id: ColumnDefinition) {
    return worksheet.getColumn(id.definition.key);
}


class NewmanXslxReporter {
    private readonly workbook : Workbook;
    private readonly assertionSheet : Worksheet;

    public constructor(
        newman: NewmanEventEmitter,
        reporterOptions: XslxOptions,
        collectionRunOptions: NewmanRunOptions,
    ) {
        this.workbook = new Workbook();
        this.assertionSheet = this.workbook.addWorksheet('Assertions');

        this.createColumns();

        const stateColumn = getColumn(this.assertionSheet, ColumnDefinition.State);
        stateColumn.protection = { locked: false };

        newman.on('beforeDone', async () => {
            const executions = newman.summary?.run?.executions ?? [];

            for (const execution of executions) {
                const assertions = execution.assertions ?? [];
                for (const assertion of assertions) {
                    this.addAssertion(assertion);
                }
            }

            autoWidth(this.assertionSheet);

            stateColumn.eachCell(cell => {
                cell.dataValidation = {
                    type: 'list',
                    formulae: ['"' + Object.values(State).join(',') + '"']
                } ;
            })

            this.conditionalFormatting(stateColumn);

            await this.assertionSheet.protect('', {});

            await this.write(reporterOptions);
        });
    }

    private createColumns(): void {
        this.assertionSheet.columns = ColumnDefinition.values().map(col => col.definition);
    }

    private addAssertion(assertion: NewmanRunExecutionAssertion) {
        const error = assertion?.error?.message;

        // For whatever reason the assertion field actually contains the test name...
        const test = assertion.assertion;

        this.assertionSheet.addRow({
            test: test,
            state: error ? State.FAIL : State.OK,
            error: error
        });
    }

    private async write(reporterOptions: XslxOptions) {
        const timestamp = new Date().toISOString().replace(/[^\d]+/g, '-');
        const filePath = reporterOptions.export ?? `newman/newman-run-report-${timestamp}0.xlsx`;

        const directory = dirname(filePath)

        if (directory) {
            mkdirSync(directory, {recursive: true});
        }

        await this.workbook.xlsx.writeFile(filePath);
    }
    
    private conditionalFormatting(stateColumn: Column) {
        const lastColumn = this.assertionSheet.lastColumn;
        const lastRow = this.assertionSheet.lastRow;
        
        if (!lastRow) {
           // Nothing to format if there are no rows
           return;
        }
        
        this.assertionSheet.addConditionalFormatting({
            ref: 'A2:' + lastColumn.letter + lastRow.number,
            rules: [
                {
                    type: "expression",
                    formulae: ['$' + stateColumn.letter + '2="' + State.OK + '"'],
                    priority: 0,
                    style: {fill: {type: 'pattern', pattern: 'solid', bgColor: {argb: '00c800'}}},
                },
                {
                    type: "expression",
                    formulae: ['$' + stateColumn.letter + '2="' + State.FAIL + '"'],
                    priority: 0,
                    style: {fill: {type: 'pattern', pattern: 'solid', bgColor: {argb: 'c80000'}}},
                }
            ]
        })
    }
}

export = NewmanXslxReporter;
