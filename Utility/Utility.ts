/**
 * The `Utility` namespace contains helper constants, functions, and classes
 * to faciliate the creation and maintenance of Office Scripts automations.
 */
namespace Utility {

    /**
     * Regex that matches the Excel column letter format.
     */
    export const EXCEL_COLUMN_REGEX: RegExp = /^[A-Z]{1,3}$/i

    /**
     * The character that separates the sheet name from the range
     * address in Excel.
     */
    export const EXCEL_ADDRESS_SEPARATOR: string = "!"

    /**
     * Constants dictionary that contains the sheet column limits of Excel.
     */
    export const EXCEL_BOUNDS = {
        COLUMN_START_LETTER: "A",
        COLUMN_END_LETTER: "XFD",
        MIN_NUM_COLUMNS: 0,
        MAX_NUM_COLUMNS: 16_384
    }

    /**
     * An object containing pre-defined colors from the Material UI 
     * palette, which is open-source under the MIT license.
     * The full palette can be viewed at: https://mui.com/material-ui/customization/color/
     */
    export const Colors = {
        Success: "#c8e6c9",
        Warning: "#fff9c4",
        Error: "#ffcdd2",
        Info: "#bbdefb",
        Ignore: "#f5f5f5",
        Alert: "#ffa726",
        Red: {
            50: "#ffebee",
            100: "#ffcdd2",
            200: "#ef9a9a",
            300: "#e57373",
            400: "#ef5350",
            500: "#f44336",
            600: "#e53935",
            700: "#d32f2f",
            800: "#c62828",
            900: "#b71c1c",
            a100: "#ff8a80",
            a200: "#ff5252",
            a400: "#ff1744",
            a700: "#d50000"
        },
        Pink: {
            50: "#fce4ec",
            100: "#f8bbd0",
            200: "#f48fb1",
            300: "#f06292",
            400: "#ec407a",
            500: "#e91e63",
            600: "#d81b60",
            700: "#c2185b",
            800: "#ad1457",
            900: "#880e4f",
            a100: "#ff80ab",
            a200: "#ff4081",
            a400: "#f50057",
            a700: "#c51162"
        },
        Purple: {
            50: "#f3e5f5",
            100: "#e1bee7",
            200: "#ce93d8",
            300: "#ba68c8",
            400: "#ab47bc",
            500: "#9c27b0",
            600: "#8e24aa",
            700: "#7b1fa2",
            800: "#6a1b9a",
            900: "#4a148c",
            a100: "#ea80fc",
            a200: "#e040fb",
            a400: "#d500f9",
            a700: "#aa00ff"
        },
        DeepPurple: {
            50: "#ede7f6",
            100: "#d1c4e9",
            200: "#b39ddb",
            300: "#9575cd",
            400: "#7e57c2",
            500: "#673ab7",
            600: "#5e35b1",
            700: "#512da8",
            800: "#4527a0",
            900: "#311b92",
            a100: "#b388ff",
            a200: "#7c4dff",
            a400: "#651fff",
            a700: "#6200ea"
        },
        Indigo: {
            50: "#e8eaf6",
            100: "#c5cae9",
            200: "#9fa8da",
            300: "#7986cb",
            400: "#5c6bc0",
            500: "#3f51b5",
            600: "#3949ab",
            700: "#303f9f",
            800: "#283593",
            900: "#1a237e",
            a100: "#8c9eff",
            a200: "#536dfe",
            a400: "#3d5afe",
            a700: "#304ffe"
        },
        Blue: {
            50: "#e3f2fd",
            100: "#bbdefb",
            200: "#90caf9",
            300: "#64b5f6",
            400: "#42a5f5",
            500: "#2196f3",
            600: "#1e88e5",
            700: "#1976d2",
            800: "#1565c0",
            900: "#0d47a1",
            a100: "#82b1ff",
            a200: "#448aff",
            a400: "#2979ff",
            a700: "#2962ff"
        },
        LightBlue: {
            50: "#e1f5fe",
            100: "#b3e5fc",
            200: "#81d4fa",
            300: "#4fc3f7",
            400: "#29b6f6",
            500: "#03a9f4",
            600: "#039be5",
            700: "#0288d1",
            800: "#0277bd",
            900: "#01579b",
            a100: "#80d8ff",
            a200: "#40c4ff",
            a400: "#00b0ff",
            a700: "#0091ea"
        },
        Cyan: {
            50: "#e0f7fa",
            100: "#b2ebf2",
            200: "#80deea",
            300: "#4dd0e1",
            400: "#26c6da",
            500: "#00bcd4",
            600: "#00acc1",
            700: "#0097a7",
            800: "#00838f",
            900: "#006064",
            a100: "#84ffff",
            a200: "#18ffff",
            a400: "#00e5ff",
            a700: "#00b8d4"
        },
        Teal: {
            50: "#e0f2f1",
            100: "#b2dfdb",
            200: "#80cbc4",
            300: "#4db6ac",
            400: "#26a69a",
            500: "#009688",
            600: "#00897b",
            700: "#00796b",
            800: "#00695c",
            900: "#004d40",
            a100: "#a7ffeb",
            a200: "#64ffda",
            a400: "#1de9b6",
            a700: "#00bfa5"
        },
        Green: {
            50: "#e8f5e9",
            100: "#c8e6c9",
            200: "#a5d6a7",
            300: "#81c784",
            400: "#66bb6a",
            500: "#4caf50",
            600: "#43a047",
            700: "#388e3c",
            800: "#2e7d32",
            900: "#1b5e20",
            a100: "#b9f6ca",
            a200: "#69f0ae",
            a400: "#00e676",
            a700: "#00c853"
        },
        LightGreen: {
            50: "#f1f8e9",
            100: "#dcedc8",
            200: "#c5e1a5",
            300: "#aed581",
            400: "#9ccc65",
            500: "#8bc34a",
            600: "#7cb342",
            700: "#689f38",
            800: "#558b2f",
            900: "#33691e",
            a100: "#ccff90",
            a200: "#b2ff59",
            a400: "#76ff03",
            a700: "#64dd17"
        },
        Lime: {
            50: "#f9fbe7",
            100: "#f0f4c3",
            200: "#e6ee9c",
            300: "#dce775",
            400: "#d4e157",
            500: "#cddc39",
            600: "#c0ca33",
            700: "#afb42b",
            800: "#9e9d24",
            900: "#827717",
            a100: "#f4ff81",
            a200: "#eeff41",
            a400: "#c6ff00",
            a700: "#aeea00"
        },
        Yellow: {
            50: "#fffde7",
            100: "#fff9c4",
            200: "#fff59d",
            300: "#fff176",
            400: "#ffee58",
            500: "#ffeb3b",
            600: "#fdd835",
            700: "#fbc02d",
            800: "#f9a825",
            900: "#f57f17",
            a100: "#ffff8d",
            a200: "#ffff00",
            a400: "#ffea00",
            a700: "#ffd600"
        },
        Amber: {
            50: "#fff8e1",
            100: "#ffecb3",
            200: "#ffe082",
            300: "#ffd54f",
            400: "#ffca28",
            500: "#ffc107",
            600: "#ffb300",
            700: "#ffa000",
            800: "#ff8f00",
            900: "#ff6f00",
            a100: "#ffe57f",
            a200: "#ffd740",
            a400: "#ffc400",
            a700: "#ffab00"
        },
        Orange: {
            50: "#fff3e0",
            100: "#ffe0b2",
            200: "#ffcc80",
            300: "#ffb74d",
            400: "#ffa726",
            500: "#ff9800",
            600: "#fb8c00",
            700: "#f57c00",
            800: "#ef6c00",
            900: "#e65100",
            a100: "#ffd180",
            a200: "#ffab40",
            a400: "#ff9100",
            a700: "#ff6d00"
        },
        DeepOrange: {
            50: "#fbe9e7",
            100: "#ffccbc",
            200: "#ffab91",
            300: "#ff8a65",
            400: "#ff7043",
            500: "#ff5722",
            600: "#f4511e",
            700: "#e64a19",
            800: "#d84315",
            900: "#bf360c",
            a100: "#ff9e80",
            a200: "#ff6e40",
            a400: "#ff3d00",
            a700: "#dd2c00"
        },
        Brown: {
            50: "#efebe9",
            100: "#d7ccc8",
            200: "#bcaaa4",
            300: "#a1887f",
            400: "#8d6e63",
            500: "#795548",
            600: "#6d4c41",
            700: "#5d4037",
            800: "#4e342e",
            900: "#3e2723"
        },
        Grey: {
            50: "#fafafa",
            100: "#f5f5f5",
            200: "#eeeeee",
            300: "#e0e0e0",
            400: "#bdbdbd",
            500: "#9e9e9e",
            600: "#757575",
            700: "#616161",
            800: "#424242",
            900: "#212121"
        },
        BlueGrey: {
            50: "#eceff1",
            100: "#cfd8dc",
            200: "#b0bec5",
            300: "#90a4ae",
            400: "#78909c",
            500: "#607d8b",
            600: "#546e7a",
            700: "#455a64",
            800: "#37474f",
            900: "#263238"
        }
    }

    export class WorkbookExtensions {

        private workbook: ExcelScript.Workbook

        /**
         * Fills a range with a specified color.
         * 
         * @param range The range to highlight.
         * @param color The color to highlight the range with.
         */
        public static Highlight(range: ExcelScript.Range, color: string) {
            range.getFormat().getFill().setColor(color)
        }

        /**
         * Determines if a range is empty.
         * 
         * @param range The range to evaluate.
         * @returns `true` if the range is empty and `false` otherwise.
         */
        public static IsEmpty(range: ExcelScript.Range): boolean {

            const values: ExcelScript.RangeValueType[][] = range.getValueTypes()

            for (const row of values) {
                for (const cell of row) {
                    if (cell !== ExcelScript.RangeValueType.empty) return false
                }
            }

            return true

        }

        /**
         * Determines if a range is not empty.
         * 
         * @param range The range to evaluate.
         * @returns `true` if the range is not empty and `false` otherwise.
         */
        public static IsNotEmpty(range: ExcelScript.Range): boolean {
            return !WorkbookExtensions.IsEmpty(range)
        }

        /**
         * Determines if a column letter address (such as "BGC") is valid.
         * Case sensitive.
         * 
         * @param address The column address to evaluate.
         * @returns `true` if the address is valid. Returns `false` otherwise.
         */
        public static IsValidColumnLetter(address: string): boolean {
            return EXCEL_COLUMN_REGEX.test(address) &&
                address >= EXCEL_BOUNDS.COLUMN_START_LETTER &&
                address <= EXCEL_BOUNDS.COLUMN_END_LETTER
        }

        /**
         * Determines if a column index is valid.
         * 
         * @param index The column index to evaluate.
         * @returns `true` if the column index is valid, `false` otherwise.
         */
        public static IsValidColumnIndex(index: number): boolean {
            return Number.isInteger(index) &&
                index >= EXCEL_BOUNDS.MIN_NUM_COLUMNS &&
                index <= EXCEL_BOUNDS.MAX_NUM_COLUMNS
        }

        /**
         * Gets the column index by either column letter or header name.
         * 
         * @column The column letter or header name to locate.
         * @returns The index of the column, if found.
         */
        public GetColumnIndex(column: string | number): number {

            if (typeof column === "number") {
                if (WorkbookExtensions.IsValidColumnIndex(column)) {
                    return column
                }
                else {
                    throw new Error(`${column} is not a valid column index number.`)
                }
            }

            if (WorkbookExtensions.IsValidColumnLetter(column)) {
                return this.GetColumnIndexByLetter(column)
            }

            return this.GetColumnIndexByName(column)

        }

        /**
         * Returns the address of a column whose first row value matches the name parameter, if found.
         * 
         * @param name The name of the column to search for.
         * @param mode Optional parameter search mode. Case sensitive by default.
         * @returns The address of the matching header cell, if found.
         * @throws `Error` if not found.
         */
        public GetColumnIndexByName(name: string, criteria?: ExcelScript.SearchCriteria): number {

            const results: ExcelScript.Range = this.GetHeaderRow()
                .find(name, criteria)

            if (!results) {
                throw new Error(`No such column with header "${name}" found.`)
            }

            return results.getColumnIndex()

        }

        /**
         * Gets the first row in the used range of the active worksheet.
         * 
         * @returns The header row.
         */
        public GetHeaderRow(): ExcelScript.Range {
            return this.workbook
                .getActiveWorksheet()
                .getUsedRange()
                .getRow(0)
        }

        /**
         * Gets the index number of a column address, such as "Sheet1!H35".
         * 
         * @returns The index number of the column at the provided address.
         */
        public GetColumnIndexByAddress(address: string): number {
            return this.workbook
                .getActiveWorksheet()
                .getRange(address)
                .getColumnIndex()
        }

        /**
         * Gets the column index from a column letter, such as "AC".
         * 
         * @param letter The column letter to evaluate.
         * @throws `Error` if the input is not a valid Excel column letter.
         */
        public GetColumnIndexByLetter(letter: string): number {

            if (!WorkbookExtensions.IsValidColumnLetter(letter)) {
                throw new Error(`Input "${letter}" is not a valid column.`)
            }

            return this.workbook
                .getActiveWorksheet()
                .getRange(`${letter}1`)
                .getColumnIndex()

        }

        /**
         * Inserts a column at the specified position and
         * optionally adds a header label.
         * 
         * @param position The position to insert the column. The function accepts
         * the index, letter, or name of the column.
         * @param name The header label of the new column. Blank if left empty.
         */
        public InsertColumn(position: string | number, name?: string) {

            const index: number = this.GetColumnIndex(position)
            const worksheet = this.workbook.getActiveWorksheet()

            worksheet
                .getCell(0, index)
                .insert(ExcelScript.InsertShiftDirection.right)

            if (name) {
                worksheet
                    .getCell(0, index)
                    .setValue(name)
            }

        }

        public constructor(workbook: ExcelScript.Workbook) {
            this.workbook = workbook
        }

    }

}