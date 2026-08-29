/* global document, Excel, Office, fetch, console */

const API_URL = "http://127.0.0.1:8000";

Office.onReady((info) => {
    if (info.host !== Office.HostType.Excel) return;

    const sideloadMsg = document.getElementById("sideload-msg");
    const appBody = document.getElementById("app-body");

    if (sideloadMsg) sideloadMsg.style.display = "none";
    if (appBody) appBody.style.display = "flex";

    const writeButton = document.getElementById("write-excel");
    const readButton = document.getElementById("read-data");
    const askButton = document.getElementById("ask-ai");

    if (writeButton) writeButton.onclick = writeToExcel;
    if (readButton) readButton.onclick = readSelectedData;
    if (askButton) askButton.onclick = askAI;
});


// ============================================================
// WRITE TO EXCEL
// ============================================================

async function writeToExcel() {
    try {
        await Excel.run(async (context) => {
            const sheet =
                context.workbook.worksheets.getActiveWorksheet();

            const range = sheet.getRange("A1");

            range.values = [
                ["Hello from SheetPilot AI 🤖"]
            ];

            range.format.font.bold = true;

            await context.sync();
        });

        showResult("Excel updated successfully! ✅");

    } catch (error) {
        console.error(error);
        showError(error.message);
    }
}


// ============================================================
// READ SELECTED DATA
// ============================================================

async function readSelectedData() {
    try {
        const result = await getSelectedRangeContext();

        if (!result) {
            showError("Could not read selected Excel range.");
            return;
        }

        showResult(
            `Selected Range: ${result.address}\n\n` +
            JSON.stringify(result.values, null, 2)
        );

    } catch (error) {
        console.error(error);
        showError(error.message);
    }
}


// ============================================================
// GET SELECTED RANGE / TABLE CONTEXT
// ============================================================

async function getSelectedRangeContext() {
    try {
        return await Excel.run(async (context) => {

            const selectedRange =
                context.workbook.getSelectedRange();

            selectedRange.load([
                "address",
                "values",
                "rowCount",
                "columnCount"
            ]);

            await context.sync();

            let dataRange = selectedRange;
            let detectedTable = false;


            // If only one cell is selected,
            // automatically detect the complete table.

            if (
                selectedRange.rowCount === 1 &&
                selectedRange.columnCount === 1
            ) {

                try {

                    const currentRegion =
                        selectedRange.getCurrentRegion();

                    currentRegion.load([
                        "address",
                        "values",
                        "rowCount",
                        "columnCount"
                    ]);

                    await context.sync();

                    if (
                        currentRegion.rowCount > 1 &&
                        currentRegion.columnCount > 1
                    ) {
                        dataRange = currentRegion;
                        detectedTable = true;
                    }

                } catch (error) {
                    console.warn(
                        "Could not detect current table region:",
                        error
                    );
                }
            }


            return {
                address: dataRange.address,
                selectedAddress: selectedRange.address,
                values: dataRange.values,
                rowCount: dataRange.rowCount,
                columnCount: dataRange.columnCount,
                detectedTable: detectedTable
            };
        });

    } catch (error) {

        console.warn(
            "Could not read Excel context:",
            error
        );

        return null;
    }
}


// ============================================================
// ASK AI
// ============================================================

async function askAI() {

    const input =
        document.getElementById("ai-input");

    if (!input) return;


    const message =
        input.value.trim();

    if (!message) {
        showError("Please enter a command.");
        return;
    }


    const askButton =
        document.getElementById("ask-ai");

    if (askButton) {
        askButton.disabled = true;
        askButton.textContent = "Thinking...";
    }


    resultBoxMessage(
        "🤖 Reading Excel data..."
    );


    try {

        const selectedData =
            await getSelectedRangeContext();

        let aiMessage = message;


        if (selectedData) {

            aiMessage =
                `${message}\n\n` +

                `EXCEL CONTEXT:\n` +

                `Selected cell/range: ${selectedData.selectedAddress}\n` +

                `Full data range: ${selectedData.address}\n` +

                `Table detected: ${
                    selectedData.detectedTable
                        ? "Yes"
                        : "No"
                }\n\n` +

                `IMPORTANT:\n` +

                `Use the FULL Excel table/data below to answer the user's question. ` +

                `Do not answer using only the selected cell. ` +

                `For calculations such as total, average, highest, lowest, etc., ` +

                `calculate from ALL relevant rows in the table.\n\n` +

                `Excel table data:\n` +

                JSON.stringify(
                    selectedData.values
                        .slice(0, 100)
                        .map(row => row.slice(0, 20)),
                    null,
                    2
                );
        }


        resultBoxMessage(
            "🤖 AI is thinking..."
        );


        const data =
            await sendChatRequest({
                message: aiMessage
            });


        await processAgentResponse(
            message,
            data.result
        );


    } catch (error) {

        console.error(
            "AI Agent Error:",
            error
        );

        showError(
            "Operation failed.\n\n" +
            error.message
        );

    } finally {

        if (askButton) {
            askButton.disabled = false;
            askButton.textContent = "🚀 Ask AI";
        }
    }
}


// ============================================================
// SEND CHAT REQUEST
// ============================================================

async function sendChatRequest(payload) {

    try {

        const response =
            await fetch(
                `${API_URL}/chat`,
                {
                    method: "POST",

                    headers: {
                        "Content-Type": "application/json"
                    },

                    body: JSON.stringify(payload)
                }
            );


        if (!response.ok) {

            let errorMessage = "";

            try {

                const errorData =
                    await response.json();

                errorMessage =
                    errorData.message ||
                    errorData.detail ||
                    errorData.error ||
                    "";

            } catch (_) {}


            if (response.status >= 500) {
                throw new Error(
                    "SheetPilot AI server error. Please make sure the backend is running."
                );
            }


            if (response.status === 404) {
                throw new Error(
                    "SheetPilot AI endpoint was not found. Please check the backend."
                );
            }


            if (response.status === 400) {
                throw new Error(
                    errorMessage ||
                    "Invalid request. Please check your command."
                );
            }


            throw new Error(
                errorMessage ||
                `Request failed (HTTP ${response.status}).`
            );
        }


        const data =
            await response.json();


        if (
            data.status &&
            data.status !== "success"
        ) {

            throw new Error(
                data.message ||
                data.error ||
                "SheetPilot AI could not process your request."
            );
        }


        return data;


    } catch (error) {

        console.error(
            "SheetPilot AI request error:",
            error
        );


        if (
            error instanceof TypeError ||
            error.message === "Failed to fetch"
        ) {

            throw new Error(
                "Unable to connect to SheetPilot AI. Please make sure the backend server is running."
            );
        }


        throw error;
    }
}


// ============================================================
// PROCESS AGENT RESPONSE
// ============================================================

async function processAgentResponse(
    userMessage,
    result
) {

    // Normal AI response

    if (result.type === "message") {

        showResult(
            result.content
        );

        return;
    }


    // Tool calls

    if (
        result.type === "tool_calls" &&
        Array.isArray(result.tool_calls)
    ) {

        const toolResults = [];


        for (
            const toolCall
            of result.tool_calls
        ) {

            resultBoxMessage(
                `🤖 Executing ${toolCall.tool}...`
            );


            const executionResult =
                await executeToolCall(
                    toolCall
                );


            toolResults.push({
                tool_call_id:
                    toolCall.tool_call_id,

                tool:
                    toolCall.tool,

                result:
                    executionResult
            });
        }


        const directActions = [
            "sort_range",
            "write_range",
            "update_cell",
            "clear_range",
            "format_range",
            "add_formula",
            "create_chart"
        ];


        const successfulActions =
            toolResults
                .map(
                    item =>
                        item.result &&
                        item.result.action
                )
                .filter(Boolean);


        const hasDirectAction =
            successfulActions.some(
                action =>
                    directActions.includes(action)
            );


        if (hasDirectAction) {

            if (
                successfulActions.includes(
                    "sort_range"
                )
            ) {

                showResult(
                    "Done! Excel data sorted successfully. ✅"
                );

                return;
            }


            if (
                successfulActions.includes(
                    "create_chart"
                )
            ) {

                showResult(
                    "Done! Chart created successfully. 📊"
                );

                return;
            }


            if (
                successfulActions.includes(
                    "update_cell"
                )
            ) {

                showResult(
                    "Done! Cell updated successfully. ✅"
                );

                return;
            }


            if (
                successfulActions.includes(
                    "write_range"
                )
            ) {

                showResult(
                    "Done! Excel range updated successfully. ✅"
                );

                return;
            }


            if (
                successfulActions.includes(
                    "clear_range"
                )
            ) {

                showResult(
                    "Done! Excel range cleared successfully. ✅"
                );

                return;
            }


            if (
                successfulActions.includes(
                    "format_range"
                )
            ) {

                showResult(
                    "Done! Excel formatting updated successfully. ✅"
                );

                return;
            }


            if (
                successfulActions.includes(
                    "add_formula"
                )
            ) {

                showResult(
                    "Done! Formula added successfully. ✅"
                );

                return;
            }
        }


        // READ RANGE requires AI continuation

        const hasRead =
            toolResults.some(
                item =>
                    item.result &&
                    item.result.action ===
                    "read_range"
            );


        if (hasRead) {

            resultBoxMessage(
                "🤖 Processing Excel data..."
            );


            const continuation =
                await sendChatRequest({

                    message:
                        userMessage,

                    previous_response:
                        result,

                    tool_results:
                        toolResults
                });


            await processAgentResponse(
                userMessage,
                continuation.result
            );

            return;
        }


        showResult(
            "Excel operation completed successfully. ✅"
        );

        return;
    }


    throw new Error(
        "AI returned an unknown response."
    );
}


// ============================================================
// EXECUTE TOOL
// ============================================================

async function executeToolCall(
    toolCall
) {

    const tool =
        toolCall.tool;

    const args =
        toolCall.arguments || {};


    console.log(
        "Executing tool:",
        tool,
        args
    );


    // ========================================================
    // UPDATE CELL
    // ========================================================

    if (
        tool === "update_cell"
    ) {

        const cell =
            args.cell;

        const value =
            args.value;


        if (!cell) {
            throw new Error(
                "AI did not provide a cell."
            );
        }


        await Excel.run(
            async context => {

                const sheet =
                    context.workbook
                        .worksheets
                        .getActiveWorksheet();


                const range =
                    sheet.getRange(cell);


                range.values = [
                    [value]
                ];


                await context.sync();
            }
        );


        return {
            success: true,
            action: "update_cell",
            cell: cell,
            value: value
        };
    }


    // ========================================================
    // WRITE RANGE
    // ========================================================

    if (
        tool === "write_range"
    ) {

        const rangeAddress =
            args.range_address;

        const values =
            args.values;


        if (!rangeAddress) {
            throw new Error(
                "AI did not provide a range."
            );
        }


        if (!Array.isArray(values)) {
            throw new Error(
                "AI did not provide valid values."
            );
        }


        await Excel.run(
            async context => {

                const sheet =
                    context.workbook
                        .worksheets
                        .getActiveWorksheet();


                const range =
                    sheet.getRange(
                        rangeAddress
                    );


                range.values =
                    values;


                await context.sync();
            }
        );


        return {
            success: true,
            action: "write_range",
            range: rangeAddress,
            values: values
        };
    }


    // ========================================================
    // READ RANGE
    // ========================================================

    if (
        tool === "read_range"
    ) {

        const rangeAddress =
            args.range_address;


        if (!rangeAddress) {
            throw new Error(
                "AI did not provide a range."
            );
        }


        return await Excel.run(
            async context => {

                const sheet =
                    context.workbook
                        .worksheets
                        .getActiveWorksheet();


                const range =
                    sheet.getRange(
                        rangeAddress
                    );


                range.load([
                    "address",
                    "values",
                    "rowCount",
                    "columnCount"
                ]);


                await context.sync();


                return {
                    success: true,
                    action: "read_range",
                    range: range.address,
                    values: range.values
                };
            }
        );
    }


    // ========================================================
    // CLEAR RANGE
    // ========================================================

    if (
        tool === "clear_range"
    ) {

        const rangeAddress =
            args.range_address;


        if (!rangeAddress) {
            throw new Error(
                "AI did not provide a range."
            );
        }


        await Excel.run(
            async context => {

                const sheet =
                    context.workbook
                        .worksheets
                        .getActiveWorksheet();


                const range =
                    sheet.getRange(
                        rangeAddress
                    );


                range.clear(
                    Excel.ClearApplyTo.contents
                );


                await context.sync();
            }
        );


        return {
            success: true,
            action: "clear_range",
            range: rangeAddress
        };
    }


    // ========================================================
    // FORMAT RANGE
    // ========================================================

    if (
        tool === "format_range"
    ) {

        const rangeAddress =
            args.range_address;

        const formatType =
            args.format_type;

        const value =
            args.value;


        if (!rangeAddress) {
            throw new Error(
                "AI did not provide a range."
            );
        }


        await Excel.run(
            async context => {

                const sheet =
                    context.workbook
                        .worksheets
                        .getActiveWorksheet();


                const range =
                    sheet.getRange(
                        rangeAddress
                    );


                if (
                    formatType === "bold"
                ) {

                    range.format.font.bold =
                        value === true ||
                        value === "true";
                }


                else if (
                    formatType === "italic"
                ) {

                    range.format.font.italic =
                        value === true ||
                        value === "true";
                }


                else if (
                    formatType === "fill_color"
                ) {

                    range.format.fill.color =
                        normalizeColor(
                            value,
                            {
                                green: "#008000",
                                darkgreen: "#006100",
                                lightgreen: "#C6EFCE",
                                yellow: "#FFF2CC",
                                lightyellow: "#FFFF00",
                                blue: "#4472C4",
                                lightblue: "#D9EAF7",
                                red: "#FF0000",
                                lightred: "#FFC7CE",
                                white: "#FFFFFF",
                                black: "#000000",
                                gray: "#D9E1F2",
                                grey: "#D9E1F2"
                            }
                        );
                }


                else if (
                    formatType === "font_color"
                ) {

                    range.format.font.color =
                        normalizeColor(
                            value,
                            {
                                green: "#008000",
                                darkgreen: "#006100",
                                blue: "#4472C4",
                                red: "#FF0000",
                                white: "#FFFFFF",
                                black: "#000000",
                                gray: "#666666",
                                grey: "#666666"
                            }
                        );
                }


                else if (
                    formatType === "font_size"
                ) {

                    const fontSize =
                        Number(value);


                    if (
                        Number.isNaN(
                            fontSize
                        )
                    ) {

                        throw new Error(
                            "Invalid font size."
                        );
                    }


                    range.format.font.size =
                        fontSize;
                }


                else if (
                    formatType === "number_format"
                ) {

                    range.load([
                        "rowCount",
                        "columnCount"
                    ]);


                    await context.sync();


                    range.numberFormat =
                        Array.from(
                            {
                                length:
                                    range.rowCount
                            },

                            () =>
                                Array.from(
                                    {
                                        length:
                                            range.columnCount
                                    },

                                    () => value
                                )
                        );
                }


                else if (
                    formatType ===
                    "horizontal_alignment"
                ) {

                    const alignment =
                        String(value)
                            .toLowerCase();


                    if (
                        ![
                            "left",
                            "center",
                            "right"
                        ].includes(
                            alignment
                        )
                    ) {

                        throw new Error(
                            "Alignment must be left, center, or right."
                        );
                    }


                    range.format.horizontalAlignment =
                        alignment;
                }


                else if (
                    formatType === "borders"
                ) {

                    const borderIndices = [

                        Excel.BorderIndex.edgeTop,

                        Excel.BorderIndex.edgeBottom,

                        Excel.BorderIndex.edgeLeft,

                        Excel.BorderIndex.edgeRight

                    ];


                    for (
                        const borderIndex
                        of borderIndices
                    ) {

                        const border =
                            range.format.borders
                                .getItem(
                                    borderIndex
                                );


                        border.style =
                            Excel.BorderLineStyle.continuous;


                        border.weight =
                            Excel.BorderWeight.thin;
                    }
                }


                else if (
                    formatType === "autofit_columns"
                ) {

                    range.format.autofitColumns();
                }


                else if (
                    formatType === "autofit_rows"
                ) {

                    range.format.autofitRows();
                }


                else {

                    throw new Error(
                        `Unsupported format: ${formatType}`
                    );
                }


                await context.sync();
            }
        );


        return {
            success: true,
            action: "format_range",
            range: rangeAddress,
            format_type: formatType,
            value: value
        };
    }


    // ========================================================
    // ADD FORMULA
    // ========================================================

    if (
        tool === "add_formula"
    ) {

        const cell =
            args.cell;

        const formula =
            args.formula;


        if (!cell) {
            throw new Error(
                "AI did not provide formula cell."
            );
        }


        if (!formula) {
            throw new Error(
                "AI did not provide formula."
            );
        }


        await Excel.run(
            async context => {

                const sheet =
                    context.workbook
                        .worksheets
                        .getActiveWorksheet();


                const range =
                    sheet.getRange(cell);


                range.formulas = [
                    [formula]
                ];


                await context.sync();
            }
        );


        return {
            success: true,
            action: "add_formula",
            cell: cell,
            formula: formula
        };
    }


    // ========================================================
    // SORT RANGE
    // ========================================================

    if (
        tool === "sort_range"
    ) {

        return await sortRange(args);
    }


    // ========================================================
    // CREATE CHART
    // ========================================================

    if (
        tool === "create_chart"
    ) {

        return await createSmartChart(args);
    }


    throw new Error(
        `Unsupported tool: ${tool}`
    );
}


// ============================================================
// COLOR HELPER
// ============================================================

function normalizeColor(
    value,
    colorMap
) {

    let color =
        String(value || "").trim();


    const mapped =
        colorMap[
            color.toLowerCase()
        ];


    if (mapped) {
        color = mapped;
    }


    if (
        !/^#[0-9A-Fa-f]{6}$/.test(
            color
        )
    ) {

        throw new Error(
            "Invalid color. Use a color name or #RRGGBB."
        );
    }


    return color;
}


// ============================================================
// SORT RANGE
// ============================================================

async function sortRange(args) {

    const requestedColumn =
        args.column;

    const requestedRange =
        args.range_address;

    const order =
        String(
            args.order ||
            "descending"
        )
            .trim()
            .toLowerCase();


    if (!requestedColumn) {

        throw new Error(
            "AI did not provide a sort column."
        );
    }


    return await Excel.run(
        async context => {

            const sheet =
                context.workbook
                    .worksheets
                    .getActiveWorksheet();


            let range;


            // ------------------------------------------------
            // DETERMINE RANGE
            // ------------------------------------------------

            if (
                requestedRange &&
                String(requestedRange).trim()
            ) {

                range =
                    sheet.getRange(
                        String(
                            requestedRange
                        ).trim()
                    );

            } else {

                const selectedRange =
                    context.workbook
                        .getSelectedRange();


                selectedRange.load([
                    "rowCount",
                    "columnCount"
                ]);


                await context.sync();


                if (
                    selectedRange.rowCount === 1 &&
                    selectedRange.columnCount === 1
                ) {

                    range =
                        selectedRange.getCurrentRegion();

                } else {

                    range =
                        selectedRange;
                }
            }


            range.load([
                "address",
                "values",
                "rowCount",
                "columnCount"
            ]);


            await context.sync();


            const values =
                range.values;


            if (
                !values ||
                values.length < 2
            ) {

                throw new Error(
                    "Could not detect a complete Excel table to sort. Select a table/range containing a header and data rows."
                );
            }


            const headerRow =
                values[0] || [];


            if (
                headerRow.length < 2
            ) {

                throw new Error(
                    "The detected Excel table does not contain enough columns."
                );
            }


            // ------------------------------------------------
            // NORMALIZE COLUMN NAME
            // ------------------------------------------------

            const normalize =
                value =>
                    String(
                        value ?? ""
                    )
                        .trim()
                        .toLowerCase()
                        .replace(
                            /[\s_-]+/g,
                            ""
                        )
                        .replace(
                            /[^a-z0-9]/g,
                            ""
                        );


            const requested =
                normalize(
                    requestedColumn
                );


            let columnIndex = -1;


            // ------------------------------------------------
            // 1. EXACT HEADER MATCH
            // ------------------------------------------------

            for (
                let i = 0;
                i < headerRow.length;
                i++
            ) {

                if (
                    normalize(
                        headerRow[i]
                    ) === requested
                ) {

                    columnIndex =
                        i;

                    break;
                }
            }


            // ------------------------------------------------
            // 2. SINGULAR / PLURAL MATCH
            // ------------------------------------------------

            if (
                columnIndex === -1
            ) {

                for (
                    let i = 0;
                    i < headerRow.length;
                    i++
                ) {

                    const header =
                        normalize(
                            headerRow[i]
                        );


                    if (
                        header ===
                            `${requested}s` ||

                        requested ===
                            `${header}s` ||

                        header ===
                            requested.replace(
                                /s$/,
                                ""
                            ) ||

                        requested ===
                            header.replace(
                                /s$/,
                                ""
                            )
                    ) {

                        columnIndex =
                            i;

                        break;
                    }
                }
            }


            // ------------------------------------------------
            // 3. COLUMN NUMBER
            // Example: 3 = third column
            // ------------------------------------------------

            if (
                columnIndex === -1 &&
                /^\d+$/.test(
                    String(
                        requestedColumn
                    ).trim()
                )
            ) {

                const index =
                    Number(
                        requestedColumn
                    ) - 1;


                if (
                    index >= 0 &&
                    index < headerRow.length
                ) {

                    columnIndex =
                        index;
                }
            }


            // ------------------------------------------------
            // 4. EXCEL COLUMN LETTER
            // Example: C / D
            // ------------------------------------------------

            if (
                columnIndex === -1 &&
                /^[A-Za-z]+$/.test(
                    String(
                        requestedColumn
                    ).trim()
                )
            ) {

                const letters =
                    String(
                        requestedColumn
                    )
                        .trim()
                        .toUpperCase();


                let absoluteColumn = 0;


                for (
                    let i = 0;
                    i < letters.length;
                    i++
                ) {

                    absoluteColumn =
                        absoluteColumn * 26 +
                        letters.charCodeAt(i) -
                        64;
                }


                const firstCell =
                    range.address.match(
                        /!?(?:\$)?([A-Za-z]+)(?:\$)?\d+/
                    );


                if (firstCell) {

                    const firstLetters =
                        firstCell[1]
                            .toUpperCase();


                    let firstAbsoluteColumn =
                        0;


                    for (
                        let i = 0;
                        i < firstLetters.length;
                        i++
                    ) {

                        firstAbsoluteColumn =
                            firstAbsoluteColumn *
                                26 +
                            firstLetters.charCodeAt(
                                i
                            ) -
                            64;
                    }


                    const relativeIndex =
                        absoluteColumn -
                        firstAbsoluteColumn;


                    if (
                        relativeIndex >= 0 &&
                        relativeIndex <
                            headerRow.length
                    ) {

                        columnIndex =
                            relativeIndex;
                    }
                }
            }


            // ------------------------------------------------
            // COLUMN NOT FOUND
            // ------------------------------------------------

            if (
                columnIndex < 0 ||
                columnIndex >=
                    headerRow.length
            ) {

                const available =
                    headerRow
                        .map(
                            (header, index) =>
                                `${index + 1}. ${header}`
                        )
                        .join(", ");


                throw new Error(
                    `Could not find sort column "${requestedColumn}". Available columns: ${available}`
                );
            }


            // ------------------------------------------------
            // SORT ORDER
            // ------------------------------------------------

            const ascending =
                [
                    "ascending",
                    "asc",
                    "a-z",
                    "lowest to highest",
                    "low to high"
                ].includes(
                    order
                );


            // ------------------------------------------------
            // SORT COMPLETE ROWS
            // ------------------------------------------------

            range.sort.apply(
                [
                    {
                        key:
                            columnIndex,

                        ascending:
                            ascending
                    }
                ],

                false,

                true
            );


            await context.sync();


            return {

                success:
                    true,

                action:
                    "sort_range",

                range:
                    range.address,

                column:
                    requestedColumn,

                matched_column:
                    headerRow[
                        columnIndex
                    ],

                order:
                    ascending
                        ? "ascending"
                        : "descending"
            };
        }
    );
}


// ============================================================
// CREATE SMART CHART
// ============================================================

async function createSmartChart(
    args
) {

    const rangeAddress =
        args.range_address;

    const chartType =
        args.chart_type ||
        "ColumnClustered";

    const title =
        args.title ||
        "Excel Chart";


    if (!rangeAddress) {

        throw new Error(
            "AI did not provide chart range."
        );
    }


    return await Excel.run(
        async context => {

            const sheet =
                context.workbook
                    .worksheets
                    .getActiveWorksheet();


            const sourceRange =
                sheet.getRange(
                    rangeAddress
                );


            sourceRange.load([
                "address",
                "values",
                "rowCount",
                "columnCount"
            ]);


            await context.sync();


            const values =
                sourceRange.values;


            if (
                !values ||
                values.length < 2
            ) {

                throw new Error(
                    "Not enough data to create a chart."
                );
            }


            const headers =
                values[0];


            // ------------------------------------------------
            // CATEGORY COLUMN
            // ------------------------------------------------

            let categoryIndex =
                findHeaderIndex(
                    headers,
                    [
                        "name",
                        "product",
                        "customer",
                        "employee",
                        "category",
                        "item"
                    ]
                );


            if (
                categoryIndex === -1
            ) {

                categoryIndex =
                    findFirstTextColumn(
                        values
                    );
            }


            // ------------------------------------------------
            // NUMERIC COLUMN
            // ------------------------------------------------

            let valueIndex =
                findNumericHeaderIndex(
                    headers,
                    title
                );


            if (
                valueIndex === -1
            ) {

                valueIndex =
                    findFirstNumericColumn(
                        values,
                        categoryIndex
                    );
            }


            if (
                categoryIndex === -1
            ) {

                throw new Error(
                    "Could not identify a category column for the chart."
                );
            }


            if (
                valueIndex === -1
            ) {

                throw new Error(
                    "Could not identify a numeric column for the chart."
                );
            }


            if (
                categoryIndex === valueIndex
            ) {

                throw new Error(
                    "Chart category and value columns cannot be the same."
                );
            }


            // ------------------------------------------------
            // VALID ROWS
            // ------------------------------------------------

            const validRows = [];


            for (
                let rowIndex = 1;
                rowIndex < values.length;
                rowIndex++
            ) {

                const row =
                    values[rowIndex];


                const category =
                    row[categoryIndex];


                const numericValue =
                    row[valueIndex];


                if (
                    category === null ||
                    category === undefined ||
                    String(
                        category
                    ).trim() === ""
                ) {

                    continue;
                }


                if (
                    numericValue === null ||
                    numericValue === undefined ||
                    numericValue === ""
                ) {

                    continue;
                }


                if (
                    typeof numericValue !==
                        "number" &&

                    Number.isNaN(
                        Number(
                            numericValue
                        )
                    )
                ) {

                    continue;
                }


                validRows.push({
                    category:
                        category,

                    value:
                        numericValue
                });
            }


            if (
                validRows.length === 0
            ) {

                throw new Error(
                    "No valid chart data was found."
                );
            }


            // ------------------------------------------------
            // HELPER SHEET
            // ------------------------------------------------

            const helperName =
                "__SheetPilotChartData";


            let helperSheet =
                context.workbook
                    .worksheets
                    .getItemOrNullObject(
                        helperName
                    );


            await context.sync();


            if (
                !helperSheet.isNullObject
            ) {

                helperSheet.delete();

                await context.sync();
            }


            helperSheet =
                context.workbook
                    .worksheets
                    .add(
                        helperName
                    );


            const helperValues = [

                [
                    headers[valueIndex],
                    headers[categoryIndex]
                ]

            ];


            for (
                const row
                of validRows
            ) {

                helperValues.push([
                    row.value,
                    row.category
                ]);
            }


            const helperRange =
                helperSheet.getRangeByIndexes(
                    0,
                    0,
                    helperValues.length,
                    2
                );


            helperRange.values =
                helperValues;


            // ------------------------------------------------
            // CREATE CHART
            // ------------------------------------------------

            const chart =
                sheet.charts.add(
                    chartType,
                    helperRange,
                    Excel.ChartSeriesBy.columns
                );


            chart.title.text =
                title;


            chart.legend.position =
                Excel.ChartLegendPosition.right;


            chart.setPosition(
                "H10",
                "N25"
            );


            await context.sync();


            // ------------------------------------------------
            // REMOVE DEFAULT SERIES
            // ------------------------------------------------

            chart.series.load(
                "items"
            );


            await context.sync();


            for (
                const series
                of chart.series.items
            ) {

                series.delete();
            }


            await context.sync();


            // ------------------------------------------------
            // ADD NUMERIC SERIES
            // ------------------------------------------------

            const series =
                chart.series.add(
                    String(
                        headers[valueIndex]
                    )
                );


            series.setValues(
                helperSheet.getRange(
                    `A2:A${
                        validRows.length + 1
                    }`
                )
            );


            series.setXAxisValues(
                helperSheet.getRange(
                    `B2:B${
                        validRows.length + 1
                    }`
                )
            );


            await context.sync();


            // ------------------------------------------------
            // HIDE HELPER SHEET
            // ------------------------------------------------

            helperSheet.visibility =
                Excel.SheetVisibility.hidden;


            await context.sync();


            return {

                success:
                    true,

                action:
                    "create_chart",

                range:
                    rangeAddress,

                chart_type:
                    chartType,

                title:
                    title,

                category_column:
                    headers[
                        categoryIndex
                    ],

                value_column:
                    headers[
                        valueIndex
                    ],

                data_points:
                    validRows.length
            };
        }
    );
}


// ============================================================
// FIND HEADER
// ============================================================

function findHeaderIndex(
    headers,
    candidates
) {

    for (
        let i = 0;
        i < headers.length;
        i++
    ) {

        const header =
            String(
                headers[i]
            )
                .trim()
                .toLowerCase();


        if (
            candidates.includes(
                header
            )
        ) {

            return i;
        }
    }


    return -1;
}


// ============================================================
// FIND NUMERIC HEADER
// ============================================================

function findNumericHeaderIndex(
    headers,
    title
) {

    const text =
        String(
            title || ""
        )
            .toLowerCase();


    const keywords = [

        "revenue",

        "sales",

        "expense",

        "expenses",

        "profit",

        "income",

        "cost",

        "amount",

        "price",

        "salary",

        "margin"

    ];


    for (
        const keyword
        of keywords
    ) {

        if (
            !text.includes(
                keyword
            )
        ) {

            continue;
        }


        for (
            let i = 0;
            i < headers.length;
            i++
        ) {

            const header =
                String(
                    headers[i]
                )
                    .trim()
                    .toLowerCase();


            if (
                header.includes(
                    keyword
                )
            ) {

                return i;
            }
        }
    }


    return -1;
}


// ============================================================
// FIND FIRST TEXT COLUMN
// ============================================================

function findFirstTextColumn(
    values
) {

    if (
        !values ||
        values.length < 2
    ) {

        return -1;
    }


    for (
        let column = 0;
        column < values[0].length;
        column++
    ) {

        for (
            let row = 1;
            row < values.length;
            row++
        ) {

            const value =
                values[row][column];


            if (
                typeof value === "string" &&
                value.trim() !== ""
            ) {

                return column;
            }
        }
    }


    return -1;
}


// ============================================================
// FIND FIRST NUMERIC COLUMN
// ============================================================

function findFirstNumericColumn(
    values,
    excludedColumn
) {

    if (
        !values ||
        values.length < 2
    ) {

        return -1;
    }


    for (
        let column = 0;
        column < values[0].length;
        column++
    ) {

        if (
            column === excludedColumn
        ) {

            continue;
        }


        for (
            let row = 1;
            row < values.length;
            row++
        ) {

            const value =
                values[row][column];


            if (
                typeof value === "number" &&
                !Number.isNaN(value)
            ) {

                return column;
            }


            if (
                typeof value === "string" &&
                value.trim() !== "" &&
                !Number.isNaN(
                    Number(value)
                )
            ) {

                return column;
            }
        }
    }


    return -1;
}


// ============================================================
// RESULT MESSAGE
// ============================================================

function resultBoxMessage(
    message
) {

    const resultBox =
        document.getElementById(
            "ai-result"
        );


    if (resultBox) {

        resultBox.textContent =
            message;
    }
}


// ============================================================
// SHOW RESULT
// ============================================================

function showResult(
    message
) {

    const resultBox =
        document.getElementById(
            "ai-result"
        );


    if (resultBox) {

        resultBox.textContent =
            String(message)
                .replace(
                    /\*\*(.*?)\*\*/g,
                    "$1"
                );
    }


    const status =
        document.getElementById(
            "status"
        );


    if (status) {

        status.textContent =
            "Ready ✅";
    }
}


// ============================================================
// SHOW ERROR
// ============================================================

function showError(
    message
) {

    const resultBox =
        document.getElementById(
            "ai-result"
        );


    if (resultBox) {

        resultBox.textContent =
            "❌ Error:\n\n" +
            message;
    }


    const status =
        document.getElementById(
            "status"
        );


    if (status) {

        status.textContent =
            "Operation failed ❌";
    }
}