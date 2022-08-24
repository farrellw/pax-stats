// Pax stats highlight rule
// Pax totals page ( tarmac )
function onOpen() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var menus = [
        {
            name: "Step 1: Populate Dates",
            functionName: "populateDates"
        }, {
            name: "Step 2: Populate Attendance",
            functionName: "populateAttendance"
        },
    ];
    spreadsheet.addMenu("F3", menus);
};

function populateStats(): void {
    const startingDatesColumnNum = 8

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    spreadsheet.getSheets().forEach(x => {
        // Add rule based formatting
        var range = x.getRange(3, startingDatesColumnNum, 1000, 30 - startingDatesColumnNum)

        // Rule that turns anything in the range with a value to green.
        var rule = SpreadsheetApp.newConditionalFormatRule()
            .whenCellNotEmpty()
            .setBackground("#90EE90")
            .setRanges([range])
            .build()

        var rules: GoogleAppsScript.Spreadsheet.ConditionalFormatRule[] = []
        rules.push(rule);
        x.setConditionalFormatRules(rules);

    })
}

function populateDates(): void {
    const apostraphePrefix = "'"
    const startingDatesColumnNum = 8

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName("REGION")
    const aoMap = retrieveAOMap()

    if (sheet) {
        var regionData = sheet.getDataRange().getValues();
        for (var i = 7; i >= 0; i--) {
            let currentDate = new Date();
            // if (currentDate.getDate() - i >= 0) {
            currentDate.setDate(currentDate.getDate() - i);
            var formattedColumnName = apostraphePrefix + Utilities.formatDate(currentDate, 'GMT', 'dd-MM')
            var col = regionData[0].indexOf(Utilities.formatDate(currentDate, 'GMT', 'dd-MM'));
            var anotherCol = regionData[0].indexOf(formattedColumnName);

            if (col == -1 && anotherCol == -1) {
                sheet.insertColumnAfter(startingDatesColumnNum - 1);
                sheet.getRange(1, startingDatesColumnNum)
                    .setValues([[
                        formattedColumnName
                    ]]).setFontWeight('bold')
                    .setFontColor('#ffffff')
                    .setBackground('#007272')
                    .setBorder(
                        true, true, true, true, null, null,
                        null,
                        SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
                sheet.getRange(2, startingDatesColumnNum).setValues([["=COUNTIF(H3:H1000,\"?*\")"]])
            }
        }
    } else {
        var ui = SpreadsheetApp.getUi();
        ui.alert("No region sheet found")
    }


    aoMap.forEach((val, key) => {
        var aoSheet = spreadsheet.getSheetByName(val.friendlyName)

        if (aoSheet) {
            var aoData = aoSheet.getDataRange().getValues();
            for (var i = 7; i >= 0; i--) {
                let currentDate = new Date();

                currentDate.setDate(currentDate.getDate() - i);
                var formattedColumnName = apostraphePrefix + Utilities.formatDate(currentDate, 'GMT', 'dd-MM')
                var col = aoData[0].indexOf(Utilities.formatDate(currentDate, 'GMT', 'dd-MM'));
                var anotherCol = aoData[0].indexOf(formattedColumnName);

                let day = currentDate.getDay()
                if (val.schedule[day]) {
                    if (col == -1 && anotherCol == -1) {
                        aoSheet.insertColumnAfter(startingDatesColumnNum - 1);
                        aoSheet.getRange(1, startingDatesColumnNum)
                            .setValues([[
                                formattedColumnName
                            ]]).setFontWeight('bold')
                            .setFontColor('#ffffff')
                            .setBackground('#007272')
                            .setBorder(
                                true, true, true, true, null, null,
                                null,
                                SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

                        aoSheet.getRange(2, startingDatesColumnNum).setValues([["=COUNTIF(H3:H1000,\"?*\")"]])
                    }
                }
            }
        }
    })
}

function retrieveBB(): BD_ATTENDANCE_MAP {
    var passwordMaybe = PropertiesService.getScriptProperties().getProperty('pw');
    var userMaybe = PropertiesService.getScriptProperties().getProperty('user');

    if (!userMaybe || !passwordMaybe) {
        userMaybe = "GoingToFail"
        passwordMaybe = "GoingToFail"
    }

    var conn = Jdbc.getConnection('jdbc:mysql://f3stlouis.cac36jsyb5ss.us-east-2.rds.amazonaws.com:3306/f3stlcity', userMaybe, passwordMaybe);
    const start = new Date();
    start.setDate(start.getDate() - 24);
    const dateFormatted = Utilities.formatDate(start, 'GMT', 'YYYY-MM-dd')

    const stmt = conn.createStatement();
    stmt.setMaxRows(600);
    const results = stmt.executeQuery("SELECT * FROM bd_attendance WHERE date > '" + dateFormatted + "'");

    const bdDateIndex = 3;
    const userIndex = 1;
    const aoIndex = 2;

    let bd_attendance: BD_ATTENDANCE_MAP = new Map<USER_ID, BD_ATTENDANCE[]>()

    while (results.next()) {
        let bdDate = new Date(results.getString(bdDateIndex));
        var formattedColumnName = Utilities.formatDate(bdDate, 'GMT', 'dd-MM')

        let user = results.getString(userIndex);
        let ao = results.getString(aoIndex);

        let res = bd_attendance.get(user)
        if (res) {
            res.push({
                date: formattedColumnName,
                user: user,
                ao: ao,
            })
            bd_attendance.set(user, res)
        } else {
            bd_attendance.set(user, [
                {
                    date: formattedColumnName,
                    user: user,
                    ao: ao,
                }
            ])
        }
    }

    results.close();
    stmt.close();

    return bd_attendance
}

function retrieveAOMap(): AO_INFORMATION {
    var aoMap = new Map<string, AO>()
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    var sheet = spreadsheet.getSheetByName("AOs")
    if (sheet) {
        var values = sheet.getDataRange().getValues();

        return values.reduce((acc: AO_INFORMATION, currentValue: any[]): AO_INFORMATION => {
            if (currentValue[3] == "YES") {
                acc.set(currentValue[0], {
                    channelId: currentValue[0],
                    channel: currentValue[1],
                    shortcutName: currentValue[2],
                    friendlyName: currentValue[4],
                    schedule: [
                        currentValue[5] == "X",
                        currentValue[6] == "X",
                        currentValue[7] == "X",
                        currentValue[8] == "X",
                        currentValue[9] == "X",
                        currentValue[10] == "X",
                        currentValue[11] == "X"
                    ]
                })
            }
            return acc
        }, aoMap)
    } else {
        var ui = SpreadsheetApp.getUi();
        ui.alert("Something fucked up")
        return aoMap
    }
}

function retrieveUserMap(): Map<string, string> {
    var userMap = new Map<string, string>()

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName("USER")
    if (sheet) {
        var values = sheet.getDataRange().getValues();

        return values.reduce((acc: Map<string, string>, currentValue: any[]): Map<string, string> => {
            if (currentValue[2] == "YES") {
                acc.set(currentValue[0], currentValue[1])
            }
            return acc
        }, userMap)
    } else {
        var ui = SpreadsheetApp.getUi();
        ui.alert("Something fucked up")
        return userMap
    }
}

function populateUserMap(): void {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    spreadsheet.getSheetByName("USER_LIST")
}

function populateRegion(bdAttendance: BD_ATTENDANCE_MAP, userMap, aoMap): void {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName("REGION")
    var peopleAdded = 0
    var data = spreadsheet.getDataRange().getValues();
    var dataLength = data.length
    if (sheet) {
        Logger.log(JSON.stringify(bdAttendance))

        bdAttendance.forEach((x, key) => {
            var row = data.find(y => y[0] == key)
            var rowNum;
            if (!row) {
                peopleAdded += 1
                rowNum = peopleAdded + dataLength
            } else {
                rowNum = data.indexOf(row) + 1
            }

            if (!row) {
                sheet?.getRange(rowNum, 1, 1, 6)
                    .setValues([[
                        key,
                        userMap.get(key),
                        "",
                        "=COUNTIF(H" + rowNum + ":ZZ" + rowNum + ",\"?*\")",
                        "=COUNTIF(H" + rowNum + ":BK" + rowNum + ",\"?*\") / 8",
                        "=COUNTIF(H" + rowNum + ":U" + rowNum + ",\"?*\") / 2"
                    ]])
            }

            x.forEach(bb => {
                var col = data[0].indexOf(bb.date);

                if (col != -1) {
                    if (aoMap.get(bb.ao)) {
                        sheet?.getRange(rowNum, col + 1)
                            .setValues([[
                                aoMap.get(bb.ao).shortcutName
                            ]])
                    }
                }
            })
        })
    } else {
        var ui = SpreadsheetApp.getUi();
        ui.alert("No region sheet found")
    }
}

function statFormulaFromAO(ao: AO): Map<number, string> {
    var BD_PER_WEEK_TO_FORMULA = {
        1: {
            2: "=COUNTIF(H3:I3,\"?*\") / 2",
            8: "=COUNTIF(H3:O3,\"?*\") / 8"
        },
        2: {
            2: "=COUNTIF(H3:K3,\"?*\") / 2",
            8: "=COUNTIF(H3:W3,\"?*\") / 8"
        },
        3: {
            2: "=COUNTIF(H3:M3,\"?*\") / 2",
            8: "=COUNTIF(H3:AE3,\"?*\") / 8"
        }
    }

    return BD_PER_WEEK_TO_FORMULA[ao.schedule.filter(x => x).length]
}

function populateAO(bdAttendance: BD_ATTENDANCE_MAP, userMap, ao: AO, sheet: GoogleAppsScript.Spreadsheet.Sheet): void {
    var peopleAdded = 0
    var data = sheet.getDataRange().getValues();
    var dataLength = data.length

    var aoFormulas = statFormulaFromAO(ao)

    bdAttendance.forEach((x, key) => {
        if (x.find(xx => xx.ao == ao.channelId)) {
            var row = data.find(y => y[0] == key)
            var rowNum;
            if (!row) {
                peopleAdded += 1
                rowNum = peopleAdded + dataLength
            } else {
                rowNum = data.indexOf(row) + 1
            }

            if (!row) {
                sheet.getRange(rowNum, 1, 1, 6)
                    .setValues([[
                        key,
                        userMap.get(key),
                        "",
                        "",
                        aoFormulas ? aoFormulas[8].replaceAll("3", rowNum) : "",
                        aoFormulas ? aoFormulas[2].replaceAll("3", rowNum) : ""
                    ]])
            }
        }


        x.filter(xx => xx.ao == ao.channelId).forEach(bb => {
            var col = data[0].indexOf(bb.date);

            if (col != -1) {
                sheet.getRange(rowNum, col + 1)
                    .setValues([[
                        "X"
                    ]])
            }
        })
    })
}

function populateAttendance(): void {
    const userMap = retrieveUserMap()
    const aoMap = retrieveAOMap()
    const bdAttendance = retrieveBB();

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    populateRegion(bdAttendance, userMap, aoMap)

    aoMap.forEach((val, key) => {
        Logger.log("Here is hte AO map")
        var sheet = spreadsheet.getSheetByName(val.friendlyName)

        if (sheet) {
            populateAO(bdAttendance, userMap, val, sheet)
        }
    })
}


