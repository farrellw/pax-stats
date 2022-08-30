// Pax totals page ( tarmac )

function onOpen() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var menus = [
        {
            name: "Populate Dates",
            functionName: "populateDates"
        }, {
            name: "Populate Attendance",
            functionName: "populateAttendance"
        }, {
            name: "Refresh User List",
            functionName: "refreshUserList"
        },
        {
            name: "Sort",
            functionName: "sortAll"
        }
    ];
    spreadsheet.addMenu("F3", menus);

    var adminHelperMenu = [
        {
            name: "setRule",
            functionName: "setRule"
        }
    ]

    spreadsheet.addMenu("ADMIN-F3", adminHelperMenu);
};

function sortAll() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    var sheet = spreadsheet.getSheetByName("REGION")
    sheet?.sort(2)
    const aoMap = retrieveAOMap()

    aoMap.forEach((val, key) => {
        var aoSheet = spreadsheet.getSheetByName(val.friendlyName)
        aoSheet?.sort(5, false)
    })
}

function refreshUserList() {
    var passwordMaybe = PropertiesService.getScriptProperties().getProperty('pw');
    var userMaybe = PropertiesService.getScriptProperties().getProperty('user');

    if (!userMaybe || !passwordMaybe) {
        userMaybe = "GoingToFail"
        passwordMaybe = "GoingToFail"
    }

    var conn = Jdbc.getConnection('jdbc:mysql://f3stlouis.cac36jsyb5ss.us-east-2.rds.amazonaws.com:3306/f3stlcity', userMaybe, passwordMaybe);

    const stmt = conn.createStatement();
    stmt.setMaxRows(500);
    const results = stmt.executeQuery("SELECT * FROM users");

    const userMap = retrieveUserMap()

    const userIdIndex = 1;
    const userNameIndex = 2;
    const realNameIndex = 3;

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var today = new Date()

    var sheet = spreadsheet.getSheetByName("USER")
    if (sheet) {
        var data = sheet.getDataRange().getValues()
        var dataLength = data.length

        var peopleAdded = 0

        while (results.next()) {
            let userId = results.getString(userIdIndex);
            let userName = results.getString(userNameIndex);
            let realName = results.getString(realNameIndex);
            let maybeFoundUser = userMap.get(userId)

            if (maybeFoundUser) {
                var row = data.find(y => y[0] == userId)
                if (row) {
                    rowNum = data.indexOf(row) + 1
                    if (userName != maybeFoundUser.username) {
                        sheet.getRange(rowNum, 2)
                            .setValues([[
                                userName
                            ]])
                    }
                    if (realName != maybeFoundUser.realName) {
                        sheet.getRange(rowNum, 3)
                            .setValues([[
                                realName
                            ]])
                    }
                }
            } else {
                peopleAdded += 1
                var rowNum = peopleAdded + dataLength
                sheet.getRange(rowNum, 1, 1, 5).setValues([[
                    userId, userName, realName, new Date(today.getFullYear(), today.getMonth(), today.getDay()), "YES"
                ]])
            }
        }
    }

    results.close();
    stmt.close();
}

function setRule(): void {
    const startingDatesColumnNum = 8

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    spreadsheet.getSheets().forEach(x => {
        // Add rule based formatting
        var range = x.getRange(3, startingDatesColumnNum, 1000, 29)
        // Rule that turns anything in the range with a value to green.
        var rule = SpreadsheetApp.newConditionalFormatRule()
            .whenCellNotEmpty()
            .setBackground("#90EE90")
            .setRanges([range])
            .build()

        // Rule that turns anything in the range with a value to green.
        var ruleTwo = SpreadsheetApp.newConditionalFormatRule()
            .whenCellEmpty()
            .setBackground("#ffcccb")
            .setRanges([range])
            .build()


        var rules: GoogleAppsScript.Spreadsheet.ConditionalFormatRule[] = []
        rules.push(rule);
        rules.push(ruleTwo)
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
        for (var i = 3; i >= 0; i--) {
            let currentDate = new Date();
            // if (currentDate.getDate() - i >= 0) {
            currentDate.setDate(currentDate.getDate() - i);
            var formattedColumnName = apostraphePrefix + Utilities.formatDate(currentDate, 'GMT', 'dd-MM')
            var col = regionData[0].indexOf(Utilities.formatDate(currentDate, 'GMT', 'dd-MM'));
            var anotherCol = regionData[0].indexOf(formattedColumnName);

            if (col == -1 && anotherCol == -1) {
                sheet.insertColumnBefore(startingDatesColumnNum);
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
                sheet.getRange(2, startingDatesColumnNum).setValues([["=COUNTIF(H3:H,\"?*\")"]])
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
    start.setDate(start.getDate() - 3);
    const dateFormatted = Utilities.formatDate(start, 'GMT', 'YYYY-MM-dd')

    const stmt = conn.createStatement();
    stmt.setMaxRows(500);
    // const results = stmt.executeQuery("SELECT * FROM bd_attendance WHERE date > '" + dateFormatted + "'");
    const results = stmt.executeQuery("SELECT * FROM bd_attendance WHERE date > '2022-07-30' AND date < '2022-08-12'");

    const bdDateIndex = 3;
    const userIndex = 1;
    const aoIndex = 2;
    const qUserIndex = 4;

    let bd_attendance: BD_ATTENDANCE_MAP = new Map<USER_ID, BD_ATTENDANCE[]>()

    while (results.next()) {
        let bdDate = new Date(results.getString(bdDateIndex));
        var formattedColumnName = Utilities.formatDate(bdDate, 'GMT', 'dd-MM')

        let user = results.getString(userIndex);
        let ao = results.getString(aoIndex);
        let qUserId = results.getString(qUserIndex)

        let res = bd_attendance.get(user)
        if (res) {
            res.push({
                date: formattedColumnName,
                user: user,
                ao: ao,
                qUser: qUserId
            })
            bd_attendance.set(user, res)
        } else {
            bd_attendance.set(user, [
                {
                    date: formattedColumnName,
                    user: user,
                    ao: ao,
                    qUser: qUserId
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

function retrieveUserMap(): USER_INFORMATION {
    var userMap = new Map<string, USER>()

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName("USER")
    if (sheet) {
        var values = sheet.getDataRange().getValues();

        return values.reduce((acc: Map<string, USER>, currentValue: any[], currentIndex): Map<string, USER> => {
            acc.set(currentValue[0], {
                id: currentValue[0],
                username: currentValue[1],
                realName: currentValue[2],
                startDate: currentValue[3] ? new Date(currentValue[3]) : null,
                rowIndex: currentIndex,
                include: currentValue.length > 4 ? currentValue[4] : false
            }
            )

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

function populateRegion(bdAttendance: BD_ATTENDANCE_MAP, userMap: USER_INFORMATION, aoMap): void {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName("REGION")
    var peopleAdded = 0
    if (sheet) {
        var data = sheet.getDataRange().getValues();
        var dataLength = data.length

        bdAttendance.forEach((x, key) => {
            var user = userMap.get(key)
            if (user?.include) {
                var row = data.find(y => {
                    return y[0] == key
                })

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
                            userMap.get(key) ? userMap.get(key)?.username : "",
                            "",
                            "=COUNTIF(INDIRECT(TEXT(CONCATENATE(ADDRESS(ROW(),8,3),\":\",ADDRESS(ROW(),63,3)),1)),\"Q?*\")",
                            "=COUNTIF(INDIRECT(TEXT(CONCATENATE(ADDRESS(ROW(),8,3),\":\",ADDRESS(ROW(),63,3)),1)),\"?*\")",
                            "=COUNTIF(INDIRECT(TEXT(CONCATENATE(ADDRESS(ROW(),8,3),\":\",ADDRESS(ROW(),21,3)),1)),\"?*\")",
                        ]])
                } else {
                    if (row[0] && (!row[1] || row[1] != userMap.get(key)?.username)) {
                        sheet?.getRange(rowNum, 2)
                            .setValues([[
                                userMap.get(key)?.username
                            ]])
                    }
                }

                x.forEach(bb => {
                    var col = data[0].indexOf(bb.date);

                    if (col != -1) {
                        if (aoMap.get(bb.ao)) {
                            if (bb.qUser == bb.user) {
                                sheet?.getRange(rowNum, col + 1)
                                    .setValues([[
                                        "Q-" + aoMap.get(bb.ao).shortcutName
                                    ]]).setHorizontalAlignment("center")
                            } else {
                                sheet?.getRange(rowNum, col + 1)
                                    .setValues([[
                                        aoMap.get(bb.ao).shortcutName
                                    ]]).setHorizontalAlignment("center")
                            }
                        }
                    }
                })
            }
        })
    } else {
        var ui = SpreadsheetApp.getUi();
        ui.alert("No region sheet found")
    }
}

function statFormulaFromAO(ao: AO, rowNum): Map<number, string> {
    var BD_PER_WEEK_TO_FORMULA = {
        1: {
            2: "=COUNTIF(INDIRECT(TEXT(CONCATENATE(ADDRESS(ROW(),8,3),\":\",ADDRESS(ROW(),9,3)),1)),\"?*\")",
            8: "=COUNTIF(INDIRECT(TEXT(CONCATENATE(ADDRESS(ROW(),8,3),\":\",ADDRESS(ROW(),14,3)),1)),\"?*\")",
            q: "=COUNTIF(INDIRECT(TEXT(CONCATENATE(ADDRESS(ROW(),8,3),\":\",ADDRESS(ROW(),14,3)),1)),\"Q\")"
        },
        2: {
            2: "=COUNTIF(INDIRECT(TEXT(CONCATENATE(ADDRESS(ROW(),8,3),\":\",ADDRESS(ROW(),11,3)),1)),\"?*\")",
            8: "=COUNTIF(INDIRECT(TEXT(CONCATENATE(ADDRESS(ROW(),8,3),\":\",ADDRESS(ROW(),21,3)),1)),\"?*\")",
            q: "=COUNTIF(INDIRECT(TEXT(CONCATENATE(ADDRESS(ROW(),8,3),\":\",ADDRESS(ROW(),21,3)),1)),\"Q\")",
        },
        3: {
            2: "=COUNTIF(INDIRECT(TEXT(CONCATENATE(ADDRESS(ROW(),8,3),\":\",ADDRESS(ROW(),13,3)),1)),\"?*\")",
            8: "=COUNTIF(INDIRECT(TEXT(CONCATENATE(ADDRESS(ROW(),8,3),\":\",ADDRESS(ROW(),31,3)),1)),\"?*\")",
            q: "=COUNTIF(INDIRECT(TEXT(CONCATENATE(ADDRESS(ROW(),8,3),\":\",ADDRESS(ROW(),31,3)),1)),\"Q\")",
        }
    }

    return BD_PER_WEEK_TO_FORMULA[ao.schedule.filter(x => x).length]
}

function populateAO(bdAttendance: BD_ATTENDANCE_MAP, userMap, ao: AO, sheet: GoogleAppsScript.Spreadsheet.Sheet): void {
    var peopleAdded = 0
    var data = sheet.getDataRange().getValues();
    var dataLength = data.length

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
                var aoFormulas = statFormulaFromAO(ao, rowNum)

                sheet.getRange(rowNum, 1, 1, 6)
                    .setValues([[
                        key,
                        userMap.get(key) ? userMap.get(key)?.username : "",
                        "",
                        aoFormulas ? aoFormulas["q"] : "",
                        aoFormulas ? aoFormulas[8] : "",
                        aoFormulas ? aoFormulas[2] : ""
                    ]])
            }
        }


        x.filter(xx => xx.ao == ao.channelId).forEach(bb => {
            var col = data[0].indexOf(bb.date);

            if (col != -1) {
                if (bb.qUser == bb.user) {
                    sheet.getRange(rowNum, col + 1)
                        .setValues([[
                            "Q"
                        ]]).setHorizontalAlignment("center")
                } else {
                    sheet.getRange(rowNum, col + 1)
                        .setValues([[
                            "X"
                        ]]).setHorizontalAlignment("center")
                }
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
        var sheet = spreadsheet.getSheetByName(val.friendlyName)

        if (sheet) {
            populateAO(bdAttendance, userMap, val, sheet)
        }
    })
}


