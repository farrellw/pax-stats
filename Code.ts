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
        {
            name: "Step 3: Calculate Stats",
            functionName: "populateStats"
        },
        {
            name: "Step 4: Sort",
            functionName: "sortPax"
        },
    ];
    spreadsheet.addMenu("F3", menus);
};

function sortPax(): void {

}

function populateStats(): void {
    const startingDatesColumnNum = 7
    const statsColumns = {
        total: 3,
        overTwoWeeks: 4,
        overMonth: 5,
    }

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName("REGION")
    var data = spreadsheet.getDataRange().getValues();
    if (sheet) {
        data.forEach((row, index) => {
            if (index != 0) {
                var beatdownSlice = row.slice(startingDatesColumnNum - 1)
                const stats = calculateStats(beatdownSlice)
                sheet?.getRange(index + 1, statsColumns.total, 1, 3).setValues([
                    [stats.perWeekTotal, stats.perWeekOverLastMonth, stats.perWeekOverLastTwoWeeks]
                ])
            }
        })

        // Find total number of rows
        var totalRows = data.length
        var totalColumns = data[0].length

        // Format the document

        // Add rule based formatting
        var range = sheet.getRange(2, startingDatesColumnNum, totalRows, totalColumns - startingDatesColumnNum)

        // Rule that turns anything in the range with a value to green.
        var rule = SpreadsheetApp.newConditionalFormatRule()
            .whenCellNotEmpty()
            .setBackground("#90EE90")
            .setRanges([range])
            .build()

        var rules: GoogleAppsScript.Spreadsheet.ConditionalFormatRule[] = []
        rules.push(rule);
        sheet.setConditionalFormatRules(rules);
        // #FFD580

        // FORMAT
    } else {
        var ui = SpreadsheetApp.getUi();
        ui.alert("No region sheet found")
    }
}

function populateDates(): void {
    const apostraphePrefix = "'"
    const startingDatesColumnNum = 7

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName("REGION")
    var data = spreadsheet.getDataRange().getValues();

    if (sheet) {

        for (var i = 7; i >= 0; i--) {
            let currentDate = new Date();
            // if (currentDate.getDate() - i >= 0) {
            currentDate.setDate(currentDate.getDate() - i);
            var formattedColumnName = apostraphePrefix + Utilities.formatDate(currentDate, 'GMT', 'dd-MM')
            var col = data[0].indexOf(Utilities.formatDate(currentDate, 'GMT', 'dd-MM'));

            if (col == -1) {
                spreadsheet.insertColumnAfter(startingDatesColumnNum - 1);
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
            }
            // }
        }
    } else {
        var ui = SpreadsheetApp.getUi();
        ui.alert("No region sheet found")
    }
}

function retrieveBB(): Array<BD_ATTENDANCE> {
    var passwordMaybe = PropertiesService.getScriptProperties().getProperty('pw');
    var userMaybe = PropertiesService.getScriptProperties().getProperty('user');

    if (!userMaybe || !passwordMaybe) {
        userMaybe = "GoingToFail"
        passwordMaybe = "GoingToFail"
    }

    var conn = Jdbc.getConnection('jdbc:mysql://f3stlouis.cac36jsyb5ss.us-east-2.rds.amazonaws.com:3306/f3stlcity', userMaybe, passwordMaybe);
    const start = new Date();
    start.setDate(start.getDate() - 7);
    const dateFormatted = Utilities.formatDate(start, 'GMT', 'YYYY-MM-dd')

    const stmt = conn.createStatement();
    stmt.setMaxRows(400);
    const results = stmt.executeQuery("SELECT * FROM bd_attendance WHERE date > '" + dateFormatted + "'");

    const bdDateIndex = 3;
    const userIndex = 1;
    const aoIndex = 2;

    let bd_attendance: Array<BD_ATTENDANCE> = []

    while (results.next()) {
        let bdDate = new Date(results.getString(bdDateIndex));
        var formattedColumnName = Utilities.formatDate(bdDate, 'GMT', 'dd-MM')

        let user = results.getString(userIndex);
        let ao = results.getString(aoIndex);

        bd_attendance.push({
            date: formattedColumnName,
            user: user,
            ao: ao
        })
    }

    results.close();
    stmt.close();

    return bd_attendance
}

function retrieveAOMap(): AO_INFORMATION {
    var aoMap = new Map<string, string>()
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    var sheet = spreadsheet.getSheetByName("AOs")
    if (sheet) {
        var values = sheet.getDataRange().getValues();

        return values.reduce((acc: Map<string, string>, currentValue: any[]): Map<string, string> => {
            if (currentValue[3] == "YES") {
                acc.set(currentValue[0], currentValue[2] ? currentValue[2] : currentValue[1])
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

function populateRegion(bdAttendance, userMap, aoMap): void {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName("REGION")
    var peopleAdded = 0
    var data = spreadsheet.getDataRange().getValues();
    if (sheet) {

        bdAttendance.forEach((x) => {
            var row = data.find(y => y[0] == x.user)
            var rowNum;
            if (!row) {
                var addToLength = 1 + peopleAdded
                peopleAdded += 1
                rowNum = data.length + 1
            } else {
                rowNum = data.indexOf(row) + 1
            }

            if (!row) {
                sheet?.getRange(rowNum, 1, 1, 2)
                    .setValues([[
                        x.user,
                        userMap.get(x.user),
                    ]])
            }

            var col = data[0].indexOf(x.date);

            if (col != -1) {
                sheet?.getRange(rowNum, col + 1)
                    .setValues([[
                        aoMap.get(x.ao)
                    ]])
            }

        })
    } else {
        var ui = SpreadsheetApp.getUi();
        ui.alert("No region sheet found")
    }
}

type Stats = {
    perWeekTotal: number,
    perWeekOverLastTwoWeeks: number,
    perWeekOverLastMonth: number
}

function calculateStats(beatdownSlice): Stats {
    return {
        perWeekTotal: calculateStat(beatdownSlice),
        perWeekOverLastMonth: calculateStat(beatdownSlice.slice(0, 31)),
        perWeekOverLastTwoWeeks: calculateStat(beatdownSlice.slice(0, 14)),
    }
}

function calculateStat(beatdownSlice): number {
    var numberOfOpportunities = beatdownSlice.length

    var numberOfPosts = beatdownSlice.filter(x => x != "").length

    var percentageOfPosts = numberOfPosts / numberOfOpportunities
    return percentageOfPosts * 7
}

function populateAO(bdAttendance, userMap, aoMap, aoID): void {

}

function populateAttendance(): void {
    const userMap = retrieveUserMap()
    const aoMap = retrieveAOMap()
    const bdAttendance = retrieveBB();

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    populateRegion(bdAttendance, userMap, aoMap)

    aoMap.forEach((val, key) => {
        var sheet = spreadsheet.getSheetByName(key)

        if (sheet) {
            populateAO(bdAttendance, userMap, aoMap, key)
        }
    })
}


