type BD_ATTENDANCE = {
    date: String,
    user: String,
    ao: String
}

// Have a "refresh user list dropdown menu"
function retrieveAllUsers() {
    let users = []  // Select * FROM f3stl.f3city.users;

}

function populateAos() {

}

function populateBB(): Array<BD_ATTENDANCE> {
    var passwordMaybe = PropertiesService.getScriptProperties().getProperty('pw');
    var userMaybe = PropertiesService.getScriptProperties().getProperty('user');

    if (!userMaybe || !passwordMaybe) {
        userMaybe = "GoingToFail"
        passwordMaybe = "GoingToFail"
    }

    var conn = Jdbc.getConnection('jdbc:mysql://f3stlouis.cac36jsyb5ss.us-east-2.rds.amazonaws.com:3306/f3stlcity', userMaybe, passwordMaybe);
    const start = new Date();
    const stmt = conn.createStatement();
    stmt.setMaxRows(10);
    const results = stmt.executeQuery('SELECT * FROM bd_attendance');

    const bdDateIndex = 3;
    const userIndex = 1;
    const aoIndex = 2;

    let bd_attendance: Array<BD_ATTENDANCE> = []

    while (results.next()) {
        let bdDate = results.getString(bdDateIndex);
        let user = results.getString(userIndex);
        let ao = results.getString(aoIndex);


        bd_attendance.push({
            date: bdDate,
            user: user,
            ao: ao
        })

    }

    results.close();
    stmt.close();

    return bd_attendance
}


function runItAll(): void {
    const user_map = {}
    const ao_map = {}
    const bd_attendance = populateBB()
    Logger.log(JSON.stringify(bd_attendance))
}