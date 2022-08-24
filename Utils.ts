// Takes in a date, and returns a DD-Month combo
// e.g. Date -> 22-Aug
function convertDateToReadable(date: Date): string {
    return date.toString()
}

function convertReadableToDate(str: string): Date {
    return new Date(str)
}