type BD_ATTENDANCE_MAP = Map<USER_ID, BD_ATTENDANCE[]>

type BD_ATTENDANCE = {
    date: string,
    user: string,
    ao: string,
}

type USER_ID = string
type USER = string

type USER_INFORMATION = Map<USER_ID, USER>
type AO_INFORMATION = Map<string, AO>

type AO = {
    channelId: string
    channel: string,
    shortcutName: string,
    friendlyName: string,
    schedule: boolean[]
}

