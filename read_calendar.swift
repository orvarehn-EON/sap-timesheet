// Reads calendar events from macOS Calendar via EventKit (fast, unlike
// Calendar.app's AppleScript bridge which is unusably slow with Exchange
// calendars because its "whose" filter must expand/scan the entire event
// history regardless of date range).
//
// Usage: swift read_calendar.swift <daysOffsetStart> <daysOffsetEnd> <calendarNamesCSV> <myEmail>
// Offsets are integer days relative to today (may be negative).
//
// Outputs one line per event, pipe-delimited:
//   startYear|startMonth|startDay|startHour|startMinute|endYear|endMonth|endDay|endHour|endMinute|allDay|myStatus|summary|location
//
// Note: EKParticipant.isCurrentUser is unreliable (always false) for
// Exchange/EWS-backed accounts, so "my" RSVP status is determined by
// matching the attendee's email against myEmail instead.

import EventKit
import Foundation

let args = CommandLine.arguments
guard args.count >= 5,
      let offsetStart = Int(args[1]),
      let offsetEnd = Int(args[2]) else {
    FileHandle.standardError.write("Usage: swift read_calendar.swift <daysOffsetStart> <daysOffsetEnd> <calendarNamesCSV> <myEmail>\n".data(using: .utf8)!)
    exit(1)
}
let calNamesCSV = args[3]
let calNames = Set(calNamesCSV.split(separator: ",").map { $0.trimmingCharacters(in: .whitespaces) })
let myEmail = args[4].lowercased()

let store = EKEventStore()
let sema = DispatchSemaphore(value: 0)
var granted = false

if #available(macOS 14.0, *) {
    store.requestFullAccessToEvents { ok, _ in
        granted = ok
        sema.signal()
    }
} else {
    store.requestAccess(to: .event) { ok, _ in
        granted = ok
        sema.signal()
    }
}
sema.wait()

guard granted else {
    FileHandle.standardError.write("Calendar access not granted.\n".data(using: .utf8)!)
    exit(1)
}

var cal = Calendar.current
let now = Date()
let startOfToday = cal.startOfDay(for: now)
guard let rangeStart = cal.date(byAdding: .day, value: offsetStart, to: startOfToday),
      let rangeEndDay = cal.date(byAdding: .day, value: offsetEnd, to: startOfToday),
      let rangeEnd = cal.date(byAdding: .second, value: 86399, to: rangeEndDay) else {
    FileHandle.standardError.write("Failed to compute date range.\n".data(using: .utf8)!)
    exit(1)
}

let allCalendars = store.calendars(for: .event)
let matchingCalendars = allCalendars.filter { calNames.contains($0.title) }
if matchingCalendars.isEmpty {
    FileHandle.standardError.write("No calendars matched: \(calNamesCSV). Available: \(allCalendars.map { $0.title }.joined(separator: ", "))\n".data(using: .utf8)!)
    exit(1)
}

let predicate = store.predicateForEvents(withStart: rangeStart, end: rangeEnd, calendars: matchingCalendars)
let events = store.events(matching: predicate)

func sanitize(_ s: String) -> String {
    return s.replacingOccurrences(of: "|", with: " ")
        .replacingOccurrences(of: "\n", with: " ")
        .replacingOccurrences(of: "\r", with: " ")
}

func statusString(_ s: EKParticipantStatus) -> String {
    switch s {
    case .accepted: return "accepted"
    case .declined: return "declined"
    case .tentative: return "tentative"
    case .pending: return "pending"
    case .delegated: return "delegated"
    case .completed: return "completed"
    case .inProcess: return "inProcess"
    default: return "unknown"
    }
}

func participantEmail(_ p: EKParticipant) -> String? {
    let urlString = p.url.absoluteString
    guard urlString.lowercased().hasPrefix("mailto:") else { return nil }
    return String(urlString.dropFirst("mailto:".count)).lowercased()
}

var output = ""
for evt in events {
    let comps = cal.dateComponents([.year, .month, .day, .hour, .minute], from: evt.startDate)
    let compsEnd = cal.dateComponents([.year, .month, .day, .hour, .minute], from: evt.endDate)

    var myStatus = "accepted"
    if let attendees = evt.attendees {
        if let me = attendees.first(where: { participantEmail($0) == myEmail }) {
            myStatus = statusString(me.participantStatus)
        }
    }

    let summary = sanitize(evt.title ?? "")
    let location = sanitize(evt.location ?? "")
    let line = "\(comps.year ?? 0)|\(comps.month ?? 0)|\(comps.day ?? 0)|\(comps.hour ?? 0)|\(comps.minute ?? 0)|\(compsEnd.year ?? 0)|\(compsEnd.month ?? 0)|\(compsEnd.day ?? 0)|\(compsEnd.hour ?? 0)|\(compsEnd.minute ?? 0)|\(evt.isAllDay)|\(myStatus)|\(summary)|\(location)\n"
    output += line
}

print(output, terminator: "")
