import ICAL from 'ical.js';

// Converts an iCalendar (.ics) event into the same field names that MsgReader produces for
// a .msg appointment, so templates work the same for both.
export function icsToTemplateData(icsText: string): Record<string, unknown> {
	let calendar: ICAL.Component;
	try {
		calendar = new ICAL.Component(ICAL.parse(icsText));
	} catch (ee: unknown) {
		throw new TypeError('Outlook Meeting Notes cannot process the file. It is not a valid ics file. '
			+ (ee instanceof Error ? ee.message : ''));
	}

	// Times are usually given in a named time zone defined in the file itself
	for (const vtimezone of calendar.getAllSubcomponents('vtimezone')) {
		ICAL.TimezoneService.register(vtimezone);
	}

	// A recurring meeting can also include changed occurrences, which have a RECURRENCE-ID
	const vevent = calendar.getAllSubcomponents('vevent').find(e => !e.hasProperty('recurrence-id'))
		?? calendar.getFirstSubcomponent('vevent');
	if (vevent == null) {
		throw new TypeError('Outlook Meeting Notes cannot process the file. '
			+ 'It is a valid ics file but does not contain a meeting.');
	}
	const event = new ICAL.Event(vevent);

	const organizer = vevent.getFirstProperty('organizer');
	return {
		subject: event.summary ?? '',
		apptStartWhole: event.startDate?.toJSDate().toISOString() ?? '',
		apptEndWhole: event.endDate?.toJSDate().toISOString() ?? '',
		apptLocation: event.location ?? '',
		body: event.description ?? '',
		senderName: organizer ? personName(organizer) : '',
		senderEmail: organizer ? personEmail(organizer) : '',
		recipients: event.attendees.map(attendee => ({
			name: personName(attendee),
			email: personEmail(attendee),
		})),
	};
}

function personEmail(property: ICAL.Property): string {
	return String(property.getFirstValue() ?? '').replace(/^mailto:/i, '');
}

function personName(property: ICAL.Property): string {
	const cn = property.getParameter('cn');
	return (Array.isArray(cn) ? cn[0] : cn) || personEmail(property);
}
