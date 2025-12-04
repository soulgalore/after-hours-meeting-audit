import fs from 'node:fs/promises';
import ical from 'node-ical';
import {DateTime} from 'luxon';

// TODO: Adjust to your timezone
const timeZone = 'Europe/Stockholm';

// TODO: Adjust to your working hours
const workStartHour = 8;
const workEndHour = 17;

// TODO: Adjust to yoour calendar file
const icsPath = new URL('./your.ics', import.meta.url);

// TODO: change this to your real Google Calendar email address
const userEmail = 'name@example.org';

const getDurations = (startDate, endDate) => {
  const localStart = DateTime.fromJSDate(startDate).setZone(timeZone);
  const localEnd = DateTime.fromJSDate(endDate).setZone(timeZone);

  if (!localStart.isValid || !localEnd.isValid) {
    return {totalHours: 0, outsideHours: 0};
  }

  const totalHours = localEnd.diff(localStart, 'hours').hours;

  // Ignore weird / zero / very long events (likely all-day / OOO)
  if (totalHours <= 0 || totalHours >= 8) {
    return {totalHours: 0, outsideHours: 0};
  }

  const workStart = localStart.set({
    hour: workStartHour,
    minute: 0,
    second: 0,
    millisecond: 0,
  });

  const workEnd = localStart.set({
    hour: workEndHour,
    minute: 0,
    second: 0,
    millisecond: 0,
  });

  const insideStart = localStart < workStart ? workStart : localStart;
  const insideEnd = localEnd > workEnd ? workEnd : localEnd;

  let insideHours = 0;
  if (insideEnd > insideStart) {
    insideHours = insideEnd.diff(insideStart, 'hours').hours;
  }

  const outsideHours = totalHours - insideHours;

  return {totalHours, outsideHours};
};

const round = (value) => Number.parseFloat(value.toFixed(2));

const percent = (part, whole) => {
  if (whole === 0) {
    return 0;
  }

  return round((part / whole) * 100);
};

// Require that you are an attendee and you did not DECLINE.
const didUserParticipate = (event) => {
  const attendees = Array.isArray(event.attendee)
    ? event.attendee
    : event.attendee
      ? [event.attendee]
      : [];

  if (attendees.length === 0) {
    // No attendees → likely a personal event, not a meeting you were invited to.
    return false;
  }

  const userEmailLower = userEmail.toLowerCase();

  for (const attendee of attendees) {
    const value = String(attendee.val ?? '').toLowerCase();
    const params = attendee.params ?? {};
    const partStat = String(params.PARTSTAT ?? '').toUpperCase();

    if (!value.includes('@')) {
      continue;
    }

    if (!value.includes(userEmailLower)) {
      continue;
    }

    // If it's explicitly DECLINED by you, skip.
    if (partStat === 'DECLINED') {
      return false;
    }

    // Anything else (ACCEPTED, TENTATIVE, NEEDS-ACTION, missing) counts.
    return true;
  }

  return false;
};

const main = async () => {
  const icsText = await fs.readFile(icsPath, 'utf8');
  const parsed = ical.sync.parseICS(icsText);

  const now = DateTime.now().setZone(timeZone);

  let totalMeetings = 0;
  let outsideMeetings = 0;
  let totalHours = 0;
  let outsideHours = 0;
  let firstEventStart;

  for (const event of Object.values(parsed)) {
    if (!event || event.type !== 'VEVENT' || !event.start || !event.end) {
      continue;
    }

    const eventStart = DateTime.fromJSDate(event.start).setZone(timeZone);

    // Ignore future events
    if (eventStart > now) {
      continue;
    }

    // Only meetings where you are an attendee and not DECLINED
    if (!didUserParticipate(event)) {
      continue;
    }

    if (!firstEventStart || eventStart < firstEventStart) {
      firstEventStart = eventStart;
    }

    const {totalHours: eventTotalHours, outsideHours: eventOutsideHours} =
      getDurations(event.start, event.end);

    if (eventTotalHours === 0) {
      continue;
    }

    totalMeetings += 1;
    totalHours += eventTotalHours;

    if (eventOutsideHours > 0) {
      outsideMeetings += 1;
      outsideHours += eventOutsideHours;
    }
  }

  if (!firstEventStart) {
    // eslint-disable-next-line no-console
    console.log(
      'No past meetings found where you are an attendee (and not declined).',
    );
    return;
  }

  const spanYears = round(now.diff(firstEventStart, 'years').years);
  const spanWeeksRaw = now.diff(firstEventStart, 'weeks').weeks;
  const spanMonthsRaw = now.diff(firstEventStart, 'months').months;

  const spanWeeks = spanWeeksRaw <= 0 ? 1 : spanWeeksRaw;
  const spanMonths = spanMonthsRaw <= 0 ? 1 : spanMonthsRaw;

  const meetingsOutsidePercent = percent(outsideMeetings, totalMeetings);
  const hoursOutsidePercent = percent(outsideHours, totalHours);

  const avgMeetingsPerWeek = round(totalMeetings / spanWeeks);
  const avgHoursPerWeek = round(totalHours / spanWeeks);
  const avgMeetingsPerMonth = round(totalMeetings / spanMonths);
  const avgHoursPerMonth = round(totalHours / spanMonths);

  const avgOutsideMeetingsPerWeek = round(outsideMeetings / spanWeeks);
  const avgOutsideHoursPerWeek = round(outsideHours / spanWeeks);
  const avgOutsideMeetingsPerMonth = round(outsideMeetings / spanMonths);
  const avgOutsideHoursPerMonth = round(outsideHours / spanMonths);

  // Summary
  // eslint-disable-next-line no-console
  console.log(
    `Range: ${firstEventStart.toISODate()} → ${now.toISODate()} (${spanYears} years)`,
  );
  // eslint-disable-next-line no-console
  console.log(
    `Total meetings (you listed as attendee, not declined): ${totalMeetings} (for ${userEmail})`,
  );
  // eslint-disable-next-line no-console
  console.log(
    `Meetings outside ${workStartHour}-${workEndHour}: ${outsideMeetings} (${meetingsOutsidePercent}%)`,
  );
  // eslint-disable-next-line no-console
  console.log(`Total hours in meetings: ${round(totalHours)}h`);
  // eslint-disable-next-line no-console
  console.log(
    `Hours outside ${workStartHour}-${workEndHour}: ${round(
      outsideHours,
    )}h (${hoursOutsidePercent}%)`,
  );

  // Averages (all meetings)
  // eslint-disable-next-line no-console
  console.log('\nAverages (all meetings where you participated):');
  // eslint-disable-next-line no-console
  console.log(
    `Per week: ${avgMeetingsPerWeek} meetings, ${avgHoursPerWeek}h in meetings`,
  );
  // eslint-disable-next-line no-console
  console.log(
    `Per month: ${avgMeetingsPerMonth} meetings, ${avgHoursPerMonth}h in meetings`,
  );

  // Averages (outside working hours)
  // eslint-disable-next-line no-console
  console.log('\nAverages (outside working hours only):');
  // eslint-disable-next-line no-console
  console.log(
    `Per week: ${avgOutsideMeetingsPerWeek} meetings, ${avgOutsideHoursPerWeek}h`,
  );
  // eslint-disable-next-line no-console
  console.log(
    `Per month: ${avgOutsideMeetingsPerMonth} meetings, ${avgOutsideHoursPerMonth}h`,
  );
};

await main();
