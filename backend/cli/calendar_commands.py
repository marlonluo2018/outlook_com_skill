"""Calendar listing, availability checking, and meeting response commands."""

from __future__ import annotations

import json
import sys
from datetime import datetime, time, timedelta

from backend.outlook_session.session_manager import OutlookSessionManager

from .common import _get_email_item, _get_smtp_address


STATUS_MAP = {
    '0': 'Free',
    '1': 'Tentative',
    '2': 'Busy',
    '3': 'OOF',
    '4': 'Working Elsewhere',
}

STATUS_DISPLAY = {
    'Free': 'Free',
    'Tentative': 'Tentative',
    'Busy': 'Busy',
    'OOF': 'Out of Office',
    'Working Elsewhere': 'Working Elsewhere',
}


def _parse_time(t_str: str) -> time:
    """Parse HH:MM string to time object."""
    t_str = t_str.strip()
    parts = t_str.split(':')
    return time(int(parts[0]), int(parts[1]))


def _parse_working_hours(wh_str: str) -> tuple[time, time] | None:
    """Parse working hours string like '09:00-18:00'. Returns (start, end) or None for all-day."""
    if not wh_str or wh_str.strip().lower() in ('all', '24h', '00:00-24:00', '00:00-23:59'):
        return None
    try:
        parts = wh_str.split('-')
        return _parse_time(parts[0]), _parse_time(parts[1])
    except Exception:
        return time(9, 0), time(18, 0)


def _get_date_list(args) -> list[datetime]:
    """Determine list of dates (normalized to 00:00:00) for query."""
    if getattr(args, 'date', None):
        start_date = datetime.strptime(args.date, '%Y-%m-%d').replace(hour=0, minute=0, second=0, microsecond=0)
        end_date = start_date
    elif getattr(args, 'start', None):
        start_date = datetime.strptime(args.start, '%Y-%m-%d').replace(hour=0, minute=0, second=0, microsecond=0)
        if getattr(args, 'end', None):
            end_date = datetime.strptime(args.end, '%Y-%m-%d').replace(hour=0, minute=0, second=0, microsecond=0)
        else:
            days = getattr(args, 'days', 1) or 1
            end_date = start_date + timedelta(days=days - 1)
    else:
        start_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        if getattr(args, 'end', None):
            end_date = datetime.strptime(args.end, '%Y-%m-%d').replace(hour=0, minute=0, second=0, microsecond=0)
        else:
            days = getattr(args, 'days', 1) or 1
            end_date = start_date + timedelta(days=days - 1)

    if end_date < start_date:
        end_date = start_date

    dates = []
    curr = start_date
    while curr <= end_date:
        dates.append(curr)
        curr += timedelta(days=1)
    return dates


def _get_recipient_info(recip) -> dict:
    """Extract full contact/Exchange info for a recipient."""
    info = {
        'name': getattr(recip, 'Name', '') or '',
        'email': '',
        'job_title': '',
        'department': '',
        'office': '',
        'city': '',
    }
    try:
        info['email'] = _get_smtp_address(recip)
    except Exception:
        info['email'] = getattr(recip, 'Address', '') or info['name']

    try:
        addr_entry = recip.AddressEntry
        if addr_entry:
            ex_user = addr_entry.GetExchangeUser()
            if ex_user:
                info['name'] = ex_user.Name or info['name']
                info['email'] = ex_user.PrimarySmtpAddress or info['email']
                info['job_title'] = ex_user.JobTitle or ''
                info['department'] = ex_user.Department or ''
                info['office'] = ex_user.OfficeLocation or ''
                info['city'] = ex_user.City or ''
    except Exception:
        pass
    return info


def _fetch_freebusy_slots(recip, date_list: list[datetime], interval: int) -> dict[str, list[dict]]:
    """Fetch and parse Free/Busy status for given recipient across date_list."""
    if not date_list:
        return {}
    base_start = date_list[0]
    fb_str = recip.FreeBusy(base_start, interval, False)
    slots_per_day = 1440 // interval

    day_slots = {}
    for day_idx, d in enumerate(date_list):
        day_key = d.strftime('%Y-%m-%d')
        slots = []
        for s in range(slots_per_day):
            char_idx = day_idx * slots_per_day + s
            code = fb_str[char_idx] if char_idx < len(fb_str) else '0'
            status = STATUS_MAP.get(code, 'Free')
            start_m = s * interval
            end_m = start_m + interval
            start_t = time(start_m // 60, start_m % 60)
            end_t = time(end_m // 60, end_m % 60) if end_m < 1440 else time(23, 59)
            slots.append({
                'slot_index': s,
                'start': start_t.strftime('%H:%M'),
                'end': end_t.strftime('%H:%M') if end_m < 1440 else '24:00',
                'start_time': start_t,
                'end_time': end_t,
                'code': code,
                'status': status,
            })
        day_slots[day_key] = slots
    return day_slots


def _fetch_own_calendar_appointments(session, date_list: list[datetime]) -> dict[str, list[dict]]:
    """Fetch appointments from user's own calendar folder."""
    ns = session.outlook_namespace
    calendar_folder = ns.GetDefaultFolder(9)  # olFolderCalendar
    items = calendar_folder.Items
    items.Sort("[Start]")
    items.IncludeRecurrences = True

    start_date = date_list[0]
    end_date = date_list[-1] + timedelta(days=1)
    start_str = start_date.strftime('%m/%d/%Y %H:%M %p')
    end_str = end_date.strftime('%m/%d/%Y %H:%M %p')
    restriction = f"[Start] >= '{start_str}' AND [Start] < '{end_str}'"
    filtered = items.Restrict(restriction)

    resp_map = {0: "None", 1: "Organized", 2: "Tentative", 3: "Accepted", 4: "Declined", 5: "Not Responded"}
    appts_by_day = {d.strftime('%Y-%m-%d'): [] for d in date_list}

    item = filtered.GetFirst()
    while item:
        try:
            start_val = item.Start.replace(tzinfo=None)
            end_val = item.End.replace(tzinfo=None)
            subject = item.Subject or "(No subject)"
            location = getattr(item, 'Location', '') or ''
            resp = getattr(item, 'ResponseStatus', 0)
            busy_status = getattr(item, 'BusyStatus', 2)
            meeting_status = getattr(item, 'MeetingStatus', 0)

            # Skip declined or cancelled meetings
            is_cancelled = (
                meeting_status == 5 or
                resp == 4 or
                subject.startswith("已取消:") or
                subject.startswith("Canceled:") or
                subject.startswith("Cancelled:")
            )

            day_key = start_val.strftime('%Y-%m-%d')
            if day_key in appts_by_day and not is_cancelled:
                appts_by_day[day_key].append({
                    'subject': subject,
                    'start': start_val.strftime('%H:%M'),
                    'end': end_val.strftime('%H:%M'),
                    'start_dt': start_val,
                    'end_dt': end_val,
                    'location': location,
                    'busy_status': busy_status,
                    'response': resp_map.get(resp, ""),
                })
        except Exception:
            pass
        item = filtered.GetNext()

    return appts_by_day


def _fetch_shared_calendar_appointments(ns, recip, date_list: list[datetime]) -> dict[str, list[dict]] | None:
    """Try to fetch appointments from shared calendar folder if accessible."""
    try:
        shared_folder = ns.GetSharedDefaultFolder(recip, 9)
        items = shared_folder.Items
        items.Sort("[Start]")
        items.IncludeRecurrences = True

        start_date = date_list[0]
        end_date = date_list[-1] + timedelta(days=1)
        start_str = start_date.strftime('%m/%d/%Y %H:%M %p')
        end_str = end_date.strftime('%m/%d/%Y %H:%M %p')
        restriction = f"[Start] >= '{start_str}' AND [Start] < '{end_str}'"
        filtered = items.Restrict(restriction)

        appts_by_day = {d.strftime('%Y-%m-%d'): [] for d in date_list}
        item = filtered.GetFirst()
        while item:
            try:
                start_val = item.Start.replace(tzinfo=None)
                end_val = item.End.replace(tzinfo=None)
                subject = item.Subject or "(Private / Busy)"
                location = getattr(item, 'Location', '') or ''
                day_key = start_val.strftime('%Y-%m-%d')
                if day_key in appts_by_day:
                    appts_by_day[day_key].append({
                        'subject': subject,
                        'start': start_val.strftime('%H:%M'),
                        'end': end_val.strftime('%H:%M'),
                        'start_dt': start_val,
                        'end_dt': end_val,
                        'location': location,
                    })
            except Exception:
                pass
            item = filtered.GetNext()
        return appts_by_day
    except Exception:
        return None


def _group_contiguous_blocks(slots: list[dict], work_window: tuple[time, time] | None = None) -> list[dict]:
    """Group contiguous slots into status blocks."""
    if not slots:
        return []

    # Filter to working hours if provided
    filtered_slots = []
    for s in slots:
        if work_window:
            w_start, w_end = work_window
            if s['end_time'] <= w_start or s['start_time'] >= w_end:
                continue
        filtered_slots.append(s)

    if not filtered_slots:
        return []

    blocks = []
    curr_status = filtered_slots[0]['status']
    curr_start = filtered_slots[0]['start']
    curr_end = filtered_slots[0]['end']
    slot_count = 1

    for s in filtered_slots[1:]:
        if s['status'] == curr_status:
            curr_end = s['end']
            slot_count += 1
        else:
            blocks.append({
                'start': curr_start,
                'end': curr_end,
                'status': curr_status,
                'duration_minutes': (len(blocks) + 1),  # placeholder updated below
                'slot_count': slot_count,
            })
            curr_status = s['status']
            curr_start = s['start']
            curr_end = s['end']
            slot_count = 1

    blocks.append({
        'start': curr_start,
        'end': curr_end,
        'status': curr_status,
        'slot_count': slot_count,
    })

    # Calculate exact duration
    for b in blocks:
        start_t = _parse_time(b['start'])
        end_t = _parse_time(b['end'] if b['end'] != '24:00' else '23:59')
        dur = (end_t.hour * 60 + end_t.minute) - (start_t.hour * 60 + start_t.minute)
        if b['end'] == '24:00':
            dur += 1
        b['duration_minutes'] = dur

    return blocks


def _compute_mutual_free_blocks(all_user_slots: list[list[dict]], interval: int, work_window: tuple[time, time] | None = None) -> list[dict]:
    """Find contiguous time blocks where ALL participants are Free."""
    if not all_user_slots:
        return []

    num_slots = len(all_user_slots[0])
    mutual_slots = []
    for s_idx in range(num_slots):
        s_proto = all_user_slots[0][s_idx]
        if work_window:
            w_start, w_end = work_window
            if s_proto['end_time'] <= w_start or s_proto['start_time'] >= w_end:
                continue

        all_free = all(slots[s_idx]['status'] == 'Free' for slots in all_user_slots)
        mutual_slots.append({
            'start': s_proto['start'],
            'end': s_proto['end'],
            'start_time': s_proto['start_time'],
            'end_time': s_proto['end_time'],
            'status': 'Free' if all_free else 'Busy',
        })

    blocks = _group_contiguous_blocks(mutual_slots, work_window)
    return [b for b in blocks if b['status'] == 'Free']


def cmd_check_availability(args) -> int:
    """Check availability / free-busy for one or more contacts."""
    try:
        user_inputs = []
        if getattr(args, 'users', None):
            if isinstance(args.users, list):
                for u in args.users:
                    user_inputs.extend([x.strip() for x in u.replace(';', ',').split(',') if x.strip()])
            else:
                user_inputs.extend([x.strip() for x in str(args.users).replace(';', ',').split(',') if x.strip()])

        user_flag = getattr(args, 'user_flag', None) or getattr(args, 'user', None) or getattr(args, 'contact', None)
        if user_flag:
            user_inputs.extend([x.strip() for x in str(user_flag).replace(';', ',').split(',') if x.strip()])

        # Support reading email list from a file (--file)
        file_path = getattr(args, 'file', None)
        if file_path:
            import os
            if os.path.exists(file_path):
                with open(file_path, 'r', encoding='utf-8') as f:
                    for line in f:
                        line = line.strip()
                        if line and not line.startswith('#'):
                            user_inputs.extend([x.strip() for x in line.replace(';', ',').split(',') if x.strip()])
            else:
                print(f"Error: File not found: {file_path}", file=sys.stderr)
                return 1

        if not user_inputs:
            print("Error: Please specify at least one user email or contact name.", file=sys.stderr)
            return 1

        interval = getattr(args, 'interval', 30) or 30
        working_hours_str = getattr(args, 'working_hours', '09:00-18:00') or '09:00-18:00'
        work_window = _parse_working_hours(working_hours_str)
        date_list = _get_date_list(args)
        compare = getattr(args, 'compare', False)
        common_only = getattr(args, 'common_only', False) or getattr(args, 'slots_only', False)
        min_duration = getattr(args, 'duration', None) or getattr(args, 'min_duration', None)
        if min_duration is not None:
            min_duration = int(min_duration)
        output_json = getattr(args, 'json', False)

        with OutlookSessionManager() as session:
            ns = session.outlook_namespace or (session.outlook.GetNamespace("MAPI") if session.outlook else None)
            if not ns:
                print("Error: Could not access Outlook MAPI namespace.", file=sys.stderr)
                return 1

            # Resolve recipients
            resolved_recips = []
            for u in user_inputs:
                recip = ns.CreateRecipient(u)
                recip.Resolve()
                if not recip.Resolved:
                    print(f"Warning: Could not resolve recipient '{u}'.", file=sys.stderr)
                    continue
                info = _get_recipient_info(recip)
                resolved_recips.append((recip, info))

            if not resolved_recips:
                print("Error: None of the specified recipients could be resolved in Outlook.", file=sys.stderr)
                return 1

            # Fetch own calendar if comparing
            own_info = None
            own_slots_by_day = None
            own_appts_by_day = None
            if compare or len(resolved_recips) > 1:
                own_recip = ns.CurrentUser
                if own_recip:
                    own_info = _get_recipient_info(own_recip)
                    own_slots_by_day = _fetch_freebusy_slots(own_recip, date_list, interval)
                    own_appts_by_day = _fetch_own_calendar_appointments(session, date_list)

            # Build data structure
            results = {
                'working_hours': working_hours_str,
                'interval_minutes': interval,
                'dates': [d.strftime('%Y-%m-%d') for d in date_list],
                'users': [],
            }

            for recip, info in resolved_recips:
                fb_by_day = _fetch_freebusy_slots(recip, date_list, interval)
                shared_appts = _fetch_shared_calendar_appointments(ns, recip, date_list)

                user_data = {
                    'contact': info,
                    'day_slots': fb_by_day,
                    'days': {},
                }

                for d in date_list:
                    d_str = d.strftime('%Y-%m-%d')
                    day_slots = fb_by_day.get(d_str, [])
                    blocks = _group_contiguous_blocks(day_slots, work_window)
                    free_blocks = [b for b in blocks if b['status'] == 'Free']
                    busy_blocks = [b for b in blocks if b['status'] != 'Free']

                    user_data['days'][d_str] = {
                        'day_of_week': d.strftime('%a'),
                        'blocks': blocks,
                        'free_blocks': free_blocks,
                        'busy_blocks': busy_blocks,
                        'appointments': shared_appts.get(d_str, []) if shared_appts else None,
                        'has_shared_folder': shared_appts is not None,
                    }

                results['users'].append(user_data)

            # Compute mutual free slots if requested
            if compare or len(results['users']) > 1:
                results['mutual_free_slots'] = {}
                for d in date_list:
                    d_str = d.strftime('%Y-%m-%d')
                    slot_lists_for_day = []
                    for u in results['users']:
                        slot_lists_for_day.append(u['day_slots'].get(d_str, []))
                    if compare and own_slots_by_day:
                        slot_lists_for_day.append(own_slots_by_day.get(d_str, []))

                    mutual_blocks = _compute_mutual_free_blocks(slot_lists_for_day, interval, work_window)
                    if min_duration:
                        mutual_blocks = [b for b in mutual_blocks if b['duration_minutes'] >= min_duration]
                    results['mutual_free_slots'][d_str] = mutual_blocks

            if output_json:
                print(json.dumps(results, indent=2, default=str))
                return 0

            # Render formatted human-readable output
            print("=" * 60)
            print(" Calendar Availability & Free/Busy Schedule")
            print("=" * 60)
            wh_label = working_hours_str if work_window else "Full Day (24h)"
            dur_label = f" | Min Duration: {min_duration}m" if min_duration else ""
            print(f"Working Hours: {wh_label} | Slot Resolution: {interval}m{dur_label}\n")

            if not common_only:
                for u_idx, u in enumerate(results['users'], start=1):
                    c = u['contact']
                    contact_line = f"👤 {c['name']} <{c['email']}>"
                    if c.get('job_title'):
                        contact_line += f" — {c['job_title']}"
                    if c.get('office') or c.get('city'):
                        loc = c.get('office') or c.get('city')
                        contact_line += f" ({loc})"
                    print(contact_line)
                    print("-" * 60)

                    for d in date_list:
                        d_str = d.strftime('%Y-%m-%d')
                        day_info = u['days'][d_str]
                        dow = day_info['day_of_week']
                        print(f"\n📅 {d_str} ({dow}):")

                        if day_info.get('has_shared_folder') and day_info.get('appointments'):
                            print("  Detailed Appointments (Shared Calendar):")
                            for appt in day_info['appointments']:
                                loc_str = f" @ {appt['location']}" if appt.get('location') else ""
                                print(f"    • {appt['start']}-{appt['end']}: {appt['subject']}{loc_str}")
                        elif day_info['busy_blocks']:
                            print("  Busy / Reserved Periods:")
                            for b in day_info['busy_blocks']:
                                status_label = STATUS_DISPLAY.get(b['status'], b['status'])
                                print(f"    • {b['start']} - {b['end']} ({b['duration_minutes']} min): {status_label}")
                        else:
                            print(f"  🟢 Open / Free all day during working hours ({wh_label})")

                        if day_info['free_blocks'] and day_info['busy_blocks']:
                            print("  Available Free Slots:")
                            for b in day_info['free_blocks']:
                                print(f"    • {b['start']} - {b['end']} ({b['duration_minutes']} min)")

                    print()

            # Display mutual free slots if compared
            if 'mutual_free_slots' in results:
                print("=" * 60)
                compared_parties = [u['contact']['name'] for u in results['users']]
                if compare and own_info:
                    compared_parties.append(f"{own_info['name']} (You)")
                print(f"🤝 Mutual Common Free Slots ({len(compared_parties)} participants: {', '.join(compared_parties)})")
                print("=" * 60)

                for d in date_list:
                    d_str = d.strftime('%Y-%m-%d')
                    dow = d.strftime('%a')
                    free_blocks = results['mutual_free_slots'].get(d_str, [])
                    print(f"\n📅 {d_str} ({dow}):")
                    if free_blocks:
                        for b in free_blocks:
                            print(f"  ✅ {b['start']} - {b['end']} ({b['duration_minutes']} mins free)")
                    else:
                        print("  ❌ No common free time slots found during working hours.")

            print()
        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1


def cmd_list_calendar(args):
    """List calendar appointments for a date range."""
    user = getattr(args, 'user', None) or getattr(args, 'contact', None)
    if user:
        return cmd_check_availability(args)

    try:
        with OutlookSessionManager() as session:
            ns = session.outlook_namespace or (session.outlook.GetNamespace("MAPI") if session.outlook else None)
            if not ns:
                print("Error: Could not access Outlook MAPI namespace.", file=sys.stderr)
                return 1
            calendar_folder = ns.GetDefaultFolder(9)  # olFolderCalendar
            items = calendar_folder.Items
            items.Sort("[Start]")
            items.IncludeRecurrences = True

            # Determine date range
            if getattr(args, 'start', None):
                start_date = datetime.strptime(args.start, '%Y-%m-%d')
            else:
                start_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

            if getattr(args, 'end', None):
                end_date = datetime.strptime(args.end, '%Y-%m-%d') + timedelta(days=1)
            else:
                days = getattr(args, 'days', 1) or 1
                end_date = start_date + timedelta(days=days)

            start_str = start_date.strftime('%m/%d/%Y %H:%M %p')
            end_str = end_date.strftime('%m/%d/%Y %H:%M %p')
            restriction = f"[Start] >= '{start_str}' AND [Start] < '{end_str}'"
            filtered = items.Restrict(restriction)

            resp_map = {0: "None", 1: "Organized", 2: "Tentative", 3: "Accepted", 4: "Declined", 5: "Not Responded"}
            current_day = None
            count = 0

            item = filtered.GetFirst()
            while item:
                try:
                    start_val = item.Start.replace(tzinfo=None)
                    end_val = item.End.replace(tzinfo=None)
                    subject = item.Subject or "(No subject)"
                    location = getattr(item, 'Location', '') or ''
                    organizer = getattr(item, 'Organizer', '') or ''
                    resp = getattr(item, 'ResponseStatus', 0)
                    resp_label = resp_map.get(resp, "")

                    day_key = start_val.strftime('%Y-%m-%d (%a)')
                    if day_key != current_day:
                        if current_day is not None:
                            print()
                        print(day_key)
                        current_day = day_key

                    entry_id = getattr(item, 'EntryID', '') or ''
                    time_range = f"{start_val.strftime('%H:%M')}-{end_val.strftime('%H:%M')}"
                    parts = [time_range, subject]
                    if location:
                        parts.append(location)
                    if organizer:
                        parts.append(f"Organizer: {organizer}")
                    if resp_label:
                        parts.append(resp_label)
                    print(f"  {' | '.join(parts)}")
                    if entry_id:
                        print(f"    ID: {entry_id}")
                    count += 1
                except Exception:
                    pass
                item = filtered.GetNext()

            if count == 0:
                print("No appointments found in this date range.")
            else:
                print(f"\n({count} items)")
        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1


def cmd_respond_meeting(args):
    """Respond to a meeting invite: accept, tentative, decline, or propose new time."""
    try:
        email_id = args.email_id
        action = args.action

        with OutlookSessionManager() as session:
            email_item = _get_email_item(session, email_id)

            # Verify it's a meeting item
            message_class = getattr(email_item, 'MessageClass', '') or ''
            if not message_class.startswith('IPM.Schedule.Meeting'):
                print("Error: This item is not a meeting invite.", file=sys.stderr)
                return 1

            # Map action to Outlook constants
            ACTION_MAP = {'accept': 3, 'tentative': 2, 'decline': 4, 'propose': 2}
            response_const = ACTION_MAP[action]

            # Show what we're about to do
            subject = getattr(email_item, 'Subject', '(No subject)')
            appt = None
            try:
                appt = email_item.GetAssociatedAppointment(False)
            except Exception:
                pass

            organizer = ''
            when_str = ''
            if appt:
                organizer = getattr(appt, 'Organizer', '') or ''
                start_val = getattr(appt, 'Start', None)
                end_val = getattr(appt, 'End', None)
                if start_val:
                    s = start_val.replace(tzinfo=None)
                    when_str = str(s)
                    if end_val:
                        e = end_val.replace(tzinfo=None)
                        when_str += f" ~ {e}"

            action_label = {'accept': 'Accept', 'tentative': 'Tentative', 'decline': 'Decline', 'propose': 'Propose New Time'}[action]
            print(f"Action: {action_label}")
            print(f"Meeting: {subject}")
            if organizer:
                print(f"Organizer: {organizer}")
            if when_str:
                print(f"When: {when_str}")

            if action == 'propose':
                if not args.start or not args.end:
                    print("Error: --start and --end are required for propose action.", file=sys.stderr)
                    return 1
                new_start = datetime.strptime(args.start, '%Y-%m-%d %H:%M')
                new_end = datetime.strptime(args.end, '%Y-%m-%d %H:%M')
                print(f"Proposed: {new_start} ~ {new_end}")

            # Execute the response via AppointmentItem (MeetingItem doesn't expose Respond)
            # GetAssociatedAppointment(True) adds to calendar, returns AppointmentItem
            # AppointmentItem.Respond(olResponse, fNoUI) returns a MeetingItem to .Send()
            appt_to_respond = email_item.GetAssociatedAppointment(True)
            resp = appt_to_respond.Respond(response_const, True)
            resp.Send()

            print(f"\nDone. Response sent: {action_label}")
            print(f"ID: {email_id}")
        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1


# --- invite cancel / calendar delete -----------------------------------------
# Sending an invitation belongs to teams_web_skill (the browser can mint a Teams
# join link, the COM path cannot). Cancelling stays here: it is an outgoing
# message with a review path already, and it must act on the item Outlook holds.

_OL_MEETING = 1
_OL_MEETING_CANCELED = 5


def _appointment_summary(item) -> dict:
    """What an operator needs to see before an irreversible calendar action."""

    def safe(attr, default=""):
        try:
            return getattr(item, attr, default)
        except Exception:
            return default

    start = safe("Start", None)
    end = safe("End", None)
    try:
        when = f"{start.replace(tzinfo=None)} ~ {end.replace(tzinfo=None)}"
    except Exception:
        when = ""
    recipients = []
    try:
        for i in range(1, min(item.Recipients.Count, 25) + 1):
            address = item.Recipients.Item(i).Address or ""
            if address:
                recipients.append(address)
    except Exception:
        pass
    return {
        "subject": safe("Subject", "(No subject)"),
        "when": when,
        "location": safe("Location", ""),
        "meeting_status": int(safe("MeetingStatus", 0) or 0),
        "organizer": safe("Organizer", ""),
        "recipients": recipients,
    }


def cmd_cancel_invite(args):
    """Cancel a meeting you organise, notifying every invitee.

    Manual invitations are created by teams_web_skill; cancelling stays in this
    skill because it is an outgoing message (same review gate as email) and it
    must act on the item Outlook already has, not in a browser session.
    """
    try:
        with OutlookSessionManager() as session:
            item = _get_email_item(session, args.email_id, getattr(args, "subject", None))
            summary = _appointment_summary(item)

            if summary["meeting_status"] != _OL_MEETING:
                print(
                    f"Error: not a live meeting (MeetingStatus={summary['meeting_status']}). "
                    "Only an organiser meeting can be cancelled; use delete-appointment "
                    "for a plain event.",
                    file=sys.stderr,
                )
                return 1

            print("Action: Cancel meeting (sends a cancellation to every invitee)")
            print(f"Meeting: {summary['subject']}")
            if summary["when"]:
                print(f"When: {summary['when']}")
            if summary["recipients"]:
                print(f"Invitees: {', '.join(summary['recipients'])}")

            if getattr(args, "dry_run", False):
                print("\nDry run: nothing was sent.")
                return 0

            # olMeetingCanceled + Send() is what makes Outlook notify invitees;
            # Delete() alone would drop the meeting without telling anyone.
            item.MeetingStatus = _OL_MEETING_CANCELED
            item.Send()
            print("\nCancellation sent.")
            return 0
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


def cmd_delete_appointment(args):
    """Remove a calendar item without notifying anybody."""
    try:
        with OutlookSessionManager() as session:
            item = _get_email_item(session, args.email_id, getattr(args, "subject", None))
            summary = _appointment_summary(item)

            print("Action: Delete calendar item (no notification is sent)")
            print(f"Item: {summary['subject']}")
            if summary["when"]:
                print(f"When: {summary['when']}")

            if summary["meeting_status"] == _OL_MEETING and not getattr(args, "force", False):
                print(
                    "\nError: this is a live meeting with invitees. Cancelling would notify "
                    "them, deleting would not. Use cancel-invite, or pass --force.",
                    file=sys.stderr,
                )
                return 1

            if getattr(args, "dry_run", False):
                print("\nDry run: nothing was deleted.")
                return 0

            item.Delete()
            print("\nDeleted.")
            return 0
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
