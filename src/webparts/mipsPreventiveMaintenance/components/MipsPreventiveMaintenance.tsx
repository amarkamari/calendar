import * as React from 'react';
import styles from './MipsPreventiveMaintenance.module.scss';
import { IMipsPreventiveMaintenanceProps } from './IMipsPreventiveMaintenanceProps';
import { SPHttpClient, ISPHttpClientOptions, SPHttpClientResponse } from '@microsoft/sp-http';

interface IPMItem {
  Id: number;
  Service: string;
  BU: string;
  Location: string;
  LastMaintenance: string;
  NextMaintenance: string;
  Frequency: string;
  Time: string;
  Duration: string;
  AffectedSystems: string;
}

interface IMipsState {
  items: IPMRow[];
  loading: boolean;
  error: string;
  threshold: number;
  search: string;
  bu: string;
  filter: string;
  expandedId: number | null;
  editing: IPMRow | null;
  editDate: string;
  activeTab: 'log' | 'calendar';
  calYear: number;
  calMonth: number;
  calSelectedDate: string | null;
  calBu: string;
  showCompleted: boolean;
  showProjected: boolean;
}
interface ICalEvent {
  id: number;
  service: string;
  bu: string;
  location: string;
  time: string;
  duration: string;
  frequency: string;
  systems: string[];
  date: Date;
  status: 'ok' | 'soon' | 'overdue' | 'completed' | 'projected';
  kind: 'completed' | 'scheduled' | 'projected';
  daysUntil?: number;
}
interface IPMRow extends IPMItem {
  lastDate: Date;
  nextDate: Date;
  daysUntil: number;
  status: 'ok' | 'soon' | 'overdue';
}



export default class MipsPreventiveMaintenance
  extends React.Component<IMipsPreventiveMaintenanceProps, IMipsState> {

  private fieldMap: { [key: string]: string } = {
    Service: 'Title',
    BU: 'field_1',
    Location: 'field_2',
    LastMaintenance: 'field_3',
    NextMaintenance: 'field_4',
    Frequency: 'field_5',
    Time: 'field_6',
    Duration: 'field_7',
    AffectedSystems: 'field_8'
  };

  private refreshTimer: number | undefined;

 public constructor(props: IMipsPreventiveMaintenanceProps) {
    super(props);

    const today = new Date();

    this.state = {
      items: [],
      loading: true,
      error: '',
      threshold: 14,
      search: '',
      bu: 'all',
      filter: 'all',
      expandedId: null,
      editing: null,
      editDate: '',
      activeTab: 'log',
      calYear: today.getFullYear(),
      calMonth: today.getMonth(),
      calSelectedDate: null,
      calBu: 'all',
      showCompleted: true,
      showProjected: false
    };
  }

  public componentDidMount(): void {
    this.loadData();

    this.refreshTimer = window.setInterval(
      () => this.loadData(),
      30 * 60 * 1000
    );
  }

  public componentWillUnmount(): void {
    if (this.refreshTimer) {
      window.clearInterval(this.refreshTimer);
    }
  }

  private parseDate(value: string): Date {
    if (!value) {
      return new Date(NaN);
    }

    const d = new Date(value);

    return new Date(
      d.getUTCFullYear(),
      d.getUTCMonth(),
      d.getUTCDate()
    );
  }

private isoDate(date: Date): string {
    const y = date.getFullYear();
    const mNum = date.getMonth() + 1;
    const dNum = date.getDate();
    const m = mNum < 10 ? '0' + mNum : String(mNum);
    const d = dNum < 10 ? '0' + dNum : String(dNum);
    return `${y}-${m}-${d}`;
}

private addFrequency(date: Date, frequency: string): Date {

  const next = new Date(date.getTime());

  switch ((frequency || '').trim().toLowerCase()) {

    case 'monthly':
      next.setMonth(next.getMonth() + 1);
      break;

    case 'quarterly':
      next.setMonth(next.getMonth() + 3);
      break;

    case 'half yearly':
      next.setMonth(next.getMonth() + 6);
      break;

    case 'yearly':
      next.setFullYear(next.getFullYear() + 1);
      break;

    default:
      next.setMonth(next.getMonth() + 1);
      break;
  }

  return next;
}

private updateMaintenanceDates(
  itemId: number,
  lastMaintenance: Date,
  nextMaintenance: Date
): Promise<SPHttpClientResponse> {

  const url =
    `${this.props.siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(
      this.props.listName
    )}')/items(${itemId})`;

  const body: any = {};

  body[this.fieldMap.LastMaintenance] =
    `${this.isoDate(lastMaintenance)}T00:00:00Z`;

  body[this.fieldMap.NextMaintenance] =
    `${this.isoDate(nextMaintenance)}T00:00:00Z`;

  const options: ISPHttpClientOptions = {
    headers: {
      Accept: 'application/json;odata=nometadata',
      'Content-Type': 'application/json;odata=nometadata',
      'IF-MATCH': '*',
      'X-HTTP-Method': 'MERGE'
    },
    body: JSON.stringify(body)
  };

  return this.props.context.spHttpClient.post(
    url,
    SPHttpClient.configurations.v1,
    options
  );
}
private createHistoryRecord(
  item: IPMItem
): Promise<SPHttpClientResponse> {

  const url =
    `${this.props.siteUrl}/_api/web/lists/getbytitle('PM_History')/items`;

const body: any = {
  Title: item.Service,
  field_1: item.BU,
  field_2: item.Location,
  field_3: item.LastMaintenance,
  field_4: item.NextMaintenance,
  field_5: item.Frequency,
  field_6: item.Time,
  field_7: item.Duration,
  field_8: item.AffectedSystems
};

  const options: ISPHttpClientOptions = {
    headers: {
      Accept: 'application/json;odata=nometadata',
      'Content-Type': 'application/json;odata=nometadata'
    },
    body: JSON.stringify(body)
  };

  return this.props.context.spHttpClient.post(
    url,
    SPHttpClient.configurations.v1,
    options
  );
}
  private formatDate(value: Date): string {
    if (isNaN(value.getTime())) {
      return '';
    }

    return value.toLocaleDateString('en-GB', {
      day: '2-digit',
      month: 'short',
      year: '2-digit'
    });
  }

 private loadData = (): void => {

  this.setState({
    loading: true,
    error: ''
  });

  const listName = this.props.listName;

  const url =
    `${this.props.siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/items?$top=500`;

  this.contextHttpGet(url)

    .then(async (rawItems: any[]) => {

      const today = new Date();
      today.setHours(0, 0, 0, 0);

      const items: IPMItem[] = rawItems.map(raw => ({
        Id: raw.Id,
        Service: raw[this.fieldMap.Service],
        BU: raw[this.fieldMap.BU],
        Location: raw[this.fieldMap.Location],
        LastMaintenance: raw[this.fieldMap.LastMaintenance],
        NextMaintenance: raw[this.fieldMap.NextMaintenance],
        Frequency: raw[this.fieldMap.Frequency],
        Time: raw[this.fieldMap.Time],
        Duration: raw[this.fieldMap.Duration],
        AffectedSystems: raw[this.fieldMap.AffectedSystems]
      }));

      const updates: Promise<any>[] = [];

      items.forEach(item => {

        let lastDate = this.parseDate(item.LastMaintenance);
        let nextDate = this.parseDate(item.NextMaintenance);

        let changed = false;

        while (
          !isNaN(nextDate.getTime()) &&
          today.getTime() > nextDate.getTime()
        ) {

          lastDate = new Date(nextDate.getTime());

          nextDate = this.addFrequency(
            nextDate,
            item.Frequency
          );

          changed = true;
        }

        if (changed) {

  updates.push(
    this.createHistoryRecord(item)
  );

  item.LastMaintenance =
    `${this.isoDate(lastDate)}T00:00:00Z`;

  item.NextMaintenance =
    `${this.isoDate(nextDate)}T00:00:00Z`;

  updates.push(
    this.updateMaintenanceDates(
      item.Id,
      lastDate,
      nextDate
    )
  );
}
      });

      if (updates.length > 0) {
        await Promise.all(updates);
      }

      const rows: IPMRow[] = items.map(item => {

        const lastDate = this.parseDate(item.LastMaintenance);
        const nextDate = this.parseDate(item.NextMaintenance);

        const daysUntil = Math.round(
          (nextDate.getTime() - today.getTime()) /
          (1000 * 60 * 60 * 24)
        );

        let status: 'ok' | 'soon' | 'overdue' = 'ok';

        if (daysUntil < 0) {
          status = 'overdue';
        }
        else if (daysUntil <= this.state.threshold) {
          status = 'soon';
        }

        return {
          ...item,
          lastDate,
          nextDate,
          daysUntil,
          status
        };
      });

      this.setState({
        items: rows,
        loading: false
      });
    })

    .catch((error: Error) => {

      console.error(error);

      this.setState({
        loading: false,
        error: error.message || 'Could not load SharePoint data.'
      });
    });
}

  private contextHttpGet(url: string): Promise<IPMItem[]> {
    return this.props.context.spHttpClient
      .get(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata=nometadata'
          }
        } as ISPHttpClientOptions
      )
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          throw new Error(
            `SharePoint REST error: ${response.status}`
          );
        }

        return response.json();
      })
      .then((json: any) => json.value || []);
  }

  private getFilteredItems(): IPMRow[] {
    let data = this.state.items.slice();

    if (this.state.filter !== 'all') {
      if (this.state.filter === 'next90') {
        data = data.filter(
          x => x.daysUntil >= 0 && x.daysUntil <= 90
        );
      } else {
        data = data.filter(
          x => x.status === this.state.filter
        );
      }
    }

    const search = this.state.search.toLowerCase();

    if (search) {
      data = data.filter(x =>
        `${x.Service} ${x.Location} ${x.BU} ${x.AffectedSystems}`
          .toLowerCase()
          .indexOf(search) !== -1
      );
    }

    if (this.state.bu !== 'all') {
      data = data.filter(x => x.BU === this.state.bu);
    }

    data.sort((a, b) => a.daysUntil - b.daysUntil);

    return data;
  }

  private getCounts(): {
    total: number;
    ok: number;
    soon: number;
    overdue: number;
    next90: number;
  } {
    const items = this.state.items;

    return {
      total: items.length,
      ok: items.filter(x => x.status === 'ok').length,
      soon: items.filter(x => x.status === 'soon').length,
      overdue: items.filter(x => x.status === 'overdue').length,
      next90: items.filter(
        x => x.daysUntil >= 0 && x.daysUntil <= 90
      ).length
    };
  }

  private getBUs(): string[] {
    return this.state.items
      .map(x => x.BU)
      .filter((value, index, array) =>
        value && array.indexOf(value) === index
      )
      .sort();
  }
    private statusColor(status: string): string {
    if (status === 'overdue') return '#E24B4A';
    if (status === 'soon') return '#BA7517';
    if (status === 'completed') return '#888780';
    if (status === 'projected') return '#7F77DD';
    return '#1D9E75';
  }

  private truncate(str: string, n: number): string {
    if (!str) return '';
    return str.length > n ? str.slice(0, n - 1) + '\u2026' : str;
  }

  private buildCalendarEvents(): ICalEvent[] {
    const events: ICalEvent[] = [];
    const items = this.state.items;

    items.forEach(r => {
      const systems = (r.AffectedSystems || '')
        .split(',')
        .map(s => s.trim())
        .filter(Boolean);

      const intervalDays = Math.round(
        (r.nextDate.getTime() - r.lastDate.getTime()) / (1000 * 60 * 60 * 24)
      );

      if (this.state.showCompleted) {
        events.push({
          id: r.Id,
          service: r.Service,
          bu: r.BU,
          location: r.Location,
          time: r.Time,
          duration: r.Duration,
          frequency: r.Frequency,
          systems,
          date: r.lastDate,
          status: 'completed',
          kind: 'completed'
        });
      }

      events.push({
        id: r.Id,
        service: r.Service,
        bu: r.BU,
        location: r.Location,
        time: r.Time,
        duration: r.Duration,
        frequency: r.Frequency,
        systems,
        date: r.nextDate,
        status: r.status,
        kind: 'scheduled',
        daysUntil: r.daysUntil
      });

      if (this.state.showProjected && intervalDays > 0) {
        for (let cyc = 1; cyc <= 4; cyc++) {
          const projDate = new Date(r.nextDate.getTime());
          projDate.setDate(projDate.getDate() + intervalDays * cyc);

          events.push({
            id: r.Id,
            service: r.Service,
            bu: r.BU,
            location: r.Location,
            time: r.Time,
            duration: r.Duration,
            frequency: r.Frequency,
            systems,
            date: projDate,
            status: 'projected',
            kind: 'projected'
          });
        }
      }
    });

    return events;
  }

  private getCalendarEvents(): ICalEvent[] {
    const all = this.buildCalendarEvents();

    if (this.state.calBu === 'all') {
      return all;
    }

    return all.filter(ev => ev.bu === this.state.calBu);
  }

  private eventsForDate(events: ICalEvent[], y: number, m: number, d: number): ICalEvent[] {
    return events.filter(ev =>
      ev.date.getFullYear() === y &&
      ev.date.getMonth() === m &&
      ev.date.getDate() === d
    );
  }

  private eventsInMonth(events: ICalEvent[], y: number, m: number): ICalEvent[] {
    return events.filter(ev =>
      ev.date.getFullYear() === y &&
      ev.date.getMonth() === m
    );
  }

  private findRowById(id: number): IPMRow | null {
    for (let i = 0; i < this.state.items.length; i++) {
      if (this.state.items[i].Id === id) return this.state.items[i];
    }
    return null;
  }
  private static MONTH_NAMES = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];

  private calShiftMonth = (delta: number): void => {
    let { calYear, calMonth } = this.state;
    calMonth += delta;

    if (calMonth < 0) {
      calMonth = 11;
      calYear--;
    }
    if (calMonth > 11) {
      calMonth = 0;
      calYear++;
    }

    this.setState({ calYear, calMonth, calSelectedDate: null });
  }

  private calGoToday = (): void => {
    const today = new Date();

    this.setState({
      calYear: today.getFullYear(),
      calMonth: today.getMonth(),
      calSelectedDate: today.toDateString()
    });
  }

  private calJumpMonth = (event: React.ChangeEvent<HTMLSelectElement>): void => {
    this.setState({
      calMonth: Number(event.target.value),
      calSelectedDate: null
    });
  }

  private calJumpYear = (event: React.ChangeEvent<HTMLSelectElement>): void => {
    this.setState({
      calYear: Number(event.target.value),
      calSelectedDate: null
    });
  }

  private calSelectDay = (y: number, m: number, d: number): void => {
    this.setState({
      calSelectedDate: new Date(y, m, d).toDateString()
    });
  }

  private calYearOptions(): number[] {
    const thisYear = new Date().getFullYear();
    const years: number[] = [];
    for (let y = thisYear - 1; y <= thisYear + 2; y++) {
      years.push(y);
    }
    return years;
  }

  private exportMonthICS = (): void => {
    const events = this.getCalendarEvents().filter(ev =>
      ev.kind === 'scheduled' &&
      ev.date.getFullYear() === this.state.calYear &&
      ev.date.getMonth() === this.state.calMonth
    );

    if (events.length === 0) {
      alert('No scheduled (Next Maintenance) events in this month to export.');
      return;
    }

    const pad = (n: number): string => (n < 10 ? '0' + n : String(n));

    const toICSDate = (date: Date, timeStr: string): string => {
      let hh = '09';
      let mm = '00';
      const match = timeStr ? timeStr.match(/(\d{2}):(\d{2})/) : null;
      if (match) {
        hh = match[1];
        mm = match[2];
      }
      return `${date.getFullYear()}${pad(date.getMonth() + 1)}${pad(date.getDate())}T${hh}${mm}00`;
    };

    let ics = 'BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//SCHOTT PM Dashboard//EN\r\n';

    events.forEach((ev, i) => {
      const dtStart = toICSDate(ev.date, ev.time);
      ics += 'BEGIN:VEVENT\r\n' +
        `UID:${ev.service}-${this.state.calYear}${this.state.calMonth}-${i}@schott\r\n` +
        `DTSTAMP:${toICSDate(new Date(), '09:00')}Z\r\n` +
        `DTSTART:${dtStart}\r\n` +
        `SUMMARY:PM - ${ev.service} (${ev.location})\r\n` +
        `DESCRIPTION:BU: ${ev.bu} | Duration: ${ev.duration} | Systems: ${ev.systems.join(', ')}\r\n` +
        'END:VEVENT\r\n';
    });

    ics += 'END:VCALENDAR';

    const blob = new Blob([ics], { type: 'text/calendar' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `pm_maintenance_${MipsPreventiveMaintenance.MONTH_NAMES[this.state.calMonth]}_${this.state.calYear}.ics`;
    a.click();
  }
    private renderCalSummary(): React.ReactElement<any> {
    const monthEvents = this.eventsInMonth(
      this.getCalendarEvents(), this.state.calYear, this.state.calMonth
    ).filter(ev => ev.kind === 'scheduled');

    const total = monthEvents.length;
    const overdue = monthEvents.filter(ev => ev.status === 'overdue').length;
    const soon = monthEvents.filter(ev => ev.status === 'soon').length;
    const ok = monthEvents.filter(ev => ev.status === 'ok').length;
    const monthLabel = `${MipsPreventiveMaintenance.MONTH_NAMES[this.state.calMonth]} ${this.state.calYear}`;

    return (
      <div className={styles.calSummary}>
        <div className={styles.chip}>{monthLabel}</div>
        <div className={styles.chip}><b>{total}</b>&nbsp;scheduled</div>
        <div className={styles.chip} style={{ color: '#791F1F' }}><b>{overdue}</b>&nbsp;overdue</div>
        <div className={styles.chip} style={{ color: '#633806' }}><b>{soon}</b>&nbsp;due soon</div>
        <div className={styles.chip} style={{ color: '#085041' }}><b>{ok}</b>&nbsp;on schedule</div>
      </div>
    );
  }

  private renderChip(ev: ICalEvent): React.ReactElement<any> {
    const kindLabel = ev.kind === 'completed' ? ' (done)' : ev.kind === 'projected' ? ' (est.)' : '';
    const sysLabel = ev.systems.length > 0 ? ev.systems.join(', ') : '';
    const title = `${ev.service} (${ev.bu})${kindLabel} - ${sysLabel}`;

    const bg =
      ev.status === 'overdue' ? '#FCEBEB' :
      ev.status === 'soon' ? '#FAEEDA' :
      ev.status === 'completed' ? 'var(--calBgTertiary)' :
      ev.status === 'projected' ? 'transparent' : '#E1F5EE';

    const fg =
      ev.status === 'overdue' ? '#791F1F' :
      ev.status === 'soon' ? '#633806' :
      ev.status === 'completed' ? '#5f5e5a' :
      ev.status === 'projected' ? '#5f5e5a' : '#085041';

    const chipClass = ev.kind === 'projected'
      ? `${styles.calChip} ${styles.calChipProjected}`
      : styles.calChip;

    return (
      <div
        key={`${ev.id}-${ev.kind}-${ev.date.getTime()}`}
        className={chipClass}
        title={title}
        style={{ background: bg, color: fg }}
        onClick={e => {
          e.stopPropagation();
          if (ev.kind === 'scheduled' && this.props.allowEdit) {
            const row = this.findRowById(ev.id);
            if (row) this.openEdit(row);
          }
        }}
      >
        <div className={styles.chipService}>
          <span className={styles.calChipDot} style={{ background: this.statusColor(ev.status) }} />
          {this.truncate(ev.service, 14)}{kindLabel}
        </div>
        <div className={styles.chipMeta}>BU: {ev.bu}</div>
        <div className={styles.chipSys}>{this.truncate(sysLabel, 32)}</div>
      </div>
    );
  }

  private renderCalendarGrid(): React.ReactElement<any> {
    const events = this.getCalendarEvents();
    const { calYear, calMonth, calSelectedDate } = this.state;

    const firstOfMonth = new Date(calYear, calMonth, 1);
    const daysInMonth = new Date(calYear, calMonth + 1, 0).getDate();
    let startOffset = firstOfMonth.getDay() - 1;
    if (startOffset < 0) startOffset = 6;

    const todayStr = new Date().toDateString();
    const cells: React.ReactElement<any>[] = [];

    for (let i = 0; i < startOffset; i++) {
      cells.push(<div key={`empty-${i}`} className={`${styles.calCell} ${styles.calCellEmpty}`} />);
    }

    for (let d = 1; d <= daysInMonth; d++) {
      const dateObj = new Date(calYear, calMonth, d);
      const dStr = dateObj.toDateString();
      const dow = dateObj.getDay();
      const isWeekend = dow === 0 || dow === 6;
      const evs = this.eventsForDate(events, calYear, calMonth, d);
      const hasOverdue = evs.some(ev => ev.kind === 'scheduled' && ev.status === 'overdue');
      const isToday = dStr === todayStr;
      const isSelected = dStr === calSelectedDate;
      const maxChips = 2;

      let cellClass = styles.calCell;
      if (isWeekend) cellClass += ` ${styles.calCellWeekend}`;
      if (isToday) cellClass += ` ${styles.calCellToday}`;
      if (isSelected) cellClass += ` ${styles.calCellSelected}`;
      if (hasOverdue) cellClass += ` ${styles.calCellOverdue}`;

      cells.push(
        <div
          key={d}
          className={cellClass}
          onClick={() => this.calSelectDay(calYear, calMonth, d)}
        >
          <div className={isToday ? `${styles.calDaynum} ${styles.calDaynumToday}` : styles.calDaynum}>
            {d}
          </div>
          {evs.length > 0 && <span className={styles.calCount}>{evs.length}</span>}
          <div className={styles.calEvents}>
            {evs.slice(0, maxChips).map(ev => this.renderChip(ev))}
            {evs.length > maxChips && (
              <div className={styles.calMore}>+{evs.length - maxChips} more</div>
            )}
          </div>
        </div>
      );
    }

    return <div className={styles.calGrid}>{cells}</div>;
  }

  private renderAgenda(): React.ReactElement<any> {
    const { calSelectedDate } = this.state;

    if (!calSelectedDate) {
      return <div className={styles.small}>Click a day above to see scheduled maintenance for that date.</div>;
    }

    const dateObj = new Date(calSelectedDate);
    const events = this.getCalendarEvents();
    const evs = this.eventsForDate(events, dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate());

    const label = dateObj.toLocaleDateString('en-GB', {
      weekday: 'long', day: '2-digit', month: 'long', year: 'numeric'
    });

    if (evs.length === 0) {
      return (
        <div className={styles.small}>
          <strong>{label}</strong> — no maintenance recorded or scheduled.
        </div>
      );
    }

    const scheduledCount = evs.filter(ev => ev.kind === 'scheduled').length;

    return (
      <div>
        <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 8 }}>
          {label} ({evs.length})
        </div>

        {scheduledCount > 1 && (
          <div className={styles.small} style={{ color: '#791F1F', marginBottom: 8 }}>
            Note: {scheduledCount} services scheduled the same day — check for resource/technician conflicts.
          </div>
        )}

        {evs.map((ev, idx) => {
          const kindBadge = ev.kind === 'completed' ? (
            <span className={`${styles.badge}`} style={{ background: 'var(--calBgTertiary)', color: '#5f5e5a' }}>
              Completed
            </span>
          ) : ev.kind === 'projected' ? (
            <span className={`${styles.badge}`} style={{ background: 'var(--calBgTertiary)', color: '#5f5e5a', border: '1px dashed #b4b2a9' }}>
              Projected (est.)
            </span>
          ) : (
            <span className={`${styles.badge} ${styles[ev.status]}`}>
              {ev.status === 'overdue' ? `Overdue ${Math.abs(ev.daysUntil || 0)}d` :
               ev.status === 'soon' ? `Due in ${ev.daysUntil}d` : 'On schedule'}
            </span>
          );

          return (
            <div key={idx} className={styles.agendaItem}>
              <div className={styles.rowTop}>
                <div className={styles.l}>
                  <span>{ev.service} &middot; {ev.location}</span>
                  <span className={styles.meta}>{ev.time} &middot; {ev.duration} &middot; BU: {ev.bu}</span>
                </div>
                <div style={{ display: 'flex', alignItems: 'center' }}>
                  {kindBadge}
                  {ev.kind === 'scheduled' && this.props.allowEdit && (
                    <button
                      className={styles.editButton}
                      style={{ marginLeft: 8 }}
                      onClick={() => {
                        const row = this.findRowById(ev.id);
                        if (row) this.openEdit(row);
                      }}
                    >
                      Edit
                    </button>
                  )}
                </div>
              </div>
              <div>
                {ev.systems.map(s => (
                  <span key={s} className={styles.systemTag}>{s}</span>
                ))}
              </div>
            </div>
          );
        })}
      </div>
    );
  }

  private renderCalendar(): React.ReactElement<any> {
    const bus = this.getBUs();

    return (
      <div>
        {this.renderCalSummary()}

        <div className={styles.calToolbar}>
          <button onClick={() => this.calShiftMonth(-1)}>&#8592;</button>

          <select value={this.state.calMonth} onChange={this.calJumpMonth}>
            {MipsPreventiveMaintenance.MONTH_NAMES.map((m, i) => (
              <option key={m} value={i}>{m}</option>
            ))}
          </select>

          <select value={this.state.calYear} onChange={this.calJumpYear}>
            {this.calYearOptions().map(y => (
              <option key={y} value={y}>{y}</option>
            ))}
          </select>

          <button onClick={() => this.calShiftMonth(1)}>&#8594;</button>
          <button onClick={this.calGoToday}>Today</button>

          <select
            style={{ marginLeft: 'auto' }}
            value={this.state.calBu}
            onChange={e => this.setState({ calBu: e.target.value })}
          >
            <option value="all">All BUs</option>
            {bus.map(bu => <option key={bu} value={bu}>{bu}</option>)}
          </select>

          <button onClick={this.exportMonthICS}>Export month (.ics)</button>
        </div>

        <div className={styles.calToggleRow}>
          <label>
            <input
              type="checkbox"
              checked={this.state.showCompleted}
              onChange={e => this.setState({ showCompleted: e.target.checked })}
            />
            Show completed (last maintenance)
          </label>

          <label>
            <input
              type="checkbox"
              checked={this.state.showProjected}
              onChange={e => this.setState({ showProjected: e.target.checked })}
            />
            Show projected future cycles (estimated)
          </label>
        </div>

        {this.renderCalendarGrid()}

        <div className={styles.agendaPanel}>
          {this.renderAgenda()}
        </div>
      </div>
    );
  }
  private statusText(item: IPMRow): string {
    if (item.status === 'overdue') {
      return `Overdue ${Math.abs(item.daysUntil)}d`;
    }

    if (item.status === 'soon') {
      return `Due in ${item.daysUntil}d`;
    }

    return 'On schedule';
  }

  private openEdit = (item: IPMRow): void => {
    if (!this.props.allowEdit) {
      return;
    }

    this.setState({
      editing: item,
      editDate: this.isoDate(item.nextDate)
    });
  }

  private closeEdit = (): void => {
    this.setState({
      editing: null,
      editDate: ''
    });
  }

    private saveEdit = (): void => {
    const item = this.state.editing;

    if (!item || !this.state.editDate) {
      return;
    }

    const url =
      `${this.props.siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(this.props.listName)}')/items(${item.Id})`;

    const body: any = {};
    body[this.fieldMap.NextMaintenance] =
      `${this.state.editDate}T00:00:00Z`;

    const options: ISPHttpClientOptions = {
      headers: {
        Accept: 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE'
      },
      body: JSON.stringify(body)
    };

    this.props.context.spHttpClient
      .post(url, SPHttpClient.configurations.v1, options)
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          return response.text().then(text => {
            throw new Error(`SharePoint update failed: ${response.status} ${text}`);
          });
        }

        this.closeEdit();
        this.loadData();
      })
      .catch(error => {
        console.error(error);
        alert(`Could not save the maintenance date.\n\n${error.message}`);
      });
  }

 

   private getListItemType(): Promise<string> {
    const url =
      `${this.props.siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(this.props.listName)}')` +
      '?$select=ListItemEntityTypeFullName';

    return this.props.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1, {
        headers: { Accept: 'application/json;odata=nometadata' }
      } as ISPHttpClientOptions)
      .then((response: SPHttpClientResponse) => response.json())
      .then((json: any) => json.ListItemEntityTypeFullName);
  }

  private renderKpis(): React.ReactElement<any> {
    const c = this.getCounts();

    const cards = [
      ['all', 'Total services', c.total, 'All tracked'],
      ['ok', 'On schedule', c.ok, 'No action needed'],
      ['soon', 'Due soon', c.soon, `Within ${this.state.threshold}d`],
      ['overdue', 'Overdue', c.overdue, 'Needs action'],
      ['next90', 'Next 90 days', c.next90, 'Upcoming window']
    ];

    return (
      <div className={styles.kpiGrid}>
        {cards.map(card => (
          <button
            key={card[0] as string}
            className={
              `${styles.kpiCard} ${
                this.state.filter === card[0] ? styles.active : ''
              }`
            }
            onClick={() =>
              this.setState({
                filter:
                  this.state.filter === card[0]
                    ? 'all'
                    : card[0] as string
              })
            }
          >
            <span className={styles.kpiLabel}>
              {card[1]}
            </span>

            <span className={styles.kpiValue}>
              {card[2]}
            </span>

            <span className={styles.kpiTag}>
              {card[3]}
            </span>
          </button>
        ))}
      </div>
    );
  }

  private renderTable(): React.ReactElement<any> {
    const data = this.getFilteredItems();

    return (
      <div className={styles.tableWrapper}>
        <table>
          <thead>
            <tr>
              <th>Service</th>
              <th>BU</th>
              <th>Location</th>
              <th>Last maintenance</th>
              <th>Next maintenance</th>
              <th>Days until</th>
              <th>Status</th>
            </tr>
          </thead>

          <tbody>
            {data.map(item => (
              <React.Fragment key={item.Id}>
                <tr
                  className={styles.dataRow}
                  onClick={() =>
                    this.setState({
                      expandedId:
                        this.state.expandedId === item.Id
                          ? null
                          : item.Id
                    })
                  }
                >
                  <td>{item.Service}</td>
                  <td>{item.BU}</td>
                  <td>{item.Location}</td>
                  <td>{this.formatDate(item.lastDate)}</td>
                  <td>{this.formatDate(item.nextDate)}</td>
                  <td>{item.daysUntil}</td>

                  <td>
                    <div className={styles.statusCell}>
                      <span
                        className={`${styles.badge} ${styles[item.status]}`}
                      >
                        {this.statusText(item)}
                      </span>

                      {this.props.allowEdit && (
                        <button
                          className={styles.editButton}
                          onClick={event => {
                            event.stopPropagation();
                            this.openEdit(item);
                          }}
                        >
                          Edit
                        </button>
                      )}
                    </div>
                  </td>
                </tr>

                {this.state.expandedId === item.Id && (
                  <tr className={styles.detailRow}>
                    <td colSpan={7}>
                      <div className={styles.detailGrid}>
                        <div>
                          <span>Time window</span>
                          {item.Time}
                        </div>

                        <div>
                          <span>Duration</span>
                          {item.Duration}
                        </div>

                        <div>
                          <span>Frequency</span>
                          {item.Frequency}
                        </div>
                      </div>

                      <div>
                        <span className={styles.small}>
                          Affected systems
                        </span>

                        <div className={styles.systems}>
                          {(item.AffectedSystems || '')
                            .split(',')
                            .map(system => (
                              <span
                                key={system}
                                className={styles.systemTag}
                              >
                                {system.trim()}
                              </span>
                            ))}
                        </div>
                      </div>
                    </td>
                  </tr>
                )}
              </React.Fragment>
            ))}
          </tbody>
        </table>

        <div className={styles.footerNote}>
          Showing {data.length} of {this.state.items.length} services
        </div>
      </div>
    );
  }

  private renderEditModal(): React.ReactElement<any> | null {
    const item = this.state.editing;

    if (!item) {
      return null;
    }

    return (
      <div className={styles.modalOverlay}>
        <div className={styles.modal}>
          <h3>
            Edit next maintenance date
          </h3>

          <p>
            {item.Service} — {item.Location}
          </p>

          <label>
            New maintenance date
          </label>

          <input
            type="date"
            value={this.state.editDate}
            onChange={event =>
              this.setState({
                editDate: event.target.value
              })
            }
          />

          <div className={styles.modalActions}>
            <button onClick={this.closeEdit}>
              Cancel
            </button>

            <button
              className={styles.primaryButton}
              onClick={this.saveEdit}
            >
              Save to SharePoint
            </button>
          </div>
        </div>
      </div>
    );
  }

  public render(): React.ReactElement<any> {
    const bus = this.getBUs();

    return (
      <section className={styles.mipsPreventiveMaintenance}>

        <div className={styles.header}>
          <div>
            <h2>
              MIPS preventive maintenance dashboard
            </h2>

            <p>
              Live data from Microsoft List
            </p>
          </div>

          <button onClick={this.loadData}>
            Refresh
          </button>
        </div>

        {this.state.error && (
          <div className={styles.error}>
            {this.state.error}
          </div>
        )}

        {this.renderKpis()}

        <div className={styles.threshold}>
          <span>
            Due soon window:
          </span>

          <input
            type="range"
            min="3"
            max="30"
            value={this.state.threshold}
            onChange={event =>
              this.setState({
                threshold: Number(event.target.value)
              }, this.loadData)
            }
          />

          <span>
            {this.state.threshold} days
          </span>
        </div>

                      <div className={styles.card}>

          <div className={styles.tabbar}>
            <button
              className={`${styles.tab} ${this.state.activeTab === 'calendar' ? styles.tabActive : ''}`}
              onClick={() => this.setState({ activeTab: 'calendar' })}
            >
              Calendar
            </button>
            <button
              className={`${styles.tab} ${this.state.activeTab === 'log' ? styles.tabActive : ''}`}
              onClick={() => this.setState({ activeTab: 'log' })}
            >
              Full log
            </button>
            
          </div>

          {this.state.activeTab === 'log' ? (
            <div>
              <div className={styles.controls}>

                <input
                  type="text"
                  placeholder="Search service, location, system..."
                  value={this.state.search}
                  onChange={event =>
                    this.setState({
                      search: event.target.value
                    })
                  }
                />

                <select
                  value={this.state.bu}
                  onChange={event =>
                    this.setState({
                      bu: event.target.value
                    })
                  }
                >
                  <option value="all">
                    All BUs
                  </option>

                  {bus.map(bu => (
                    <option key={bu} value={bu}>
                      {bu}
                    </option>
                  ))}
                </select>

                <button
                  onClick={() =>
                    this.setState({
                      search: '',
                      bu: 'all',
                      filter: 'all'
                    })
                  }
                >
                  Reset filters
                </button>
              </div>

              {this.state.loading ? (
                <div className={styles.loading}>
                  Loading SharePoint data...
                </div>
              ) : (
                this.renderTable()
              )}
            </div>
          ) : (
            this.renderCalendar()
          )}

        </div>

        {this.renderEditModal()}

      </section>
    );
  }
}
