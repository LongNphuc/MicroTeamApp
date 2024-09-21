import React from 'react';
import { Nav } from './calendar/Nav'
import { ScheduleWeek } from './calendar/ScheduleWeek'
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IGraphApiService } from "../services/IGraphApiService";
import { CalendarProvider, useCalendar } from '../contexts/calendar/CalendarContext';
import { ScheduleType } from '../utils/EnumConst';
import { ScheduleDay } from './calendar/ScheduleDay';

export interface ICalendarProps {

}

export function Calendar({}: ICalendarProps) {

  return (
    <CalendarProvider>
      <Nav
      />
      <ScheduleWrapper />
    </CalendarProvider>
  );
}

function ScheduleWrapper({}: ICalendarProps) {
  const { scheduleTypeState } = useCalendar();
  const [scheduleTypeSelected] = scheduleTypeState;

  return (
    <>
      {(scheduleTypeSelected === ScheduleType.Week || scheduleTypeSelected === ScheduleType.WorkWeek) && (
        <ScheduleWeek
          key={scheduleTypeSelected}
          // context={context}
          // graphApiService={graphApiService}
        />
      )}
      {(scheduleTypeSelected === ScheduleType.Day) && (
        <ScheduleDay
          // context={context}
          // graphApiService={graphApiService}
        />
      )}
    </>
  );
}
