import React, { createContext, ReactNode, useContext, useState } from "react";
import { ScheduleType } from "../../utils/EnumConst";
import { DateRange } from "../../models/DateRange";
import { useDateUtils } from "../../hooks/useDateUtils";

interface CalendarContextProps {
  scheduleTypeState: [ScheduleType, (type: ScheduleType) => void];
  scheduleData: DateRange;
  setDateRange: (dateRange: Date[]) => void;
  pageRenderState : [boolean, (state: boolean) => void];
}

interface CalendarProviderProps {
  children: ReactNode;
}

export const CalendarContext = createContext<CalendarContextProps | undefined>(
  undefined
);

export const CalendarProvider: React.FC<CalendarProviderProps> = ({children,}) => {
  const { getWeekCurrent } = useDateUtils();
  const [scheduleTypeSelected, setScheduleTypeSelected] =useState<ScheduleType>(ScheduleType.Week);
  const [dateRange, setDateRange] = useState<Date[]>(getWeekCurrent());
  const [isRender, setIsRender] = useState<boolean>(false);
  let scheduleData: DateRange;

  if (scheduleTypeSelected === ScheduleType.Day || scheduleTypeSelected === ScheduleType.Agenda) {
    scheduleData = {
      startDate: dateRange[0],
      endDate: dateRange[0],
    };
  } else {
    if (scheduleTypeSelected === ScheduleType.Week) {
      scheduleData = {
        startDate: dateRange[0],
        endDate: dateRange[6],
      };
    } else {
      scheduleData = {
        startDate: dateRange[1],
        endDate: dateRange[5],
      };
    }
  }

  return (
    <CalendarContext.Provider
      value={{
        scheduleTypeState: [scheduleTypeSelected, setScheduleTypeSelected],
        scheduleData,
        setDateRange,
        pageRenderState : [isRender, setIsRender]
      }}
    >
      {children}
    </CalendarContext.Provider>
  );
};

export const useCalendar = () => {
  const context = useContext(CalendarContext);
  if (!context) {
    throw new Error("useCalendar must be used within an CalendarProvider");
  }
  return context;
};
