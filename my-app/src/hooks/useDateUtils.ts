import moment from "moment";

export const useDateUtils = () => {
  const getStartOfWeek = (date: Date, startDay: number = 0): Date => {
    const day = date.getDay();
    const diff = (day < startDay ? -7 : 0) + day - startDay;
    const startOfWeek = new Date(date.toString());
    startOfWeek.setDate(date.getDate() - diff);
    startOfWeek.setHours(0, 0, 0, 0);
    return startOfWeek;
  };

  const getWeekDays = (date: Date, startDay: number = 0): Date[] => {
    const startOfWeek = getStartOfWeek(date, startDay);
    const weekDays: Date[] = [];

    for (let i = 0; i < 7; i++) {
      const day = new Date(startOfWeek.toString());
      day.setDate(startOfWeek.getDate() + i);
      weekDays.push(day);
    }
    return weekDays;
  };

  const getWeekCurrent = () => {
    const today = new Date();
    const weekDays = getWeekDays(today, 0);
    return weekDays;
  };

  const utcDateTime = (selectedAppointmentDate: Date, time: Date): string => {
    const combinedDateTime = moment(selectedAppointmentDate).set({
      hour: time.getHours(),
      minute: time.getMinutes(),
      second: 0,
      millisecond: 0,
    });
    return moment(combinedDateTime).utc().format("YYYY-MM-DDTHH:mm:ss.SSS[Z]");
  };

  const getDayOfWeek = (date: Date): string => {
    const options: Intl.DateTimeFormatOptions = { weekday: "long" };
    return new Intl.DateTimeFormat("en-US", options).format(date);
  };

  const generateTimeSlots = (date: Date): Date[] => {
    const slots: Date[] = [];
    for (let hour = 0; hour < 24; hour++) {
      slots.push(new Date(date.setHours(hour, 0, 0, 0)));
      slots.push(new Date(date.setHours(hour, 30, 0, 0)));
    }
    return slots;
  };

  return { getWeekDays, getWeekCurrent, utcDateTime, getDayOfWeek, generateTimeSlots};
};
