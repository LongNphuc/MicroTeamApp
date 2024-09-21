import React, { useCallback, useEffect, useMemo, useState } from "react";
import { Stack, Text } from "@fluentui/react";
import { Appointment } from "./Appointment";
import moment from "moment-timezone";
import { AppointmentModel, UTCDateRange } from "../../models/AppointmentModel";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IGraphApiService } from "../../services/IGraphApiService";
import { styles } from "../../styles/Calendar/ScheduleStyles";
import { useCalendar } from "../../contexts/calendar/CalendarContext";
import { Loading } from "../Loading";
import { useDateUtils } from "../../hooks/useDateUtils";

export interface DateRangeProps {

}

export function ScheduleDay({}: DateRangeProps) {
  const { scheduleData, pageRenderState, scheduleTypeState } = useCalendar();
  const [isRender, setIsRender] = pageRenderState;
  const [scheduleTypeSelected] = scheduleTypeState;

  const {getDayOfWeek, generateTimeSlots} = useDateUtils();

  const [schedule, setSchedule] = useState<AppointmentModel[]>([]);
  const [currentDay] = useState(new Date());

  const selectedDate = new Date(scheduleData.startDate);
  selectedDate.setHours(0, 0, 0, 0);

  const dayOfWeek = useMemo(() => getDayOfWeek(selectedDate),[selectedDate]);

  const timeSlots = generateTimeSlots(new Date(selectedDate.toString()));

  // const genSchedule = useCallback(async () => {
  //   const currentTimezone = moment.tz.guess();
  //   const startDate = moment.tz(selectedDate, currentTimezone).utc().format();
  //   const endDate = moment
  //     .tz(selectedDate, currentTimezone)
  //     .endOf("day")
  //     .utc()
  //     .format();
  //   const utcDateRange: UTCDateRange = { start: startDate, end: endDate };
  //   // const result = await graphApiService.getCalendarEvent(
  //   //   context,
  //   //   utcDateRange
  //   // );
  //   setSchedule(result);
  // }, [selectedDate, context, graphApiService]);

  useEffect(() => {
    const fetchData = async () => {
      try {
        // await genSchedule();
        if (isRender) setIsRender(false);
      } catch (error) {
        console.error("Error occurred while fetching schedule: ", error);
      }
    };
  
    fetchData();
  }, [scheduleData, isRender, scheduleTypeSelected]);

  useEffect(() => {
    const calendarBody = document.querySelector(".calendar-body")!;
    const timeLines: NodeListOf<HTMLElement> = document.querySelectorAll(".time-line");
    const timeLineActive: HTMLElement | null = document.querySelector(".time-line.active");

    const updatePosition = () => {
      const currentHour = new Date().getHours();
      const currentMinutes = new Date().getMinutes();
      const currentPosition = (currentHour / 24) * calendarBody.clientHeight + (currentMinutes / 60) * 82;

      if (timeLines) {
        timeLines.forEach((element: HTMLElement) => {
          element.style.top = `${currentPosition}px`;
        });
      }

      if (timeLineActive) {
        const position = timeLineActive.getBoundingClientRect();
        const body = document.body;
        const offsetY = position.top - body.getBoundingClientRect().top;
        window.scrollTo({
          top: offsetY,
          left: 0,
        });
      }
    };

    updatePosition();

    const intervalId = setInterval(updatePosition, 60000);

    return () => clearInterval(intervalId);
  }, [selectedDate, currentDay]);

  const calculateTimelinePosition = () => {
    const currentHour = currentDay.getHours();
    const currentMinute = currentDay.getMinutes();
    const position = (currentHour * 60 + currentMinute) * (41 / 60);
    return position;
  };

  const times: string[] = timeSlots.map((slot) => {
    const hours = slot.getHours();
    const minutes = slot.getMinutes();
    if (minutes === 30) return "";
    if (hours === 0) return "12AM";
    if (hours === 12) return "12PM";
    if (hours > 12) return `${hours - 12}PM`;
    return `${hours}AM`;
  });

  return (
    <>
    {schedule ? (
        <>
      <div className="calendar-head">
        <div
          style={{
            paddingLeft: 51,
            display: "flex",
            boxSizing: "border-box",
            justifyContent: "center",
            borderBottom: "1px solid #e0e0e0",
          }}>
          <Text
            variant="xLarge"
            style={{
              boxSizing: "border-box",
              paddingBottom: 6,
            }}>
            <Text
              styles={{
                root: {
                  fontSize: 24,
                  lineHeight: "32px",
                  display: "block",
                  color: "#6264A7",
                  fontWeight: "600",
                  marginLeft: 8,
                },
              }}>
              {selectedDate.getDate()}
            </Text>
            <Text
              styles={{
                root: {
                  fontSize: 12,
                  lineHeight: "16px",
                  color: "#6264A7",
                  fontWeight: "600",
                  marginLeft: 8,
                },
              }}>
              {dayOfWeek}
            </Text>
          </Text>
        </div>
      </div>

      <div className="calendar-container">
        <div className="calendar">
          <div className="calendar-body" style={{ display: "flex" }}>
            <Stack
              styles={{
                root: {
                  width: 51,
                  borderTop: "1px solid #e0e0e0",
                  marginTop: "-1px",
                },
              }}>
              {times.map((time, i) => (
                <Text
                  key={i}
                  styles={{
                    root: { height: 41, textAlign: "center", fontSize: 12 },
                  }}
                >
                  {time}
                </Text>
              ))}
            </Stack>
            <div className="calendar-cell-list" style={{ display: "flex", boxSizing: "border-box", flex: 1 }}>
              <div className="calendar-col" style={{ width: "100%", position: "relative" }}>
                {timeSlots.map((slot, index) => (
                  <div
                    key={index}
                    className="calendar-cell"
                    style={{
                      height: 41,
                      width: "100%",
                      borderBottom: slot.getMinutes() === 0 ? "1px dashed #e0e0e0" : "1px solid #e0e0e0",
                      boxSizing: "border-box",
                    }}/>
                ))}

                <div className="contain-appointment" style={styles.containAppointment}>
                  {schedule.map((e) => {
                    const startDate = new Date(moment.utc(e.start.dateTime).local().format("YYYY-MM-DD HH:mm:ss"));
                    const endDate = new Date(moment.utc(e.end.dateTime).local().format("YYYY-MM-DD HH:mm:ss"));
                    if (startDate >= selectedDate &&  endDate <= new Date(selectedDate.getTime() + 86400000)) {
                      const top = 41 * (startDate.getHours() + startDate.getMinutes() / 60);
                      const height = 41 * (endDate.getHours() + endDate.getMinutes() / 60 - startDate.getHours() - startDate.getMinutes() / 60);
                      const inlineStyle = {
                        padding: "0 4px",
                        boxSizing: "border-box",
                        backgroundColor: "#DCE0FA",
                        borderLeft: "4px solid #5B5FC7",
                        borderRadius: "2px",
                        position: "absolute",
                        zIndex: 100,
                        top: `${top}px`,
                        left: "0%",
                        height: `${height}px`,
                        width: "100%",
                        pointerEvents: "auto",
                        marginTop: 1,
                        marginLeft: 2,
                        overflow: "hidden",
                      };
                      return (
                        <div key={e.id} style={styles.appointmentContainer}>
                          <Appointment
                            style={inlineStyle}
                            appointment={e}
                            // context={context}
                            // graphApiService={graphApiService}
                          />
                        </div>
                      );
                    }
                    return null;
                  })}
                </div>
                {selectedDate.getDate() === currentDay.getDate() && (
                  <div
                    className="time-line active"
                    style={{ ...styles.timeLineActive,top: `${calculateTimelinePosition()}px`,}}
                  >
                    <div className="before-dot" style={styles.beforeDot} />
                    <div className="after-dot" style={styles.afterDot} />
                  </div>
                )}
              </div>
            </div>
          </div>
        </div>
      </div>
      </>
      ) : (
        <Stack
          horizontalAlign="center"
          verticalAlign="center"
          style={{ width: "100%", minHeight: "inherit" }}
        >
          <Loading className="none" description="One moment. Loading your calendar..."  />
        </Stack>
      )}
    </>
  );
}
