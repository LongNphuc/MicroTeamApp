import React, {
  ReactNode,
  useCallback,
  useEffect,
  useMemo,
  useState,
} from "react";
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

export function ScheduleWeek({ }: DateRangeProps) {
  const { scheduleData, pageRenderState, scheduleTypeState } = useCalendar();
  const [isRender, setIsRender] = pageRenderState;
  const [scheduleTypeSelected] = scheduleTypeState;

  const { getDayOfWeek, generateTimeSlots } = useDateUtils();

  const [schedule, setSchedule] = useState<AppointmentModel[]>([]);

  const currentDate = new Date();
  const currentDay = new Date(currentDate.setHours(0, 0, 0, 0));

  const generateDateRangeArray = (startDate: Date, endDate: Date): Date[][] => {
    const dateArray: Date[][] = [];
    const currentDate = new Date(startDate.toString());

    while (currentDate <= endDate) {
      const baseDate = new Date(currentDate.toString());
      const timeSlots = generateTimeSlots(baseDate).map(
        (date) => new Date(date.getTime())
      );
      dateArray.push(timeSlots);
      currentDate.setDate(currentDate.getDate() + 1);
      currentDate.setHours(0, 0, 0, 0);
    }
    if (dateArray.length > 0) {
      const lastDayTimeSlots = generateTimeSlots(new Date(endDate.toString()));
      dateArray[dateArray.length - 1] = lastDayTimeSlots.map(
        (date) => new Date(date.getTime())
      );
    }

    return dateArray;
  };

  const dateRangeArray = generateDateRangeArray(
    new Date(scheduleData.startDate.toString()),
    new Date(scheduleData.endDate.toString())
  );

  const times: string[] = [];
  dateRangeArray[0].forEach((e, i) => {
    const minutes = e.getMinutes();
    if (minutes === 30) {
      times.push("");
    } else {
      if (e.getHours() === 0) {
        times.push("12AM");
      } else if (e.getHours() === 12) {
        times.push("12PM");
      } else if (e.getHours() >= 13) {
        times.push(e.getHours() - 12 + "PM");
      } else {
        times.push(e.getHours() + "AM");
      }
    }
  });

  const genSchedule = useCallback(async () => {
    const currentTimezone = moment.tz.guess();
    const startDate = moment
      .tz(scheduleData.startDate, currentTimezone)
      .utc()
      .format();
    const endDate = moment
      .tz(scheduleData.endDate, currentTimezone)
      .utc()
      .format();
    const utcDateRange: UTCDateRange = { start: startDate, end: endDate };
    // const result = await graphApiService.getCalendarEvent(
    //   context,
    //   utcDateRange
    // );
    // setSchedule(result);
  }, [scheduleData]);

  function areDivsOverlapping(div1: HTMLDivElement, div2: HTMLDivElement) {
    const rect1 = div1.getBoundingClientRect();
    const rect2 = div2.getBoundingClientRect();

    const isOverlapping = !(
      rect1.right <= rect2.left ||
      rect1.left >= rect2.right ||
      rect1.bottom <= rect2.top ||
      rect1.top >= rect2.bottom
    );

    return isOverlapping;
  }

  function adjustDivs(parentNodes: any) {
    const innerDivs = parentNodes.childNodes;
    const overlappingDivs = [];

    for (let i = 0; i < innerDivs.length; i++) {
      for (let j = i + 1; j < innerDivs.length; j++) {
        if (areDivsOverlapping(innerDivs[i], innerDivs[j])) {
          overlappingDivs.push(innerDivs[i], innerDivs[j]);
        }
      }
    }

    const uniqueOverlappingDivs = new Set(overlappingDivs);
    const uniqueOverlappingDivsArray: HTMLDivElement[] = [];
    uniqueOverlappingDivs.forEach((e) => {
      uniqueOverlappingDivsArray.push(e);
    });

    if (uniqueOverlappingDivsArray.length > 0) {
      const numDivs = uniqueOverlappingDivsArray.length;
      const divWidth = 100 / numDivs;

      uniqueOverlappingDivsArray.forEach(
        (div: HTMLDivElement, index: number) => {
          div.style.width = divWidth - 1 + "%";
          div.style.left = divWidth * index + "%";
        }
      );
    }
  }

  const containAppointment: ReactNode[][] = [];
  const holdManyday: ReactNode[] = [];
  for (let i = 0; i < 7; i++) {
    containAppointment.push([]);
  }
  schedule.forEach((e: AppointmentModel) => {
    console.log(e);
    const startDate = new Date(
      moment.utc(e.start.dateTime).local().format("YYYY-MM-DD HH:mm:ss")
    );
    const endDate = new Date(
      moment.utc(e.end.dateTime).local().format("YYYY-MM-DD HH:mm:ss")
    );
    if (
      startDate.getTime() - dateRangeArray[0][0].getTime() < 0 &&
      endDate.getTime() - dateRangeArray[0][0].getTime() > 0
    ) {
      const position =
        (2 * (endDate.getTime() - dateRangeArray[0][0].getTime())) /
        1000 /
        60 /
        60;
      const width =
        (100 / dateRangeArray.length) * (Math.floor(position / 48) + 1);
      const inlineStyle = {
        padding: "0 4px",
        boxSizing: "border-box",
        backgroundColor: "#DCE0FA",
        borderLeft: "4px solid #5B5FC7",
        borderRadius: "2px",
        position: "absolute",
        zIndex: 100,
        left: `0%`,
        height: "100%",
        width: `${width}%`,
        pointerEvents: "auto",
        marginTop: 1,
        marginLeft: 2,
        overflow: "hidden",
      };
      holdManyday.push(
        <div style={styles.appointmentContainer}>
          <Appointment
            style={inlineStyle}
            appointment={e}

          />
        </div>
      );
    } else {
      for (let i = 0; i < dateRangeArray.length; i++) {
        for (let j = 0; j <= 47; j++) {
          const tmp = startDate.getTime() - dateRangeArray[i][j].getTime();
          if (tmp >= 0 && tmp < 1800000) {
            const position =
              (2 * (endDate.getTime() - startDate.getTime())) / 1000 / 60 / 60;
            if (position >= 48) {
              const left = (100 / 7) * i;
              const width =
                (100 / 7) * (Math.floor(position / 48) + 1);
              const inlineStyle = {
                padding: "0 4px",
                boxSizing: "border-box",
                backgroundColor: "#DCE0FA",
                borderLeft: "4px solid #5B5FC7",
                borderRadius: "2px",
                position: "absolute",
                zIndex: 100,
                left: `${left}%`,
                height: "100%",
                width: `${width}%`,
                pointerEvents: "auto",
                marginTop: 1,
                marginLeft: 2,
                overflow: "hidden",
              };
              holdManyday.push(
                <div style={styles.appointmentContainer}>
                  <Appointment
                    style={inlineStyle}
                    appointment={e}
     
                  />
                </div>
              );
            } else {
              const top = 41 * (j + tmp / 1800000);
              const height = 41 * position - 3;
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
                height: height,
                width: "100%",
                pointerEvents: "auto",
                marginTop: 1,
                marginLeft: 2,
                overflow: "hidden",
              };
              containAppointment[i].push(
                <Appointment
                  style={inlineStyle}
                  appointment={e}
             
                />
              );
              if (position + j - 48 > 0) {
                const height = 41 * (position + (j + tmp / 1800000) - 48) - 3;
                const inlineStyle = {
                  padding: "0 4px",
                  boxSizing: "border-box",
                  backgroundColor: "#DCE0FA",
                  borderLeft: "4px solid #5B5FC7",
                  borderRadius: "2px",
                  position: "absolute",
                  zIndex: 100,
                  top: `0px`,
                  left: "0%",
                  height: height,
                  width: "100%",
                  pointerEvents: "auto",
                  marginTop: 1,
                  marginLeft: 2,
                  overflow: "hidden",
                };
                containAppointment[i + 1].push(
                  <Appointment
                    style={inlineStyle}
                    appointment={e}
                
                  />
                );
              }
            }
          }
        }
      }
    }
  });
  const checkCurrentWeek = useCallback(
    (dateRangeArray: Date[][], currentDate: string) => {
      let isCheck: boolean = false;
      dateRangeArray.forEach((d) => {
        if (d[0].toString() === currentDate) {
          isCheck = true;
        }
      });
      return isCheck;
    },
    []
  );

  useEffect(() => {
    const fetchData = async () => {
      try {
        await genSchedule();
        if (isRender) setIsRender(false);
      } catch (error) {
        console.error("Error occurred while fetching schedule: ", error);
      }
    };
  
    fetchData();
  }, [scheduleData, isRender, scheduleTypeSelected]);

  useEffect(() => {
    const containAppointment = document.querySelectorAll(
      ".contain-appointment"
    );
    containAppointment.forEach((e) => {
      adjustDivs(e);
    });

    const holdManydays: HTMLDivElement | null =
      document.querySelector(".hold-many-days");
    const cellBlank: HTMLDivElement | null =
      document.querySelector(".cell-blank");

    if (holdManydays && cellBlank) {
      if (holdManydays.childNodes.length === 0) {
        holdManydays.style.border = "none";
        cellBlank.style.border = "none";
      } else {
        cellBlank.style.borderTop = "1px solid #e0e0e0";
        cellBlank.style.marginTop = "-1px";
        holdManydays.style.borderLeft = "1px solid #e0e0e0";
        holdManydays.style.borderBottom = "1px solid #e0e0e0";
      }
    }

    const calendarBody: Element | null= document.querySelector(".calendar-body");
    const timeLines: NodeListOf<HTMLElement> = document.querySelectorAll(".time-line");
    const timeLineActive: HTMLElement | null = document.querySelector(".time-line.active");
    let currentHour = new Date().getHours();
    let currentMinutes = new Date().getMinutes();
    let currentPosition = (currentHour / 24) * ((calendarBody) ? calendarBody.clientHeight : 1 )+ (currentMinutes / 60) * 82;

    if (timeLines) {
      timeLines.forEach((element: HTMLElement) => {
        element.style.top = `${currentPosition}px`;
      });
    }

    if (timeLineActive) {
      const postion = timeLineActive.getBoundingClientRect();
      const body = document.body;
      const offSetY = postion.top - body.getBoundingClientRect().top;
      window.scrollTo({ top: offSetY, left: 0 });
    }

    const intervalId = setInterval(() => {
      currentHour = new Date().getHours();
      currentMinutes = new Date().getMinutes();
      currentPosition = (currentHour / 24) * ((calendarBody) ? calendarBody.clientHeight : 1 ) + (currentMinutes / 60) * 82;
      if (timeLines) {
        timeLines.forEach((element: HTMLElement) => {
          element.style.top = `${currentPosition}px`;
        });
      }
    }, 60000);

    return () => clearInterval(intervalId);
  }, containAppointment);

  const getDay = (date: Date) => {
    return getDayOfWeek(date);
  }

  
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
              }}
            >
              {dateRangeArray.map((dateRange: Date[], i: number) => {
                const date = dateRange[0];
                const dayOfMonth = date.getDate();
                const dayOfWeek =getDay(date);

                return (
                  <Text
                    key={i}
                    variant="xLarge"
                    style={{
                      boxSizing: "border-box",
                      width: `calc(100% / ${dateRangeArray.length})`,
                      borderLeft: "1px solid #e0e0e0",
                      borderBottom: "1px solid #e0e0e0",
                      borderTop:
                        currentDay.toString() === date.toString()
                          ? "2px solid #6264a7"
                          : "none",
                      paddingBottom: 6,
                    }}
                  >
                    <Text
                      styles={{
                        root: {
                          fontSize: 24,
                          lineHeight: "32px",
                          display: "block",
                          color:
                            currentDay.toString() === date.toString()
                              ? "#6264A7"
                              : "#616161",
                          fontWeight:
                            currentDay.toString() === date.toString()
                              ? "600"
                              : "400",
                          marginLeft: 8,
                        },
                      }}
                    >
                      {dayOfMonth}
                    </Text>
                    <Text
                      styles={{
                        root: {
                          fontSize: 12,
                          lineHeight: "16px",
                          color:
                            currentDay.toString() === date.toString()
                              ? "#6264A7"
                              : "#616161",
                          fontWeight:
                            currentDay.toString() === date.toString()
                              ? "600"
                              : "400",
                          marginLeft: 8,
                        },
                      }}
                    >
                      {dayOfWeek}
                    </Text>
                  </Text>
                );
              })}
            </div>
          </div>
          <div style={{ display: "flex" }}>
            <div className="cell-blank" style={styles.cellBlank} />
            <div className="hold-many-days" style={styles.holdManyDays}>
              {holdManyday}
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
                  }}
                >
                  {times.map((time: string, i: number) => (
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
                <div
                  className="calendar-cell-list"
                  style={{ display: "flex", boxSizing: "border-box", flex: 1 }}
                >
                  {dateRangeArray.map((e, i: number) => (
                    <div
                      key={i}
                      className="calendar-col"
                      style={{
                        width: `calc(100% / ${dateRangeArray.length})`,
                        position: "relative",
                      }}
                    >
                      {e.map((cn, index) => (
                        <div
                          key={index}
                          className="calendar-cell"
                          title={cn.toString()}
                          style={{
                            height: 41,
                            width: "100%",
                            borderLeft: "1px solid #e0e0e0",
                            borderBottom:
                              cn.getMinutes() === 0
                                ? "1px dashed #e0e0e0"
                                : "1px solid #e0e0e0",
                            boxSizing: "border-box",
                          }}
                        />
                      ))}
                      <div
                        className="contain-appointment"
                        style={styles.containAppointment}
                      >
                        {containAppointment[i].map((e: ReactNode) => e)}
                      </div>
                      {dateRangeArray[i][0].getTime() < currentDay.getTime() &&
                        checkCurrentWeek(
                          dateRangeArray,
                          currentDay.toString()
                        ) && (
                          <div className="time-line" style={styles.timeLine} />
                        )}
                      {dateRangeArray[i][0].getTime() ===
                        currentDay.getTime() && (
                        <div
                          className="time-line active"
                          style={styles.timeLineActive}
                        >
                          <div
                            className="before-dot"
                            style={styles.beforeDot}
                          />
                          <div className="after-dot" style={styles.afterDot} />
                        </div>
                      )}
                    </div>
                  ))}
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
          <Loading
            className="none"
            description="One moment. Loading your calendar..."
          />
        </Stack>
      )}
    </>
  );
}
