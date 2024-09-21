import React, { useState, useCallback, useMemo } from 'react';
import {Event} from '../../models/Event'
import {
    Stack,
    SearchBox,
    Dropdown,
    IDropdownOption,
    PrimaryButton,
    DefaultButton,
    DialogFooter,
    TextField,
    DatePicker,
    TimePicker,
    ITimeRange,
    TimePickerValidationResultData,
    Text,
    Modal,
    Toggle,
    IComboBox,
} from '@fluentui/react';
import {
    FilterRegular,
    ArrowLeftRegular,
    ArrowRightRegular,
    CalendarRegular,
    ChevronDownRegular,
    CalendarDayRegular,
    CalendarWorkWeekRegular,
    CalendarAgendaRegular,
    CalendarWeekNumbersRegular,
    CheckmarkRegular
  } from "@fluentui/react-icons";
import { useBoolean } from '@fluentui/react-hooks';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IGraphApiService } from "../../services/IGraphApiService";
import {
    Calendar as Calendars,
    DateRangeType,
    DayOfWeek,
} from "@fluentui/react-calendar-compat";
import {
    Button,
    Menu,
    MenuTrigger,
    MenuPopover,
    MenuProps,
    MenuList,
    MenuItem,
    MenuDivider,
    Checkbox,
    Field,
    Label,
    CheckboxOnChangeData,
} from "@fluentui/react-components";
import moment from 'moment';
import { useDateUtils } from '../../hooks/useDateUtils';
import { buttonStyles, contentStyles, dropdownStyles, StyledImage, styles } from '../../styles/Calendar/NavStyles';
import { useUsers } from '../../hooks/useUsers';
import { ScheduleType } from '../../utils/EnumConst';
import { useCalendar } from '../../contexts/calendar/CalendarContext';

const timeRange: ITimeRange = {
    start: 0,
    end: 24,
};

const options: IDropdownOption[] = [
    { key: 'appointment', text: 'Client Appointment' },
    { key: 'event', text: 'Event' },
    { key: 'office', text: 'Out of office' },
];

const clients: IDropdownOption[] = [
    { key: 'a', text: 'Joe Smith', data: { imageSrc: 'https://via.placeholder.com/50' } },
    { key: 'b', text: 'John Woods', data: { imageSrc: 'https://via.placeholder.com/50' } },
    { key: 'c', text: 'Sue Smith', data: { imageSrc: 'https://via.placeholder.com/50' } },
    { key: 'd', text: 'Adam Gray', data: { imageSrc: 'https://via.placeholder.com/50' } },
    { key: 'f', text: 'Eleanor Rigby', data: { imageSrc: 'https://via.placeholder.com/50' } },
];

const onRenderOption = (option?: IDropdownOption): JSX.Element => {
    if(!option) return <></>;
    return (
        <div style={{ display: 'flex', alignItems: 'center' }}>
            {option.data && option.data.imageSrc && (
                <StyledImage src={option.data.imageSrc} />
            )}
            <span>{option.text}</span>
        </div>
    );
};


const onRenderTitle = (options?: IDropdownOption[]): JSX.Element => {
    if(!options) return <></>;
    const option = options[0];

    return (
        <div style={{ display: 'flex', alignItems: 'center' }}>
            {option.data && option.data.imageSrc && (
                <StyledImage src={option.data.imageSrc} />
            )}
            <span>{option.text}</span>
        </div>
    );
};


interface INavProps {

}

const dateRangeType = DateRangeType.Week;
const firstDayOfWeek = DayOfWeek.Sunday;
export function Nav({}: INavProps, props: MenuProps,) {
    const { scheduleTypeState, setDateRange, pageRenderState} = useCalendar();
    const [isRender, setIsRender] = pageRenderState;
    const [scheduleTypeSelected, setScheduleTypeSelected] = scheduleTypeState;
    const now = new Date();
    const {getWeekDays, utcDateTime} = useDateUtils();
    const [startTime, setStartTime] = useState<Date>(new Date());
    const [endTime, setEndTime] = useState<Date>(new Date(now.getTime() + 30 * 60000));
    const [disableButton, setDisableButton] = useState<boolean>(false);
    const [isModalOpen, { setTrue: showModal, setFalse: hideModal }] = useBoolean(false);
    const [selectedDateRange, setSelectedDateRange] = useState<Date[]>();
    const [selectedDate, setSelectedDate] = useState<Date>(new Date());
    const [selectedAppointmentDate, setSelectedAppointmentDate] = useState<Date>(new Date());
    const [clinicianName, setClinicianName] = React.useState('');
    const [note, setNote] = useState('');
    const [checkedItems, setCheckedItems] = useState({
        open: false,
        close: false,
        telehealth: false,
        manager: false,
    });
    const [inputLocation, setInputLocation] = React.useState('');
    const [isToggled, setIsToggled] = React.useState(true);
    const [selectedClient, setSelectedClient] = React.useState<IDropdownOption>();
    const [selectedAppointmentType, setSelectedAppointmentType] = React.useState<IDropdownOption>();
    const [isMenuOpen, setIsMenuOpen] = useState(false);

    const handleChangeScheduleType = (event: React.MouseEvent<HTMLElement>) => {
        const scheduleType = event.currentTarget.getAttribute("data-schedule-type");
        if(scheduleType === ScheduleType.Day.toString() || scheduleType === ScheduleType.Agenda.toString()) {
          setDateRange([selectedDate]);
          setScheduleTypeSelected((scheduleType === ScheduleType.Day.toString()) ? ScheduleType.Day : ScheduleType.Agenda);
        }else {
          const weekDays = getWeekDays(selectedDate, 0);
          setSelectedDateRange(weekDays);
          setDateRange(weekDays);
          setScheduleTypeSelected((scheduleType === ScheduleType.Week.toString()) ? ScheduleType.Week : ScheduleType.WorkWeek);
        }
    };
    
    const onAppointmentTypeChange = (event: React.FormEvent<HTMLDivElement>, item?: IDropdownOption): void => {
        setSelectedAppointmentType(item);
    };
    const onClientChange = (event: React.FormEvent<HTMLDivElement>, item?: IDropdownOption): void => {
        setSelectedClient(item);
    };
    
    const onSelectDate = useCallback((date: Date, selectedDateRangeArray?: Date[]): void => {
        setSelectedDate(date);
        if (scheduleTypeSelected === ScheduleType.Day || scheduleTypeSelected === ScheduleType.Agenda) {
            setDateRange([date]);
        } else {
            if (selectedDateRangeArray) {
                setDateRange(selectedDateRangeArray);
            }
        }
        setSelectedDateRange(selectedDateRangeArray);
        setIsMenuOpen(false);
    },[selectedDate, selectedDateRange, scheduleTypeSelected]);

    const onStartTimeChange = React.useCallback((event: React.FormEvent<IComboBox>, newTime: Date) => {
        setStartTime(newTime);
    }, []);
    

    const onEndTimeChange = React.useCallback((event: React.FormEvent<IComboBox>, newTime: Date) => {
        if (moment(newTime.getTime()).isAfter(moment(startTime.getTime()))) {
            setEndTime(newTime);
        }
    }, [startTime]);

    const goPrevious = useCallback((): void => {
        const currentDate = selectedDate ?? new Date();
    
        if (scheduleTypeSelected === ScheduleType.Day || scheduleTypeSelected === ScheduleType.Agenda) {
          const previousDay = new Date(
            currentDate.setDate(currentDate.getDate() - 1)
          );
          setSelectedDate(previousDay);
          setSelectedDateRange([previousDay]);
          setDateRange([previousDay]);
        } else {
          const preWeekDays = getWeekDays(new Date(currentDate.setDate(currentDate.getDate() - 7)), 0);
          setSelectedDate(preWeekDays[0]);
          setSelectedDateRange(preWeekDays);
          setDateRange(preWeekDays);
        }
    }, [selectedDate, selectedDateRange, scheduleTypeSelected]);

    const goNext = useCallback((): void => {
        const currentDate = selectedDate ?? new Date();
        if (scheduleTypeSelected === ScheduleType.Day || scheduleTypeSelected === ScheduleType.Agenda) {
          const nextDay = new Date(currentDate.setDate(currentDate.getDate() + 1));
          setSelectedDate(nextDay);
          setSelectedDateRange([nextDay]);
          setDateRange([nextDay]);
        } else {
          const nextWeekDays = getWeekDays(new Date(currentDate.setDate(currentDate.getDate() + 7)),0);
          setSelectedDate(nextWeekDays[0]);
          setSelectedDateRange(nextWeekDays);
          setDateRange(nextWeekDays);
        }
    }, [selectedDate, selectedDateRange, scheduleTypeSelected]);

    const showGoToToday = useCallback((): void => {
        const today = new Date();
    
        if (scheduleTypeSelected === ScheduleType.Day || scheduleTypeSelected === ScheduleType.Agenda) {
          setSelectedDate(today);
          setSelectedDateRange([today]);
          setDateRange([today]);
        } else {
          const weekDays = getWeekDays(today, 0);
          setSelectedDate(today);
          setSelectedDateRange(weekDays);
          setDateRange(weekDays);
        }
    }, [selectedDate, selectedDateRange, scheduleTypeSelected]);

    const startUtcDateTime = useMemo(
        () => utcDateTime(selectedAppointmentDate, startTime),
        [selectedAppointmentDate, startTime]
    );
    
    const endUtcDateTime = useMemo(
        () => utcDateTime(selectedAppointmentDate, endTime),
        [selectedAppointmentDate, endTime]
    ); 

    const onChange = (ev: React.ChangeEvent<HTMLInputElement>, data: CheckboxOnChangeData) => {
        const { checked } = data;
        setCheckedItems({ ...checkedItems, [ev.currentTarget.name]: checked });
    };
    const clearFilters = () => {
        setCheckedItems({
            open: false,
            close: false,
            telehealth: false,
            manager: false,
        });
    };

    const onValidationResult = useCallback((event: React.FormEvent<IComboBox>, timePickerValidationResultData: TimePickerValidationResultData) => {
        if (timePickerValidationResultData.errorMessage !== undefined) {
            console.log('Validation error message received: ', timePickerValidationResultData.errorMessage);
            setDisableButton(timePickerValidationResultData.errorMessage.length > 0);
        }
    }, []);

    const onParseDateFromString = useCallback(
        (newValue: string): Date => {
            const previousValue = selectedAppointmentDate ?? new Date();
            const newValueParts = (newValue || '').trim().split('/');
            const day =
                newValueParts.length > 0 ? Math.max(1, Math.min(31, parseInt(newValueParts[0], 10))) : previousValue.getDate();
            const month =
                newValueParts.length > 1
                    ? Math.max(1, Math.min(12, parseInt(newValueParts[1], 10))) - 1
                    : previousValue.getMonth();
            let year = newValueParts.length > 2 ? parseInt(newValueParts[2], 10) : previousValue.getFullYear();
            if (year < 100) {
                year += previousValue.getFullYear() - (previousValue.getFullYear() % 100);
            }
            return new Date(year, month, day);
        },
        [selectedAppointmentDate],
    );
    const onFormatDate = (date?: Date): string => {
        return !date ? '' : date.getDate() + '/' + (date.getMonth() + 1) + '/' + (date.getFullYear() % 100);
    }

    const onInputChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
        setInputLocation(newValue ?? '');
    };
    
    const onNoteChange = (event: any, newValue: any) => {
        setNote(newValue ?? '');
    };

    const onToggleChange = (event: React.FormEvent<HTMLElement>, checked?: boolean) => {
        setIsToggled(checked || false);
    };

    const onChangeClinicianName = React.useCallback(
        (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
          setClinicianName(newValue || '');
        },
        [],
    );

    const handleSelectDate = (date: Date | null | undefined) => {
        if (date) {
            setSelectedAppointmentDate(date);
        }
    };

    const formattedDuration = useMemo(() => {
        const startMoment = moment(startTime);
        const endMoment = moment(endTime);
        const duration = moment.duration(endMoment.diff(startMoment));
        return [
            duration.hours() > 0 && `${duration.hours()}h`,
            duration.minutes() > 0 && `${duration.minutes()}m`
        ].filter(Boolean).join(' ');
    }, [startTime, endTime]);

    const getIconForScheduleType = useMemo(() => {
        switch (scheduleTypeSelected) {
          case ScheduleType.Day:
            return <CalendarDayRegular />;
          case ScheduleType.WorkWeek:
            return <CalendarWorkWeekRegular />;
          case ScheduleType.Week:
            return <CalendarWeekNumbersRegular />;
          case ScheduleType.Agenda:
            return <CalendarAgendaRegular />;
          default:
            return null;
        }
    },[scheduleTypeSelected]);

    const handleSubmit = useCallback(async () => {
        const newEvent: Event = {
            subject: "Appointment",
            start: {
                dateTime: startUtcDateTime,
                timeZone: "UTC"
            },
            end: {
                dateTime: endUtcDateTime,
                timeZone: "UTC"
            },
             isOnlineMeeting: isToggled ? true : false,
             allowNewTimeProposals: true,
             onlineMeetingProvider: isToggled ? "teamsForBusiness" : "unknown",
            // attendees: [
            //     {
            //         emailAddress: {
            //             address: "john.doe@example.com",
            //             name: "John Doe"
            //         },
            //         type: "required"
            //     },
            //     {
            //         emailAddress: {
            //             address: "jane.smith@example.com",
            //             name: "Jane Smith"
            //         },
            //         type: "optional"
            //     }
            // ],
            
            location: {
                displayName: inputLocation
            },
            extensions: [
                {
                    "@odata.type": "microsoft.graph.openTypeExtension",
                    extensionName: "Com.Contoso.Referral",
                    appointmentType: selectedAppointmentType?.text ?? '',
                    clinicianName: clinicianName,
                    client: selectedClient?.text ?? '',
                    note :note
                }
            ]
        }
        // await graphApiService.createCalendarEvent(context,newEvent);
        hideModal();
        setIsRender(true);
    }, [selectedAppointmentDate,startTime,endTime,clinicianName,inputLocation,isToggled,note,selectedAppointmentType,selectedClient])

    const toggleMenu = () => {
        setIsMenuOpen((prevState) => !prevState);
    };

    const handleCloseModal = () => {
        setSelectedAppointmentType(undefined);
        setSelectedClient(undefined);
        setClinicianName("");
        setSelectedAppointmentDate(new Date());
        setStartTime(new Date());
        setEndTime(new Date(now.getTime() + 30 * 60000));
        setInputLocation("");
        setIsToggled(true);
        setNote("");
        hideModal();
    };
    
    return (
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center" styles={{ root: { padding: '10px', borderBottom: '1px solid #e0e0e0' } }}>
            <Stack horizontal verticalAlign="center">
                <Button onClick={showGoToToday} appearance="transparent" icon={<CalendarRegular />}>
                    Today
                </Button>
                <Button onClick={goPrevious} size="large" appearance="transparent" icon={<ArrowLeftRegular />}/>
                
                <Button onClick={goNext} size="large" appearance="transparent" icon={<ArrowRightRegular />} />
                <Menu open={isMenuOpen} {...props}>
                    {
                        [
                            <MenuTrigger disableButtonEnhancement key="trigger">
                                <Button appearance="transparent" onClick={toggleMenu} iconPosition="after" icon={<ChevronDownRegular />}>{moment(selectedDate).format("MMMM YYYY")}</Button>
                            </MenuTrigger>,
                            <MenuPopover tabIndex={-1} key="surface">
                                <Calendars
                                    showMonthPickerAsOverlay
                                    dateRangeType={dateRangeType}
                                    highlightSelectedMonth
                                    onSelectDate={onSelectDate}
                                    value={selectedDate}
                                    firstDayOfWeek={firstDayOfWeek}
                                    showGoToToday={false}
                                />
                            </MenuPopover>
                        ] as [React.ReactElement, React.ReactElement]
                    }
                </Menu>
            </Stack>
            <Stack horizontal verticalAlign="center">
                <Menu {...props} persistOnItemClick={true} >
                    {
                        [
                            <MenuTrigger disableButtonEnhancement key="trigger">
                                <Button appearance="transparent" icon={<FilterRegular />}>
                                    Filter
                                </Button>
                            </MenuTrigger>,
                            <MenuPopover tabIndex={-1} key="surface" style={{ width: '220px' }}>
                                <Stack tokens={{ childrenGap: 10 }} >
                                    <MenuList >
                                        <Stack horizontal horizontalAlign="space-between">
                                            <Label weight="semibold" style={styles.filter}>Filter</Label>
                                            <Button appearance="transparent" onClick={clearFilters} size="small" >
                                                Clear
                                            </Button>
                                        </Stack>
                                        <MenuDivider />
                                        <MenuItem>
                                            <Checkbox
                                                label="Open"
                                                name="open"
                                                checked={checkedItems.open}
                                                onChange={onChange}
                                                labelPosition="before"
                                                style={styles.checkbox}
                                            />
                                        </MenuItem>
                                        <MenuItem>
                                            <Checkbox
                                                label="Close"
                                                name="close"
                                                checked={checkedItems.close}
                                                onChange={onChange}
                                                labelPosition="before"
                                                style={styles.checkbox}
                                            />
                                        </MenuItem>
                                        <MenuItem>
                                            <Checkbox
                                                label="Telehealth"
                                                name="telehealth"
                                                checked={checkedItems.telehealth}
                                                onChange={onChange}
                                                labelPosition="before"
                                                style={styles.checkbox}
                                            />
                                        </MenuItem>
                                        <MenuItem>
                                            <Checkbox
                                                label="Manager"
                                                name="manager"
                                                checked={checkedItems.manager}
                                                onChange={onChange}
                                                labelPosition="before"
                                                style={styles.checkbox}
                                            />
                                        </MenuItem>
                                    </MenuList>
                                </Stack>
                            </MenuPopover>
                        ] as [React.ReactElement, React.ReactElement]
                    }
                </Menu>
                <Field >
                    <SearchBox placeholder="Search" />
                </Field>
                <Menu>
                    <MenuTrigger disableButtonEnhancement>
                        <Button
                        appearance="transparent"
                        icon={<ChevronDownRegular />}
                        iconPosition="after"
                        >
                        <span style={{ display: 'flex', alignItems: 'center' }}>
                            {getIconForScheduleType}
                            <span style={{ marginLeft: '8px' }}>{scheduleTypeSelected}</span>
                        </span>
                        </Button>
                    </MenuTrigger>
                    <MenuPopover>
                        <MenuList>
                        <MenuItem
                            data-schedule-type={ScheduleType.Day}
                            onClick={handleChangeScheduleType}
                            aria-selected={scheduleTypeSelected === ScheduleType.Day}
                            // className={menuItemScheduleStyle}
                        >
                            <span style={{ display: 'flex', alignItems: 'center', width: '100%' }}>
                            <CalendarDayRegular />
                            <span style={{ marginLeft: '8px' }}>Day</span>
                            {scheduleTypeSelected === ScheduleType.Day && (
                                <span style={{ marginLeft: 'auto' }}>
                                <CheckmarkRegular />
                                </span>
                            )}
                            </span>
                        </MenuItem>

                        <MenuItem
                            data-schedule-type={ScheduleType.WorkWeek}
                            onClick={handleChangeScheduleType}
                            aria-selected={scheduleTypeSelected === ScheduleType.WorkWeek}
                            // className={menuItemScheduleStyle}
                        >
                            <span style={{ display: 'flex', alignItems: 'center', width: '100%' }}>
                            <CalendarWorkWeekRegular />
                            <span style={{ marginLeft: '8px' }}>Work Week</span>
                            {scheduleTypeSelected === ScheduleType.WorkWeek && (
                                <span style={{ marginLeft: 'auto' }}>
                                <CheckmarkRegular />
                                </span>
                            )}
                            </span>
                        </MenuItem>

                        <MenuItem
                            data-schedule-type={ScheduleType.Week}
                            onClick={handleChangeScheduleType}
                            aria-selected={scheduleTypeSelected === ScheduleType.Week}
                            // className={menuItemScheduleStyle}
                        >
                            <span style={{ display: 'flex', alignItems: 'center', width: '100%' }}>
                            <CalendarWeekNumbersRegular />
                            <span style={{ marginLeft: '8px' }}>Week</span>
                            {scheduleTypeSelected === ScheduleType.Week && (
                                <span style={{ marginLeft: 'auto' }}>
                                <CheckmarkRegular />
                                </span>
                            )}
                            </span>
                        </MenuItem>

                        <MenuItem
                            data-schedule-type={ScheduleType.Agenda}
                            onClick={handleChangeScheduleType}
                            aria-selected={scheduleTypeSelected === ScheduleType.Agenda}
                            // className={menuItemScheduleStyle}
                        >
                            <span style={{ display: 'flex', alignItems: 'center', width: '100%' }}>
                            <CalendarAgendaRegular />
                            <span style={{ marginLeft: '8px' }}>Agenda</span>
                            {scheduleTypeSelected === ScheduleType.Agenda && (
                                <span style={{ marginLeft: 'auto' }}>
                                <CheckmarkRegular />
                                </span>
                            )}
                            </span>
                        </MenuItem>
                        </MenuList>
                    </MenuPopover>
                </Menu>

                <PrimaryButton
                    text="Add New"
                    onClick={showModal}
                    styles={buttonStyles}
                    iconProps={{
                        iconName: 'Add',
                        styles: { root: { marginRight: 8 } }
                    }}
                />
                <Modal
                    isOpen={isModalOpen}
                    onDismiss={handleCloseModal}
                    isBlocking={true}
                    containerClassName={contentStyles.container}
                >
                    <div className={contentStyles.header}>
                        <Text className={contentStyles.heading} variant='large'>
                            Add a new appointment
                        </Text>
                    </div>
                    <div className={contentStyles.body}>
                        <form>
                            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 16, padding: 10 }}>
                                <svg width="14" height="14" viewBox="0 0 14 14" fill="none" xmlns="http://www.w3.org/2000/svg">
                                    <path d="M4 8C4.55228 8 5 7.55229 5 7C5 6.44771 4.55228 6 4 6C3.44772 6 3 6.44771 3 7C3 7.55229 3.44772 8 4 8ZM5 10C5 10.5523 4.55228 11 4 11C3.44772 11 3 10.5523 3 10C3 9.44771 3.44772 9 4 9C4.55228 9 5 9.44771 5 10ZM7 8C7.55229 8 8 7.55229 8 7C8 6.44771 7.55229 6 7 6C6.44771 6 6 6.44771 6 7C6 7.55229 6.44771 8 7 8ZM8 10C8 10.5523 7.55229 11 7 11C6.44771 11 6 10.5523 6 10C6 9.44771 6.44771 9 7 9C7.55229 9 8 9.44771 8 10ZM10 8C10.5523 8 11 7.55229 11 7C11 6.44771 10.5523 6 10 6C9.44771 6 9 6.44771 9 7C9 7.55229 9.44771 8 10 8ZM14 2.5C14 1.11929 12.8807 0 11.5 0H2.5C1.11929 0 0 1.11929 0 2.5V11.5C0 12.8807 1.11929 14 2.5 14H11.5C12.8807 14 14 12.8807 14 11.5V2.5ZM1 4H13V11.5C13 12.3284 12.3284 13 11.5 13H2.5C1.67157 13 1 12.3284 1 11.5V4ZM2.5 1H11.5C12.3284 1 13 1.67157 13 2.5V3H1V2.5C1 1.67157 1.67157 1 2.5 1Z" fill="#616161" />
                                </svg>

                                <Dropdown
                                    placeholder="Select appointment type"
                                    options={options}
                                    styles={dropdownStyles}
                                    selectedKey={selectedAppointmentType ? selectedAppointmentType.key : undefined}
                                    onChange={onAppointmentTypeChange}
                                />
                            </Stack>
                            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 16, padding: 10 }}>
                                <svg width="17" height="17" viewBox="0 0 17 17" fill="none" xmlns="http://www.w3.org/2000/svg">
                                    <path d="M7 0C4.79086 0 3 1.79086 3 4C3 6.20914 4.79086 8 7 8C9.20914 8 11 6.20914 11 4C11 1.79086 9.20914 0 7 0ZM4 4C4 2.34315 5.34315 1 7 1C8.65685 1 10 2.34315 10 4C10 5.65685 8.65685 7 7 7C5.34315 7 4 5.65685 4 4ZM2.00873 9C0.903151 9 0 9.88687 0 11C0 12.6912 0.83281 13.9663 2.13499 14.7966C3.41697 15.614 5.14526 16 7 16C7.41085 16 7.8155 15.9811 8.21047 15.9427C7.97298 15.6472 7.7654 15.3266 7.59233 14.9855C7.39798 14.9951 7.20041 15 7 15C5.26489 15 3.74318 14.636 2.67262 13.9534C1.62226 13.2837 1 12.3088 1 11C1 10.4467 1.44786 10 2.00873 10L7.59971 10C7.7826 9.64222 8.00353 9.30711 8.25716 9L2.00873 9ZM12.5 17C14.9853 17 17 14.9853 17 12.5C17 10.0147 14.9853 8 12.5 8C10.0147 8 8 10.0147 8 12.5C8 14.9853 10.0147 17 12.5 17ZM12.5 10C12.7761 10 13 10.2239 13 10.5V12H14.5C14.7761 12 15 12.2239 15 12.5C15 12.7761 14.7761 13 14.5 13H13V14.5C13 14.7761 12.7761 15 12.5 15C12.2239 15 12 14.7761 12 14.5V13H10.5C10.2239 13 10 12.7761 10 12.5C10 12.2239 10.2239 12 10.5 12H12V10.5C12 10.2239 12.2239 10 12.5 10Z" fill="#616161" />
                                </svg>

                                <Dropdown
                                    placeholder="Select a client"
                                    options={clients}
                                    styles={dropdownStyles}
                                    selectedKey={selectedClient ? selectedClient.key : undefined}
                                    onChange={onClientChange}
                                    onRenderTitle={onRenderTitle}
                                    onRenderOption={onRenderOption}
                                />
                                <PrimaryButton text="Add new client"/>
                            </Stack>
                            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 16, padding: 10 }}>
                                <svg width="17" height="17" viewBox="0 0 17 17" fill="none" xmlns="http://www.w3.org/2000/svg">
                                    <path d="M7 0C4.79086 0 3 1.79086 3 4C3 6.20914 4.79086 8 7 8C9.20914 8 11 6.20914 11 4C11 1.79086 9.20914 0 7 0ZM4 4C4 2.34315 5.34315 1 7 1C8.65685 1 10 2.34315 10 4C10 5.65685 8.65685 7 7 7C5.34315 7 4 5.65685 4 4ZM2.00873 9C0.903151 9 0 9.88687 0 11C0 12.6912 0.83281 13.9663 2.13499 14.7966C3.41697 15.614 5.14526 16 7 16C7.41085 16 7.8155 15.9811 8.21047 15.9427C7.97298 15.6472 7.7654 15.3266 7.59233 14.9855C7.39798 14.9951 7.20041 15 7 15C5.26489 15 3.74318 14.636 2.67262 13.9534C1.62226 13.2837 1 12.3088 1 11C1 10.4467 1.44786 10 2.00873 10L7.59971 10C7.7826 9.64222 8.00353 9.30711 8.25716 9L2.00873 9ZM12.5 17C14.9853 17 17 14.9853 17 12.5C17 10.0147 14.9853 8 12.5 8C10.0147 8 8 10.0147 8 12.5C8 14.9853 10.0147 17 12.5 17ZM12.5 10C12.7761 10 13 10.2239 13 10.5V12H14.5C14.7761 12 15 12.2239 15 12.5C15 12.7761 14.7761 13 14.5 13H13V14.5C13 14.7761 12.7761 15 12.5 15C12.2239 15 12 14.7761 12 14.5V13H10.5C10.2239 13 10 12.7761 10 12.5C10 12.2239 10.2239 12 10.5 12H12V10.5C12 10.2239 12.2239 10 12.5 10Z" fill="#616161" />
                                </svg>

                                <TextField
                                 value={clinicianName}
                                 onChange={onChangeClinicianName}
                                 placeholder="Clinician name" 
                                 styles={styles.textField} 
                                 />
                            </Stack>
                            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 16, padding: 10 }}>
                                <svg width="14" height="14" viewBox="0 0 14 14" fill="none" xmlns="http://www.w3.org/2000/svg">
                                    <path d="M4 8C4.55228 8 5 7.55229 5 7C5 6.44771 4.55228 6 4 6C3.44772 6 3 6.44771 3 7C3 7.55229 3.44772 8 4 8ZM5 10C5 10.5523 4.55228 11 4 11C3.44772 11 3 10.5523 3 10C3 9.44771 3.44772 9 4 9C4.55228 9 5 9.44771 5 10ZM7 8C7.55229 8 8 7.55229 8 7C8 6.44771 7.55229 6 7 6C6.44771 6 6 6.44771 6 7C6 7.55229 6.44771 8 7 8ZM8 10C8 10.5523 7.55229 11 7 11C6.44771 11 6 10.5523 6 10C6 9.44771 6.44771 9 7 9C7.55229 9 8 9.44771 8 10ZM10 8C10.5523 8 11 7.55229 11 7C11 6.44771 10.5523 6 10 6C9.44771 6 9 6.44771 9 7C9 7.55229 9.44771 8 10 8ZM14 2.5C14 1.11929 12.8807 0 11.5 0H2.5C1.11929 0 0 1.11929 0 2.5V11.5C0 12.8807 1.11929 14 2.5 14H11.5C12.8807 14 14 12.8807 14 11.5V2.5ZM1 4H13V11.5C13 12.3284 12.3284 13 11.5 13H2.5C1.67157 13 1 12.3284 1 11.5V4ZM2.5 1H11.5C12.3284 1 13 1.67157 13 2.5V3H1V2.5C1 1.67157 1.67157 1 2.5 1Z" fill="#616161" />
                                </svg>


                                <DatePicker
                                    value={selectedAppointmentDate}
                                    onSelectDate={handleSelectDate}
                                    firstDayOfWeek={firstDayOfWeek}
                                    placeholder="mm/dd/yy"
                                    ariaLabel="Select a date"
                                    showMonthPickerAsOverlay={true}
                                    parseDateFromString={onParseDateFromString}
                                    formatDate={onFormatDate}
                                    styles={{ root: { width: 100, div: { border: 'none', backgroundColor: '#F5F5F5' }, span: { backgroundColor: '#F5F5F5', border: 'none' } } }}
                                />
                            </Stack>
                            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 16, padding: 10 }}>
                                <svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
                                    <path d="M8 0C12.4183 0 16 3.58172 16 8C16 12.4183 12.4183 16 8 16C3.58172 16 0 12.4183 0 8C0 3.58172 3.58172 0 8 0ZM8 1C4.13401 1 1 4.13401 1 8C1 11.866 4.13401 15 8 15C11.866 15 15 11.866 15 8C15 4.13401 11.866 1 8 1ZM7.5 3C7.74546 3 7.94961 3.17688 7.99194 3.41012L8 3.5V8H10.5C10.7761 8 11 8.22386 11 8.5C11 8.74546 10.8231 8.94961 10.5899 8.99194L10.5 9H7.5C7.25454 9 7.05039 8.82312 7.00806 8.58988L7 8.5V3.5C7 3.22386 7.22386 3 7.5 3Z" fill="#616161" />
                                </svg>
                                <TimePicker
                                    // styles={timePickerStyles}
                                    allowFreeform
                                    useHour12
                                    increments={15}
                                    autoComplete="on"
                                    dateAnchor={startTime}
                                    value={startTime}
                                    onChange={onStartTimeChange}
                                    onValidationResult={onValidationResult}
                                    timeRange={timeRange}
                                    styles={styles.timePicker}
                                />
                                <svg width="12" height="10" viewBox="0 0 12 10" fill="none" xmlns="http://www.w3.org/2000/svg">
                                    <path d="M0.5 4.50001C0.223858 4.50001 -1.20706e-08 4.72386 0 5.00001C1.20706e-08 5.27615 0.223858 5.50001 0.5 5.50001L10.197 5.5L6.16552 9.12836C5.96026 9.31309 5.94362 9.62923 6.12835 9.83449C6.31308 10.0397 6.62923 10.0564 6.83448 9.87165L11.8345 5.37165C11.9398 5.27683 12 5.14175 12 5C12 4.85826 11.9398 4.72318 11.8345 4.62836L6.83448 0.128357C6.62923 -0.0563721 6.31308 -0.0397334 6.12835 0.165522C5.94362 0.370777 5.96026 0.686922 6.16552 0.871651L10.197 4.5L0.5 4.50001Z" fill="#616161" />
                                </svg>

                                <TimePicker
                                    // styles={timePickerStyles}
                                    allowFreeform
                                    useHour12
                                    increments={15}
                                    autoComplete="on"
                                    dateAnchor={endTime}
                                    value={endTime}
                                    onChange={onEndTimeChange}
                                    onValidationResult={onValidationResult}
                                    timeRange={timeRange}
                                    styles={styles.timePicker}
                                />
                                <Text>{formattedDuration}</Text>
                            </Stack>
                            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 16, padding: 10 }} styles={{ root: { position: 'relative' } }}>
                                <svg width="16" height="16" viewBox="0 0 14 16" fill="none" xmlns="http://www.w3.org/2000/svg">
                                    <path d="M10 7C10 8.65685 8.65685 10 7 10C5.34315 10 4 8.65685 4 7C4 5.34315 5.34315 4 7 4C8.65685 4 10 5.34315 10 7ZM9 7C9 5.89543 8.10457 5 7 5C5.89543 5 5 5.89543 5 7C5 8.10457 5.89543 9 7 9C8.10457 9 9 8.10457 9 7ZM11.9497 11.955C14.6834 9.2201 14.6834 4.78601 11.9497 2.05115C9.21608 -0.683716 4.78392 -0.683716 2.05025 2.05115C-0.683418 4.78601 -0.683418 9.2201 2.05025 11.955L3.57128 13.4538L5.61408 15.4389L5.74691 15.5567C6.52168 16.1847 7.65623 16.1455 8.38611 15.4391L10.8223 13.0691L11.9497 11.955ZM2.75499 2.75619C5.09944 0.410715 8.90055 0.410715 11.245 2.75619C13.5294 5.04153 13.5879 8.71039 11.4207 11.0667L11.245 11.2499L9.92371 12.5539L7.69315 14.7225L7.60016 14.8021C7.24594 15.0699 6.7543 15.0698 6.40012 14.802L6.30713 14.7224L3.3263 11.817L2.75499 11.2499L2.57927 11.0667C0.412077 8.71039 0.47065 5.04153 2.75499 2.75619Z" fill="#616161" />
                                </svg>
                                <TextField
                                    value={inputLocation}
                                    onChange={onInputChange}
                                    placeholder="Add location"
                                    styles={{
                                        root: {
                                            width: '100%', div: { border: 'none' }, input: {
                                                backgroundColor: '#F5F5F5', border: 'none', padding: '10px 200px 10px 10px',
                                            }
                                        }
                                    }
                                    }
                                />
                                <Stack horizontal verticalAlign="center" styles={{ root: { position: 'absolute', right: 16 } }} tokens={{ childrenGap: 10 }}>
                                    <span>Telehealth meeting</span>
                                    <Toggle
                                        checked={isToggled}
                                        onChange={onToggleChange}
                                        styles={{ root: { marginBottom: 0 } }}
                                    />
                                </Stack>
                            </Stack>
                            <Stack horizontal verticalAlign="start" tokens={{ childrenGap: 16, padding: 10 }}>
                                <svg width="14" height="12" viewBox="0 0 14 12" fill="none" xmlns="http://www.w3.org/2000/svg">
                                    <path d="M1 2.5C1.55228 2.5 2 2.05228 2 1.5C2 0.947715 1.55228 0.5 1 0.5C0.447715 0.5 0 0.947715 0 1.5C0 2.05228 0.447715 2.5 1 2.5ZM1 7C1.55228 7 2 6.55228 2 6C2 5.44772 1.55228 5 1 5C0.447715 5 0 5.44772 0 6C0 6.55228 0.447715 7 1 7ZM2 10.5C2 11.0523 1.55228 11.5 1 11.5C0.447715 11.5 0 11.0523 0 10.5C0 9.94772 0.447715 9.5 1 9.5C1.55228 9.5 2 9.94772 2 10.5ZM4.5 1C4.22386 1 4 1.22386 4 1.5C4 1.77614 4.22386 2 4.5 2H13.5C13.7761 2 14 1.77614 14 1.5C14 1.22386 13.7761 1 13.5 1H4.5ZM4 6C4 5.72386 4.22386 5.5 4.5 5.5H13.5C13.7761 5.5 14 5.72386 14 6C14 6.27614 13.7761 6.5 13.5 6.5H4.5C4.22386 6.5 4 6.27614 4 6ZM4.5 10C4.22386 10 4 10.2239 4 10.5C4 10.7761 4.22386 11 4.5 11H13.5C13.7761 11 14 10.7761 14 10.5C14 10.2239 13.7761 10 13.5 10H4.5Z" fill="#616161" />
                                </svg>

                                <TextField
                                    styles={styles.textArea}
                                    value={note}
                                    onChange={onNoteChange}
                                    placeholder="Add a note (optional)" multiline resizable={false} />
                            </Stack>

                        </form>
                        <DialogFooter>
                            <DefaultButton onClick={handleCloseModal} text="Close" />
                            <PrimaryButton onClick={handleSubmit} text="Save" />
                        </DialogFooter>
                    </div>

                </Modal>
            </Stack>
        </Stack>
    );
}