import { CSSProperties } from "react";

export const styles: { [key: string]: CSSProperties } = {
  timeLine: {
    position: "absolute",
    top: "0",
    borderTop: "1px dashed #ff0000",
    width: "100%",
    zIndex: "1000",
  },
  timeLineActive: {
    position: "absolute",
    top: "0",
    borderTop: "1px solid #ff0000",
    width: "100%",
  },
  beforeDot: {
    position: "absolute",
    top: "-3px",
    left: "-3px",
    width: "6px",
    height: "6px",
    borderRadius: "50%",
    backgroundColor: "#ff0000",
    zIndex: "1000",
  },
  afterDot: {
    position: "absolute",
    top: "-3px",
    right: "-3px",
    width: "6px",
    height: "6px",
    borderRadius: "50%",
    backgroundColor: "#ff0000",
    zIndex: "1000",
  },
  appointmentContainer: {
    position: "relative",
    height: 18,
    marginBottom: 1,
  },
  cellBlank: {
    width: 51,
    borderTop: "1px solid #e0e0e0",
    marginTop: "-1px",
  },
  containAppointment: {
    width: "calc(96%)",
    height: "100%",
    position: "absolute",
    top: "0",
    backgroundColor: "transparent",
    overflow: "hidden",
    pointerEvents: "none",
  },
  holdManyDays: {
    borderLeft: "1px solid #e0e0e0",
    borderBottom: "1px solid #e0e0e0",
    flex: "1",
    overflow: "hidden",
  },
};
