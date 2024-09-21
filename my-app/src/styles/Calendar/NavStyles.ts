import {
  IDropdownStyles,
  mergeStyleSets,
  getTheme,
  FontWeights,
  IButtonStyles,
} from "@fluentui/react";
import styled from 'styled-components';


export const styles: { [key: string]: any } = {
  checkbox: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    width: "100%",
  },
  timePicker: {
    root: {
      width: 89,
      backgroundColor: "#F5F5F5",
      div: {
        border: "none",
        backgroundColor: "#F5F5F5",
        width: "60px",
      },
      selectors: {
        "::after": {
          border: "none",
        },
      },
      button: {
        backgroundColor: "#F5F5F5",
      },
      input: {
        backgroundColor: "#F5F5F5",
        border: "none",
        width: "60px",
      },
    },
  },
  icon: {
    fontSize: "24px",
  },
  textField: {
    root: {
      div: {
        border: "none",
        width: "235px",
      },
      input: {
        backgroundColor: "#F5F5F5",
        border: "none",
      },
    },
  },
  textArea: {
    root: {
      width: "100%",
      div: {
        border: "none",
      },
      textarea: {
        backgroundColor: "#F5F5F5",
        border: "none",
      },
    },
  },
  filter: {
    padding: "6px 0px 6px 14px",
  },
};

const theme = getTheme();
export const contentStyles = mergeStyleSets({
  container: {
    display: "flex",
    flexFlow: "column nowrap",
    alignItems: "stretch",
    width: "808px",
  },
  header: [
    theme.fonts.xLargePlus,
    {
      flex: "1 1 auto",
      color: theme.palette.neutralPrimary,
      display: "flex",
      alignItems: "center",
      fontWeight: FontWeights.semibold,
      padding: "18px 12px 14px 24px",
    },
  ],
  heading: {
    color: theme.palette.neutralPrimary,
    fontWeight: FontWeights.semibold,
    fontSize: "inherit",
    margin: "0",
  },
  body: {
    flex: "4 4 auto",
    padding: "0 24px 24px 24px",
    overflowY: "hidden",
    selectors: {
      p: { margin: "14px 0" },
      "p:first-child": { marginTop: 0 },
      "p:last-child": { marginBottom: 0 },
    },
  },
});

export const buttonStyles: IButtonStyles = {
  root: {
    borderRadius: "4px",
    padding: "0 12px",
    height: "32px",
  },
  label: {
    fontSize: "14px",
    fontWeight: FontWeights.semibold,
  },
};

export const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: { width: "235px" },
  root: {
    div: { border: "none" },
    span: { backgroundColor: "#F5F5F5", border: "none" },
  },
};

export const StyledImage = styled.img`
  width: 26px;
  height: 26px;
  border-radius: 50%;
  margin-right: 10px;
  object-fit: cover;
`;
