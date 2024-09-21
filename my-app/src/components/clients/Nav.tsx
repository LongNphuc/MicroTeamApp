import React from "react";

import { Input, Button, MenuTrigger } from "@fluentui/react-components";

import {
  Search20Regular,
  Filter20Regular,
  ChevronUp20Regular,
  ChevronDown20Regular,
} from "@fluentui/react-icons";

interface PropsNavChild {
    clientCounts : number | 0,
    toggleDataVisibility : () => void,
    isDataVisible : boolean
}

export function Nav({clientCounts, toggleDataVisibility, isDataVisible} : PropsNavChild) {
    
  return (
    <div>
      <header style={{ padding: "20px" }}>
        <h1 className="title">Clients</h1>
        <p>{clientCounts} clients</p>
      </header>

      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          padding: "10px 20px",
          backgroundColor: "#f5f5f5",
        }}
      >
        <div>
          <MenuTrigger>
            <Button
              style={{
                fontWeight: "500",
                fontSize: "18px",
                border: "none",
              }}
              // onMouseEnter={(e) =>
              //   (e.currentTarget.style.border = "1px solid blue")
              // }
              // onMouseLeave={(e) => (e.currentTarget.style.border = "none")}
              onClick={toggleDataVisibility}
              icon={
                isDataVisible ? (
                  <ChevronDown20Regular />
                ) : (
                  <ChevronUp20Regular />
                )
              }
              iconPosition="after"
            >
              Your clients
            </Button>
          </MenuTrigger>
        </div>

        <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
          <Button icon={<Filter20Regular />} appearance="subtle" />

          <div style={{ position: "relative" }}>
            <Input placeholder="Find" style={{ paddingRight: "30px" }} />
            <Search20Regular
              style={{
                position: "absolute",
                right: "10px",
                top: "50%",
                transform: "translateY(-50%)",
                pointerEvents: "none",
              }}
            />
          </div>

          <Button
            appearance="primary"
            style={{ backgroundColor: "#0078D4", color: "white" }}
          >
            Create New
          </Button>
        </div>
      </div>
    </div>
  );
}
