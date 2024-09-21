import React from "react";

import { Nav } from "./clients/Nav";
import { TableContent } from "./clients/TableDisplay";

// interface Props {
// }

export function Clients() {
  const [isDataVisible, setIsDataVisible] = React.useState(true);

  const toggleDataVisibility = () => {
    setIsDataVisible((prev) => !prev);
  };
  return (
    <>
      <Nav
        clientCounts={10}
        toggleDataVisibility={toggleDataVisibility}
        isDataVisible={isDataVisible}
      />
      <TableContent isDataVisible={isDataVisible} />
    </>
  );
}
