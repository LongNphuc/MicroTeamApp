import React from "react";

import {
  TableBody,
  TableCell,
  TableRow,
  Table,
  TableHeader,
  TableHeaderCell,
  TableCellLayout,
  Button,
} from "@fluentui/react-components";

const itemsPerPage = 10;

const clients = [
  {
    name: "John Doe",
    type: "VIP",
    contactInfo: "john@example.com",
    relationship: "Client",
  },
  {
    name: "Jane Smith2",
    type: "Regular",
    contactInfo: "jane@example.com",
    relationship: "Partner",
  },
  {
    name: "Jane Smith3",
    type: "Regular",
    contactInfo: "jane@example.com",
    relationship: "Partner",
  },
  {
    name: "Jane Smith4",
    type: "Regular",
    contactInfo: "jane@example.com",
    relationship: "Partner",
  },
  {
    name: "Jane Smith5",
    type: "Regular",
    contactInfo: "jane@example.com",
    relationship: "Partner",
  },
  {
    name: "Jane Smith6",
    type: "Regular",
    contactInfo: "jane@example.com",
    relationship: "Partner",
  },
  {
    name: "Jane Smith7",
    type: "Regular",
    contactInfo: "jane@example.com",
    relationship: "Partner",
  },
  {
    name: "Jane Smith8",
    type: "Regular",
    contactInfo: "jane@example.com",
    relationship: "Partner",
  },
  {
    name: "Jane Smith9",
    type: "Regular",
    contactInfo: "jane@example.com",
    relationship: "Partner",
  },
  {
    name: "Jane Smith10",
    type: "Regular",
    contactInfo: "jane@example.com",
    relationship: "Partner",
  },
];

const columns = [
  { columnKey: "name", label: "Name" },
  { columnKey: "type", label: "Type" },
  { columnKey: "contactInfo", label: "Contact Info" },
  { columnKey: "relationship", label: "Relationship" },
];

interface PropsTableContentChild {
  isDataVisible: boolean;
}

export function TableContent({ isDataVisible }: PropsTableContentChild) {
  const [currentPage, setCurrentPage] = React.useState(1);
  const [currentData, setCurrentData] = React.useState(
    clients.slice(0, itemsPerPage)
  );
  const totalPages = React.useMemo(
    () => Math.ceil(clients.length / itemsPerPage),
    []
  );

  const updateCurrentData = React.useCallback(() => {
    const startIndex = (currentPage - 1) * itemsPerPage;
    const endIndex = startIndex + itemsPerPage;
    setCurrentData(clients.slice(startIndex, endIndex));
  }, [currentPage]);

  React.useEffect(() => {
    updateCurrentData();
  }, [currentPage, updateCurrentData]);

  const goToNextPage = React.useCallback(() => {
    setCurrentPage((prev) => Math.min(prev + 1, totalPages));
  }, [totalPages]);

  const goToPreviousPage = React.useCallback(() => {
    setCurrentPage((prev) => Math.max(prev - 1, 1));
  }, []);
  return (
    <div
      style={{
        padding: "20px",
      }}
    >
      {isDataVisible && (
        <div>
          {clients.length !== 0 && (
            <Table arial-label="Clients Table" style={{ minWidth: "600px" }}>
              <TableHeader>
                <TableRow>
                  {columns.map((column) => (
                    <TableHeaderCell key={column.columnKey}>
                      {column.label}
                    </TableHeaderCell>
                  ))}
                </TableRow>
              </TableHeader>
              <TableBody>
                {currentData.map((client) => (
                  <TableRow key={client.name}>
                    <TableCell>
                      <TableCellLayout>{client.name}</TableCellLayout>
                    </TableCell>
                    <TableCell>
                      <TableCellLayout>{client.type}</TableCellLayout>
                    </TableCell>
                    <TableCell>
                      <TableCellLayout>{client.contactInfo}</TableCellLayout>
                    </TableCell>
                    <TableCell>
                      <TableCellLayout>{client.relationship}</TableCellLayout>
                    </TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
          )}
          {clients.length === 0 && <p>No clients to display.</p>}
          <div
            style={{
              marginTop: "20px",
              display: "flex",
              justifyContent: "center",
              gap: "10px",
            }}
          >
            <Button
              disabled={currentPage === 1}
              onClick={goToPreviousPage}
              appearance="primary"
            >
              Previous
            </Button>

            <span>
              {currentPage} of {totalPages}
            </span>

            <Button
              disabled={currentPage === totalPages}
              onClick={goToNextPage}
              appearance="primary"
            >
              Next
            </Button>
          </div>
        </div>
      )}
    </div>
  );
}
