import React, { useState } from "react";
import * as XLSX from "xlsx";
import Select from "react-select";
import { saveAs } from "file-saver";

const App = () => {
  const [originalData, setOriginalData] = useState([]);
  const [selectedColumns, setSelectedColumns] = useState([]);
  const [filteredData, setFilteredData] = useState([]);
  const [columns, setColumns] = useState([]);
  const [showPreview, setShowPreview] = useState(false);
  const [filterColumn, setFilterColumn] = useState(null);
  const [filterValues, setFilterValues] = useState([]);
  const [filterOptions, setFilterOptions] = useState([]);
  const [activeFilters, setActiveFilters] = useState([]);

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    readExcel(file);
  };

  const handleDrop = (e) => {
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    readExcel(file);
  };

  const readExcel = (file) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const jsonData = XLSX.utils.sheet_to_json(sheet);
      const cols = Object.keys(jsonData[0]);
      setOriginalData(jsonData);
      setFilteredData(jsonData);
      setColumns(cols.map((col) => ({ label: col, value: col })));
      setSelectedColumns([]);
      setActiveFilters([]);
      setShowPreview(false);
    };
    reader.readAsArrayBuffer(file);
  };

  const handleColumnSelect = (selected) => {
    if (selected.find((option) => option.value === "all")) {
      setSelectedColumns(columns.filter((col) => col.value !== "all"));
    } else {
      setSelectedColumns(selected);
    }
  };

  const handlePreview = () => {
    setShowPreview(true);
  };

  const handleFilterColumnSelect = (selected) => {
    setFilterColumn(selected.value);
    const uniqueValues = [
      ...new Set(filteredData.map((row) => row[selected.value])),
    ];
    setFilterOptions(uniqueValues.map((val) => ({ label: val, value: val })));
  };

  const applyFilter = () => {
    if (!filterColumn || filterValues.length === 0) return;
    const newFilter = { column: filterColumn, values: filterValues.map((v) => v.value) };
    const updatedFilters = [...activeFilters, newFilter];
    setActiveFilters(updatedFilters);

    let tempData = [...originalData];
    updatedFilters.forEach(({ column, values }) => {
      tempData = tempData.filter((row) => values.includes(row[column]));
    });

    setFilteredData(tempData);
    setShowPreview(true);
  };

  const resetFilters = () => {
    setActiveFilters([]);
    setFilteredData(originalData);
    setShowPreview(false);
    setFilterColumn(null);
    setFilterValues([]);
  };

  const downloadExcel = () => {
    const ws = XLSX.utils.json_to_sheet(
      filteredData.map((row) => {
        const newRow = {};
        selectedColumns.forEach((col) => {
          newRow[col.value] = row[col.value];
        });
        return newRow;
      })
    );
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "FilteredData");
    const excelBuffer = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const file = new Blob([excelBuffer], { type: "application/octet-stream" });
    saveAs(file, "filtered_data.xlsx");
  };

  return (
    <div
      style={{
        padding: "30px",
        fontFamily: "Arial",
        background: "#f4f7fa",
        minHeight: "100vh",
      }}
      onDrop={handleDrop}
      onDragOver={(e) => e.preventDefault()}
    >
      <h1 style={{ textAlign: "center", color: "#2b3a42" }}>
        Excel Manipulation & Filtering Tool
      </h1>

      <div style={{ margin: "20px 0", textAlign: "center" }}>
        <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} />
        <p>or drag and drop Excel file above</p>
      </div>

      {columns.length > 0 && (
        <>
          <div style={{ marginBottom: "20px" }}>
            <label>Select Columns to Include:</label>
            <Select
              options={[{ label: "Select All", value: "all" }, ...columns]}
              isMulti
              onChange={handleColumnSelect}
              value={selectedColumns}
              placeholder="Choose columns..."
            />
          </div>

          <div style={{ marginBottom: "20px" }}>
            <label>Filter by Column:</label>
            <Select
              options={selectedColumns}
              onChange={handleFilterColumnSelect}
              placeholder="Choose a column to filter..."
            />
          </div>

          {filterColumn && (
            <div style={{ marginBottom: "20px" }}>
              <label>Select Filter Values:</label>
              <Select
                options={filterOptions}
                isMulti
                onChange={(selected) => setFilterValues(selected)}
                placeholder="Choose values to filter..."
              />
            </div>
          )}

          <div style={{ display: "flex", gap: "10px", marginBottom: "20px" }}>
            <button
              onClick={applyFilter}
              style={{
                background: "#4caf50",
                color: "white",
                padding: "10px 15px",
                border: "none",
                borderRadius: "5px",
              }}
            >
              Apply Filter
            </button>
            <button
              onClick={resetFilters}
              style={{
                background: "#f44336",
                color: "white",
                padding: "10px 15px",
                border: "none",
                borderRadius: "5px",
              }}
            >
              Reset Filters
            </button>
            <button
              onClick={handlePreview}
              style={{
                background: "#2196f3",
                color: "white",
                padding: "10px 15px",
                border: "none",
                borderRadius: "5px",
              }}
            >
              Preview
            </button>
          </div>
        </>
      )}

      {showPreview && selectedColumns.length > 0 && (
        <div>
          <h3>Preview</h3>
          <div style={{ overflowX: "auto", marginBottom: "20px" }}>
            <table
              style={{
                width: "100%",
                borderCollapse: "collapse",
                background: "white",
              }}
            >
              <thead>
                <tr>
                  {selectedColumns.map((col) => (
                    <th
                      key={col.value}
                      style={{
                        border: "1px solid #ddd",
                        padding: "8px",
                        backgroundColor: "#f2f2f2",
                      }}
                    >
                      {col.label}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {filteredData.map((row, idx) => (
                  <tr key={idx}>
                    {selectedColumns.map((col) => (
                      <td
                        key={col.value}
                        style={{ border: "1px solid #ddd", padding: "8px" }}
                      >
                        {row[col.value]}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <button
            onClick={downloadExcel}
            style={{
              background: "#673ab7",
              color: "white",
              padding: "10px 20px",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
            }}
          >
            Download as Excel
          </button>
        </div>
      )}
    </div>
  );
};
export default App;
