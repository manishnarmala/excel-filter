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
  const [isDragging, setIsDragging] = useState(false);

  const buttonStyles = {
    base: {
      color: "white",
      padding: "12px 20px",
      border: "none",
      borderRadius: "6px",
      cursor: "pointer",
      transition: "all 0.2s ease",
      fontWeight: "600",
      textTransform: "uppercase",
      fontSize: "0.9rem",
      boxShadow: "0 2px 4px rgba(0,0,0,0.1)",
      "&:hover": {
        transform: "translateY(-1px)",
        boxShadow: "0 4px 6px rgba(0,0,0,0.1)",
      },
    },
    green: { background: "#4caf50" },
    red: { background: "#f44336" },
    blue: { background: "#2196f3" },
    purple: { background: "#673ab7" },
  };

  const selectStyles = {
    control: (base) => ({
      ...base,
      borderRadius: "6px",
      border: "1px solid #e0e0e0",
      boxShadow: "0 1px 3px rgba(0,0,0,0.05)",
      "&:hover": {
        border: "1px solid #2196f3",
      },
    }),
    option: (base, state) => ({
      ...base,
      backgroundColor: state.isSelected ? "#2196f3" : "white",
      padding: "10px 15px",
      "&:hover": {
        backgroundColor: state.isSelected ? "#2196f3" : "#f5f5f5",
      },
    }),
    menu: (base) => ({
      ...base,
      boxShadow: "0 4px 6px rgba(0,0,0,0.1)",
      borderRadius: "6px",
    }),
  };

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    readExcel(file);
  };

  const handleDrop = (e) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files[0];
    readExcel(file);
  };

  const handleDragOver = (e) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = (e) => {
    e.preventDefault();
    setIsDragging(false);
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
    if (selected?.find((option) => option.value === "all")) {
      setSelectedColumns(columns.filter((col) => col.value !== "all"));
    } else {
      setSelectedColumns(selected || []);
    }
  };

  const handlePreview = () => {
    setShowPreview(prev => !prev); 
  };

  const handleFilterColumnSelect = (selected) => {
    setFilterColumn(selected?.value || null);
    if (selected) {
      const uniqueValues = [
        ...new Set(filteredData.map((row) => row[selected.value])),
      ];
      setFilterOptions(uniqueValues.map((val) => ({ label: val, value: val })));
    }
  };

  const applyFilter = () => {
    const newFilteredData = originalData.filter((row) => {
      if (!filterColumn || !filterValues.length) return true;
      return filterValues.some((filter) => filter.value === row[filterColumn]);
    });
    setFilteredData(newFilteredData);
    setActiveFilters([...activeFilters, { column: filterColumn, values: filterValues }]);
    setFilterColumn(null);
    setFilterValues([]);
  };

  const resetFilters = () => {
    // Reset filtered data to original data
    setFilteredData(originalData);
    
    // Reset all column and filter related states
    setSelectedColumns([]);
    setActiveFilters([]);
    setFilterColumn(null);
    setFilterValues([]);
    setFilterOptions([]);
    
    // Hide the preview section
    setShowPreview(false);
  };

  const downloadExcel = () => {
    const worksheet = XLSX.utils.json_to_sheet(
      filteredData.map(row => {
        const newRow = {};
        selectedColumns.forEach(col => {
          newRow[col.label] = row[col.value];
        });
        return newRow;
      })
    );
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Filtered Data");
    const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    const data = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(data, 'filtered_data.xlsx');
  };

  return (
    <div
      style={{
        padding: "40px",
        fontFamily: "'Segoe UI', Arial, sans-serif",
        background: "linear-gradient(to bottom, #f4f7fa, #e8eef5)",
        minHeight: "100vh",
        maxWidth: "1200px",
        margin: "0 auto",
      }}
    >
      <h1 style={{ 
        textAlign: "center", 
        color: "#1a237e",
        fontSize: "2.5rem",
        marginBottom: "40px",
        fontWeight: "600"
      }}>
        Excel Manipulation & Filtering Tool
      </h1>

      <div
        style={{
          margin: "30px auto",
          maxWidth: "600px",
          textAlign: "center",
          padding: "40px",
          border: `2px dashed ${isDragging ? "#2196f3" : "#ccc"}`,
          borderRadius: "10px",
          background: "white",
          transition: "all 0.2s ease",
          cursor: "pointer",
          boxShadow: isDragging ? "0 0 10px rgba(33,150,243,0.3)" : "none",
        }}
        onDrop={handleDrop}
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
      >
        <input
          type="file"
          accept=".xlsx, .xls"
          onChange={handleFileUpload}
          style={{ marginBottom: "15px" }}
        />
        <p style={{ color: "#666", margin: "10px 0" }}>
          or drag and drop Excel file here
        </p>
      </div>

      {columns.length > 0 && (
        <div style={{ 
          background: "white", 
          padding: "30px",
          borderRadius: "10px",
          boxShadow: "0 2px 8px rgba(0,0,0,0.1)",
          margin: "20px 0"
        }}>
          <div style={{ marginBottom: "20px" }}>
            <label style={{ display: "block", marginBottom: "8px", color: "#333" }}>
              Select Columns to Include:
            </label>
            <Select
              options={[{ label: "Select All", value: "all" }, ...columns]}
              isMulti
              onChange={handleColumnSelect}
              value={selectedColumns}
              placeholder="Choose columns..."
              styles={selectStyles}
            />
          </div>

          <div style={{ marginBottom: "20px" }}>
            <label style={{ display: "block", marginBottom: "8px", color: "#333" }}>
              Filter by Column:
            </label>
            <Select
              options={selectedColumns}
              onChange={handleFilterColumnSelect}
              placeholder="Choose a column to filter..."
              styles={selectStyles}
            />
          </div>

          {filterColumn && (
            <div style={{ marginBottom: "20px" }}>
              <label style={{ display: "block", marginBottom: "8px", color: "#333" }}>
                Select Filter Values:
              </label>
              <Select
                options={filterOptions}
                isMulti
                onChange={(selected) => setFilterValues(selected || [])}
                placeholder="Choose values to filter..."
                styles={selectStyles}
              />
            </div>
          )}

          <div style={{ display: "flex", gap: "10px", marginBottom: "20px" }}>
            <button
              onClick={applyFilter}
              style={{
                ...buttonStyles.base,
                ...buttonStyles.green,
              }}
            >
              Apply Filter
            </button>
            <button
              onClick={resetFilters}
              style={{
                ...buttonStyles.base,
                ...buttonStyles.red,
              }}
            >
              Reset Filters
            </button>
            <button
  onClick={handlePreview}
  style={{
    ...buttonStyles.base,
    ...buttonStyles.blue,
  }}
>
  {showPreview ? 'Hide Preview' : 'Show Preview'}
</button>
          </div>
        </div>
      )}

      {showPreview && selectedColumns.length > 0 && (
        <div style={{
          background: "white",
          padding: "30px",
          borderRadius: "10px",
          boxShadow: "0 2px 8px rgba(0,0,0,0.1)",
          margin: "20px 0"
        }}>
          <h3 style={{ 
            color: "#1a237e",
            marginBottom: "20px",
            fontSize: "1.5rem"
          }}>Preview</h3>
          <div style={{ overflowX: "auto", marginBottom: "20px" }}>
            <table style={{
              width: "100%",
              borderCollapse: "collapse",
              background: "white",
              borderRadius: "8px",
              overflow: "hidden",
              boxShadow: "0 1px 3px rgba(0,0,0,0.1)",
            }}>
              <thead>
                <tr>
                  {selectedColumns.map((col) => (
                    <th
                      key={col.value}
                      style={{
                        border: "1px solid #eee",
                        padding: "12px 15px",
                        backgroundColor: "#f8f9fa",
                        fontWeight: "600",
                        textAlign: "left",
                      }}
                    >
                      {col.label}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {filteredData.map((row, idx) => (
                  <tr 
                    key={idx}
                    style={{
                      '&:hover': {
                        backgroundColor: "#f5f5f5"
                      }
                    }}
                  >
                    {selectedColumns.map((col) => (
                      <td
                        key={col.value}
                        style={{
                          border: "1px solid #eee",
                          padding: "12px 15px",
                        }}
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
              ...buttonStyles.base,
              ...buttonStyles.purple,
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