import React, { useState } from "react";
import * as XLSX from "xlsx";
import Select from "react-select";
import { saveAs } from "file-saver";
import "./App.css";

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

  const selectStyles = {
    control: (base) => ({
      ...base,
      borderRadius: "6px",
      border: "1px solid #D2665A",
      boxShadow: "0 1px 3px rgba(184, 33, 50, 0.05)",
      "&:hover": {
        border: "1px solid #B82132",
      },
    }),
    option: (base, state) => ({
      ...base,
      backgroundColor: state.isSelected ? "#B82132" : "white",
      color: state.isSelected ? "white" : "#333",
      padding: "10px 15px",
      "&:hover": {
        backgroundColor: state.isSelected ? "#B82132" : "#F6DED8",
      },
    }),
    menu: (base) => ({
      ...base,
      boxShadow: "0 4px 6px rgba(184, 33, 50, 0.1)",
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
      const uniqueValues = [...new Set(filteredData.map((row) => row[selected.value]))];
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
    setFilteredData(originalData);
    setSelectedColumns([]);
    setActiveFilters([]);
    setFilterColumn(null);
    setFilterValues([]);
    setFilterOptions([]);
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
    <div className="container">
      <h1>Excel Manipulation & Filtering Tool</h1>

      <div
        className={`upload-box ${isDragging ? 'dragging' : ''}`}
        onDrop={handleDrop}
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
      >
        <input
          type="file"
          accept=".xlsx, .xls"
          onChange={handleFileUpload}
          className="file-input"
        />
        <p>or drag and drop Excel file here</p>
      </div>

      {columns.length > 0 && (
        <div className="filter-section">
          <div className="select-container">
            <label>Select Columns to Include:</label>
            <Select
              options={[{ label: "Select All", value: "all" }, ...columns]}
              isMulti
              onChange={handleColumnSelect}
              value={selectedColumns}
              placeholder="Choose columns..."
              styles={selectStyles}
            />
          </div>

          <div className="select-container">
            <label>Filter by Column:</label>
            <Select
              options={selectedColumns}
              onChange={handleFilterColumnSelect}
              placeholder="Choose a column to filter..."
              styles={selectStyles}
            />
          </div>

          {filterColumn && (
            <div className="select-container">
              <label>Select Filter Values:</label>
              <Select
                options={filterOptions}
                isMulti
                onChange={(selected) => setFilterValues(selected || [])}
                placeholder="Choose values to filter..."
                styles={selectStyles}
              />
            </div>
          )}

          <div className="button-group">
            <button className="btn primary" onClick={applyFilter}>
              Apply Filter
            </button>
            <button className="btn secondary" onClick={resetFilters}>
              Reset Filters
            </button>
            <button className="btn accent" onClick={handlePreview}>
              {showPreview ? 'Hide Preview' : 'Show Preview'}
            </button>
          </div>
        </div>
      )}

      {showPreview && selectedColumns.length > 0 && (
        <div className="preview-section">
          <h3>Preview</h3>
          <div className="table-container">
            <table>
              <thead>
                <tr>
                  {selectedColumns.map((col) => (
                    <th key={col.value}>{col.label}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {filteredData.map((row, idx) => (
                  <tr key={idx}>
                    {selectedColumns.map((col) => (
                      <td key={col.value}>{row[col.value]}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <button className="btn primary" onClick={downloadExcel}>
            Download as Excel
          </button>
        </div>
      )}
    </div>
  );
};

export default App;