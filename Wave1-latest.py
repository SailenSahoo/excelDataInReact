import React, { useEffect, useState } from "react";
import * as XLSX from "xlsx";

function excelDateToJSDate(serial) {
  if (!serial || isNaN(serial)) return "";
  const utc_days = Math.floor(serial - 25569);
  const utc_value = utc_days * 86400;
  const date_info = new Date(utc_value * 1000);
  return date_info.toLocaleDateString();
}

function downloadExcel(data, filename) {
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  XLSX.writeFile(wb, filename);
}

export default function Dashboard() {
  const [data, setData] = useState([]);
  const [singleUsers, setSingleUsers] = useState([]);
  const [securityUsers, setSecurityUsers] = useState([]);
  const [region, setRegion] = useState("NAM");
  const [templateFilter, setTemplateFilter] = useState("");
  const [expandedTemplate, setExpandedTemplate] = useState(null);

  const [singlePage, setSinglePage] = useState(0);
  const [secPage, setSecPage] = useState(0);
  const [expandedPage, setExpandedPage] = useState(0);
  const usersPerPage = 10;

  const [singleUserFilters, setSingleUserFilters] = useState({});
  const [securityUserFilters, setSecurityUserFilters] = useState({});
  const [expandedFilters, setExpandedFilters] = useState({});

  useEffect(() => {
    fetch("/data/projects.xlsx")
      .then((res) => res.arrayBuffer())
      .then((buffer) => {
        const workbook = XLSX.read(buffer, { type: "buffer" });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(worksheet);
        setData(json);
      });

    fetch("/data/single_users.xlsx")
      .then((res) => res.arrayBuffer())
      .then((buffer) => {
        const workbook = XLSX.read(buffer, { type: "buffer" });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(worksheet);
        setSingleUsers(json);
      });

    fetch("/data/security_groups.xlsx")
      .then((res) => res.arrayBuffer())
      .then((buffer) => {
        const workbook = XLSX.read(buffer, { type: "buffer" });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(worksheet);
        setSecurityUsers(json);
      });
  }, []);

  const applyFilters = (data, filters) =>
    data.filter((row) =>
      Object.entries(filters).every(([key, value]) =>
        value ? (row[key] || "").toLowerCase().includes(value.toLowerCase()) : true
      )
    );

  const filteredData = data.filter(
    (row) =>
      row.Region === region &&
      (templateFilter === "" || row["Template Key"] === templateFilter)
  );

  const groupedByTemplate = filteredData.reduce((acc, row) => {
    const key = row["Template Key"];
    if (!acc[key]) acc[key] = [];
    acc[key].push(row);
    return acc;
  }, {});

  const filteredSingleUsers = applyFilters(
    singleUsers.filter(
      (user) =>
        user.Region === region &&
        (templateFilter === "" || user["TEMPLATE_KEY"] === templateFilter)
    ),
    singleUserFilters
  );

  const filteredSecurityUsers = applyFilters(
    securityUsers.filter((user) => user.Region === region),
    securityUserFilters
  );

  const filteredExpandedProjects = expandedTemplate
    ? applyFilters(groupedByTemplate[expandedTemplate] || [], expandedFilters)
    : [];

  const paginatedSingle = filteredSingleUsers.slice(
    singlePage * usersPerPage,
    (singlePage + 1) * usersPerPage
  );

  const paginatedSecurity = filteredSecurityUsers.slice(
    secPage * usersPerPage,
    (secPage + 1) * usersPerPage
  );

  const paginatedExpanded = filteredExpandedProjects.slice(
    expandedPage * usersPerPage,
    (expandedPage + 1) * usersPerPage
  );

  const getLatestDate = (entries) => {
    const latest = entries
      .map((e) => e["Last Issue Updated"])
      .reduce((a, b) => (a > b ? a : b));
    return excelDateToJSDate(latest);
  };

  return (
    <div style={{ padding: "20px", fontFamily: "Arial" }}>
      <h1>Project Dashboard</h1>

      {/* Region Buttons */}
      <div>
        {["NAM", "APAC"].map((r) => (
          <button
            key={r}
            onClick={() => setRegion(r)}
            style={{
              marginRight: "10px",
              padding: "5px 10px",
              backgroundColor: region === r ? "#007BFF" : "#ccc",
              color: "white",
              border: "none",
            }}
          >
            {r}
          </button>
        ))}
      </div>

      {/* Template Filter */}
      <div style={{ marginTop: "10px" }}>
        <label style={{ marginRight: "10px" }}>Filter by Template:</label>
        <select
          value={templateFilter}
          onChange={(e) => setTemplateFilter(e.target.value)}
        >
          <option value="">All Templates</option>
          {[...new Set(data.map((row) => row["Template Key"]))]
            .sort()
            .map((template) => (
              <option key={template} value={template}>
                {template}
              </option>
            ))}
        </select>
      </div>

      {/* Project Table */}
      <table border="1" cellPadding="6" cellSpacing="0" width="100%">
        <thead>
          <tr>
            <th>Template Key</th>
            <th>Latest Issue Updated</th>
            <th>Active Project Count</th>
          </tr>
        </thead>
        <tbody>
          {Object.entries(groupedByTemplate).map(([template, projects]) => (
            <React.Fragment key={template}>
              <tr
                style={{ backgroundColor: "#e6f2ff", cursor: "pointer" }}
                onClick={() => {
                  setExpandedTemplate(
                    expandedTemplate === template ? null : template
                  );
                  setExpandedPage(0);
                }}
              >
                <td>{expandedTemplate === template ? "−" : "+"} {template}</td>
                <td>{getLatestDate(projects)}</td>
                <td>{projects.length}</td>
              </tr>
              {expandedTemplate === template && (
                <tr>
                  <td colSpan="3">
                    <div style={{ marginTop: "10px" }}>
                      <p>
                        Total Records:{" "}
                        <a
                          href="#"
                          onClick={() =>
                            downloadExcel(filteredExpandedProjects, `${template}_projects.xlsx`)
                          }
                        >
                          {filteredExpandedProjects.length}
                        </a>
                      </p>
                      <table border="1" cellPadding="5" width="100%">
                        <thead>
                          <tr>
                            {["Active Project Key", "Last Issue Updated", "Active Project Name"].map((header) => (
                              <th key={header}>
                                {header}
                                <br />
                                <input
                                  type="text"
                                  placeholder="Filter"
                                  onChange={(e) =>
                                    setExpandedFilters((prev) => ({
                                      ...prev,
                                      [header]: e.target.value,
                                    }))
                                  }
                                />
                              </th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {paginatedExpanded.map((proj, idx) => (
                            <tr key={idx}>
                              <td>{proj["Active Project Key"]}</td>
                              <td>{excelDateToJSDate(proj["Last Issue Updated"])}</td>
                              <td>{proj["Active Project Name"]}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                      <div style={{ marginTop: "10px" }}>
                        <button
                          onClick={() => setExpandedPage((p) => Math.max(p - 1, 0))}
                          disabled={expandedPage === 0}
                        >
                          Previous
                        </button>
                        <button
                          onClick={() =>
                            setExpandedPage((p) =>
                              (p + 1) * usersPerPage < filteredExpandedProjects.length ? p + 1 : p
                            )
                          }
                          disabled={
                            (expandedPage + 1) * usersPerPage >=
                            filteredExpandedProjects.length
                          }
                          style={{ marginLeft: "5px" }}
                        >
                          Next
                        </button>
                      </div>
                    </div>
                  </td>
                </tr>
              )}
            </React.Fragment>
          ))}
        </tbody>
      </table>

      {/* Security Group Users Table */}
      <h2 style={{ marginTop: "40px" }}>Security Group Users</h2>
      <p>
        Total Users:{" "}
        <a
          href="#"
          onClick={() =>
            downloadExcel(filteredSecurityUsers, "security_group_users.xlsx")
          }
        >
          {filteredSecurityUsers.length}
        </a>
      </p>
      <table border="1" cellPadding="6" cellSpacing="0" width="100%">
        <thead>
          <tr>
            {["USER_NAME", "DISPLAY_NAME", "EMAIL_ADDRESS", "GROUP_NAME", "Region"].map(
              (header) => (
                <th key={header}>
                  {header}
                  <br />
                  <input
                    type="text"
                    placeholder="Filter"
                    onChange={(e) =>
                      setSecurityUserFilters((prev) => ({
                        ...prev,
                        [header]: e.target.value,
                      }))
                    }
                  />
                </th>
              )
            )}
          </tr>
        </thead>
        <tbody>
          {paginatedSecurity.map((user, idx) => (
            <tr key={idx}>
              <td>{user["USER_NAME"]}</td>
              <td>{user["DISPLAY_NAME"]}</td>
              <td>{user["EMAIL_ADDRESS"]}</td>
              <td>{user["GROUP_NAME"]}</td>
              <td>{user["Region"]}</td>
            </tr>
          ))}
        </tbody>
      </table>

      <div style={{ marginTop: "10px" }}>
        <button
          onClick={() => setSecPage((p) => Math.max(p - 1, 0))}
          disabled={secPage === 0}
        >
          Previous
        </button>
        <button
          onClick={() =>
            setSecPage((p) =>
              (p + 1) * usersPerPage < filteredSecurityUsers.length ? p + 1 : p
            )
          }
          disabled={(secPage + 1) * usersPerPage >= filteredSecurityUsers.length}
          style={{ marginLeft: "5px" }}
        >
          Next
        </button>
      </div>
    </div>
  );
}
