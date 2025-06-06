import React, { useEffect, useState } from "react";
import * as XLSX from "xlsx";

function excelDateToJSDate(serial) {
  if (typeof serial === "string") return serial;
  const utc_days = Math.floor(serial - 25569);
  const utc_value = utc_days * 86400;
  const date_info = new Date(utc_value * 1000);
  return date_info.toLocaleDateString();
}

export default function Dashboard() {
  const [data, setData] = useState([]);
  const [singleUsers, setSingleUsers] = useState([]);
  const [securityUsers, setSecurityUsers] = useState([]);
  const [region, setRegion] = useState("NAM");
  const [expandedTemplate, setExpandedTemplate] = useState(null);

  useEffect(() => {
    loadExcel("/data/projects.xlsx", setData);
    loadExcel("/data/single_users.xlsx", setSingleUsers);
    loadExcel("/data/security_groups.xlsx", setSecurityUsers);
  }, []);

  const loadExcel = (url, setter) => {
    fetch(url)
      .then((res) => res.arrayBuffer())
      .then((buffer) => {
        const workbook = XLSX.read(buffer, { type: "buffer" });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(worksheet);
        setter(json);
      });
  };

  const filteredData = data.filter((row) => row.Region === region);
  const groupedByTemplate = filteredData.reduce((acc, row) => {
    const key = row["Template Key"];
    if (!acc[key]) acc[key] = [];
    acc[key].push(row);
    return acc;
  }, {});

  const getLatestDate = (entries) => {
    const latest = entries
      .map((e) => e["Last Issue Updated"])
      .reduce((a, b) => (a > b ? a : b));
    return excelDateToJSDate(latest);
  };

  return (
    <div style={{ padding: "20px", fontFamily: "Arial" }}>
      <h1 style={{ fontSize: "24px", marginBottom: "20px" }}>Project Dashboard</h1>

      <div style={{ marginBottom: "20px" }}>
        <button
          onClick={() => setRegion("NAM")}
          style={{
            marginRight: "10px",
            padding: "5px 10px",
            backgroundColor: region === "NAM" ? "#007BFF" : "#ccc",
            color: "white",
            border: "none",
            cursor: "pointer",
          }}
        >
          NAM
        </button>
        <button
          onClick={() => setRegion("APAC")}
          style={{
            padding: "5px 10px",
            backgroundColor: region === "APAC" ? "#007BFF" : "#ccc",
            color: "white",
            border: "none",
            cursor: "pointer",
          }}
        >
          APAC
        </button>
      </div>

      <table border="1" cellPadding="10" cellSpacing="0" style={{ width: "100%", borderCollapse: "collapse" }}>
        <thead>
          <tr style={{ backgroundColor: "#f2f2f2" }}>
            <th>Template Key</th>
            <th>Latest Issue Updated</th>
            <th>Active Project Count</th>
          </tr>
        </thead>
        <tbody>
          {Object.entries(groupedByTemplate).map(([template, projects]) => (
            <React.Fragment key={template}>
              <tr
                style={{ cursor: "pointer", backgroundColor: "#e6f2ff" }}
                onClick={() => setExpandedTemplate(expandedTemplate === template ? null : template)}
              >
                <td>{expandedTemplate === template ? "−" : "+"} {template}</td>
                <td>{getLatestDate(projects)}</td>
                <td>{projects.length}</td>
              </tr>
              {expandedTemplate === template && (
                <tr>
                  <td colSpan="3">
                    {/* Project Table */}
                    <h4>Projects</h4>
                    <table border="1" cellPadding="8" cellSpacing="0" style={{ width: "100%", borderCollapse: "collapse", marginBottom: "20px" }}>
                      <thead style={{ backgroundColor: "#f9f9f9" }}>
                        <tr>
                          <th>Project Key</th>
                          <th>Last Issue Updated</th>
                          <th>Project Name</th>
                        </tr>
                      </thead>
                      <tbody>
                        {projects.map((proj, idx) => (
                          <tr key={idx}>
                            <td>{proj["Active Project Key"]}</td>
                            <td>{excelDateToJSDate(proj["Last Issue Updated"])}</td>
                            <td>{proj["Active Project Name"]}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>

                    {/* Single Users Table */}
                    <h4>Single Users</h4>
                    <table border="1" cellPadding="8" cellSpacing="0" style={{ width: "100%", borderCollapse: "collapse", marginBottom: "20px" }}>
                      <thead style={{ backgroundColor: "#f9f9f9" }}>
                        <tr>
                          <th>User SOE ID</th>
                          <th>Region</th>
                        </tr>
                      </thead>
                      <tbody>
                        {singleUsers
                          .filter((user) => user["TEMPLATE_KEY"] === template && user.Region === region)
                          .map((user, idx) => (
                            <tr key={idx}>
                              <td>{user["User SOE ID"]}</td>
                              <td>{user.Region}</td>
                            </tr>
                          ))}
                      </tbody>
                    </table>

                    {/* Security Group Users Table */}
                    <h4>Security Group Users</h4>
                    <table border="1" cellPadding="8" cellSpacing="0" style={{ width: "100%", borderCollapse: "collapse" }}>
                      <thead style={{ backgroundColor: "#f9f9f9" }}>
                        <tr>
                          <th>Username</th>
                          <th>Display Name</th>
                          <th>Email</th>
                          <th>Group</th>
                        </tr>
                      </thead>
                      <tbody>
                        {securityUsers
                          .filter((user) => user.Region === region)
                          .map((user, idx) => (
                            <tr key={idx}>
                              <td>{user.USER_NAME}</td>
                              <td>{user.DISPLAY_NAME}</td>
                              <td>{user.EMAIL_ADDRESS}</td>
                              <td>{user.GROUP_NAME}</td>
                            </tr>
                          ))}
                      </tbody>
                    </table>
                  </td>
                </tr>
              )}
            </React.Fragment>
          ))}
        </tbody>
      </table>
    </div>
  );
}
