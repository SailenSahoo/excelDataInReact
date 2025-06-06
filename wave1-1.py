import React, { useEffect, useState } from "react";
import * as XLSX from "xlsx";

// Helper to convert Excel date to readable format
function excelDateToJSDate(serial) {
  if (typeof serial === "string") return serial; // Already formatted
  const utc_days = Math.floor(serial - 25569);
  const utc_value = utc_days * 86400;
  const date_info = new Date(utc_value * 1000);
  return date_info.toLocaleDateString();
}

export default function Dashboard() {
  const [projectData, setProjectData] = useState([]);
  const [singleUsers, setSingleUsers] = useState([]);
  const [securityUsers, setSecurityUsers] = useState([]);
  const [region, setRegion] = useState("NAM");
  const [expandedTemplate, setExpandedTemplate] = useState(null);

  useEffect(() => {
    const fetchExcel = async (path) => {
      const res = await fetch(path);
      const buffer = await res.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "buffer" });
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      return XLSX.utils.sheet_to_json(worksheet);
    };

    Promise.all([
      fetchExcel("/data/projects.xlsx"),
      fetchExcel("/data/single_users.xlsx"),
      fetchExcel("/data/security_groups.xlsx"),
    ]).then(([projects, single, security]) => {
      setProjectData(projects);
      setSingleUsers(single);
      setSecurityUsers(security);
    });
  }, []);

  const filteredProjects = projectData.filter((row) => row.Region === region);
  const groupedByTemplate = filteredProjects.reduce((acc, row) => {
    const key = row["Template Key"];
    if (!acc[key]) acc[key] = [];
    acc[key].push(row);
    return acc;
  }, {});

  const getLatestDate = (entries) => {
    const dates = entries.map((e) => e["Last Issue Updated"]);
    const latest = Math.max(...dates.map((d) => (typeof d === "number" ? d : new Date(d).getTime())));
    return excelDateToJSDate(latest);
  };

  const filteredSingleUsers = singleUsers.filter((u) => u.Region === region);
  const filteredSecurityUsers = securityUsers.filter((u) => u.Region === region);

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

      {/* Templates Table */}
      <h2>Templates</h2>
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
                    <table border="1" cellPadding="8" cellSpacing="0" style={{ width: "100%", marginTop: "10px" }}>
                      <thead>
                        <tr style={{ backgroundColor: "#f9f9f9" }}>
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
                  </td>
                </tr>
              )}
            </React.Fragment>
          ))}
        </tbody>
      </table>

      {/* Users Section */}
      <h2 style={{ marginTop: "40px" }}>Single Users</h2>
      <table border="1" cellPadding="10" cellSpacing="0" style={{ width: "100%", borderCollapse: "collapse" }}>
        <thead style={{ backgroundColor: "#f2f2f2" }}>
          <tr>
            <th>Template Key</th>
            <th>User SOE ID</th>
            <th>LENGTH_SOE_ID</th>
          </tr>
        </thead>
        <tbody>
          {filteredSingleUsers.map((user, idx) => (
            <tr key={idx}>
              <td>{user["TEMPLATE_KEY"]}</td>
              <td>{user["User SOE ID"]}</td>
              <td>{user["LENGTH_SOE_ID"]}</td>
            </tr>
          ))}
        </tbody>
      </table>

      <h2 style={{ marginTop: "40px" }}>Security Group Users</h2>
      <table border="1" cellPadding="10" cellSpacing="0" style={{ width: "100%", borderCollapse: "collapse" }}>
        <thead style={{ backgroundColor: "#f2f2f2" }}>
          <tr>
            <th>User Name</th>
            <th>Display Name</th>
            <th>Email</th>
            <th>Group Name</th>
            <th>LENGTH_SOE_ID</th>
          </tr>
        </thead>
        <tbody>
          {filteredSecurityUsers.map((user, idx) => (
            <tr key={idx}>
              <td>{user["USER_NAME"]}</td>
              <td>{user["DISPLAY_NAME"]}</td>
              <td>{user["EMAIL_ADDRESS"]}</td>
              <td>{user["GROUP_NAME"]}</td>
              <td>{user["LENGTH_SOE_ID"]}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
