import React, { useEffect, useState } from "react";
import * as XLSX from "xlsx";

// Convert Excel serial date to JS readable format
function excelDateToJSDate(serial) {
  const utc_days = Math.floor(serial - 25569);
  const utc_value = utc_days * 86400;
  const date_info = new Date(utc_value * 1000);
  return date_info.toLocaleDateString();
}

export default function Dashboard() {
  const [projects, setProjects] = useState([]);
  const [singleUsers, setSingleUsers] = useState([]);
  const [securityGroup, setSecurityGroup] = useState([]);
  const [region, setRegion] = useState("NAM");
  const [expandedTemplate, setExpandedTemplate] = useState(null);

  useEffect(() => {
    fetch("/data/projects.xlsx")
      .then((res) => res.arrayBuffer())
      .then((buffer) => {
        const workbook = XLSX.read(buffer, { type: "buffer" });
        const sheets = workbook.Sheets;

        const projects = XLSX.utils.sheet_to_json(sheets[workbook.SheetNames[0]]);
        const singleUsers = XLSX.utils.sheet_to_json(sheets[workbook.SheetNames[1]]);
        const securityGroup = XLSX.utils.sheet_to_json(sheets[workbook.SheetNames[2]]);

        setProjects(projects);
        setSingleUsers(singleUsers);
        setSecurityGroup(securityGroup);
      });
  }, []);

  const filteredProjects = projects.filter((row) => row.Region === region);
  const groupedByTemplate = filteredProjects.reduce((acc, row) => {
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

  const getUsersForTemplate = (templateKey) => {
    const users = singleUsers.filter(
      (u) => u["TEMPLATE_KEY"] === templateKey && u.Region === region
    );
    return users
      .map((user) => {
        const userDetails = securityGroup.find(
          (s) => s.USER_NAME === user["User SOE ID"]
        );
        return {
          soeId: user["User SOE ID"],
          displayName: userDetails?.DISPLAY_NAME || "N/A",
          email: userDetails?.EMAIL_ADDRESS || "N/A",
          group: userDetails?.GROUP_NAME || "N/A",
        };
      });
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
                onClick={() =>
                  setExpandedTemplate(expandedTemplate === template ? null : template)
                }
              >
                <td>{expandedTemplate === template ? "−" : "+"} {template}</td>
                <td>{getLatestDate(projects)}</td>
                <td>{projects.length}</td>
              </tr>

              {expandedTemplate === template && (
                <tr>
                  <td colSpan="3">
                    {/* Project Table */}
                    <table
                      border="1"
                      cellPadding="8"
                      cellSpacing="0"
                      style={{
                        width: "100%",
                        borderCollapse: "collapse",
                        marginTop: "10px",
                        marginBottom: "10px",
                      }}
                    >
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

                    {/* User Table */}
                    <table
                      border="1"
                      cellPadding="8"
                      cellSpacing="0"
                      style={{ width: "100%", borderCollapse: "collapse" }}
                    >
                      <thead>
                        <tr style={{ backgroundColor: "#f9f9f9" }}>
                          <th>User ID</th>
                          <th>Display Name</th>
                          <th>Email</th>
                          <th>Group</th>
                        </tr>
                      </thead>
                      <tbody>
                        {getUsersForTemplate(template).map((user, idx) => (
                          <tr key={idx}>
                            <td>{user.soeId}</td>
                            <td>{user.displayName}</td>
                            <td>{user.email}</td>
                            <td>{user.group}</td>
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
