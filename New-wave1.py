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
  const [expandedPage, setExpandedPage] = useState(0);

  const [singlePage, setSinglePage] = useState(0);
  const [secPage, setSecPage] = useState(0);
  const usersPerPage = 10;
  const [singleUserFilters, setSingleUserFilters] = useState({});

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

  // Project filtering
  const filteredData = data.filter(
    (row) =>
      row.Region === region &&
      (templateFilter === "" || row["Template Key"] === templateFilter)
  );

  // Group by Template
  const groupedByTemplate = filteredData.reduce((acc, row) => {
    const key = row["Template Key"];
    if (!acc[key]) acc[key] = [];
    acc[key].push(row);
    return acc;
  }, {});

  const expandedProjects = expandedTemplate
    ? groupedByTemplate[expandedTemplate] || []
    : [];

  const paginatedExpanded = expandedProjects.slice(
    expandedPage * usersPerPage,
    (expandedPage + 1) * usersPerPage
  );

  // Filter single users
  const filteredSingleUsers = singleUsers
    .filter((u) => u.Region === region && (templateFilter === "" || u["TEMPLATE_KEY"] === templateFilter))
    .filter((u) =>
      Object.entries(singleUserFilters).every(([key, val]) =>
        val === "" || (u[key] || "").toString().toLowerCase().includes(val.toLowerCase())
      )
    );

  const paginatedSingle = filteredSingleUsers.slice(
    singlePage * usersPerPage,
    (singlePage + 1) * usersPerPage
  );

  const filteredSecurityUsers = securityUsers.filter(
    (user) => user.Region === region &&
      (templateFilter === "" || user["TEMPLATE_KEY"] === templateFilter)
  );

  const paginatedSecurity = filteredSecurityUsers.slice(
    secPage * usersPerPage,
    (secPage + 1) * usersPerPage
  );

  const getLatestDate = (entries) => {
    const latest = entries
      .map((e) => e["Last Issue Updated"])
      .reduce((a, b) => (a > b ? a : b));
    return excelDateToJSDate(latest);
  };

  const handleFilterChange = (key, val) => {
    setSingleUserFilters((prev) => ({
      ...prev,
      [key]: val,
    }));
    setSinglePage(0);
  };

  const exportExpandedProjects = () => {
    if (expandedProjects.length > 0) {
      downloadExcel(expandedProjects, `${expandedTemplate}_projects.xlsx`);
    }
  };

  
