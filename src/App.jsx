import { useEffect, useState } from "react";
import "./App.css";
import * as XLSX from "xlsx";
import { Table } from "antd";

const columns = [
  {
    title: "Employee_Name",
    dataIndex: "Employee_Name",
    key: "Employee_Name",
  },
  {
    title: "Employee_EmailID",
    dataIndex: "Employee_EmailID",
    key: "Employee_EmailID",
  },
  {
    title: "Secret_Child_Name",
    dataIndex: "Secret_Child_Name",
    key: "Secret_Child_Name",
  },
  {
    title: "Secret_Child_EmailID",
    dataIndex: "Secret_Child_EmailID",
    key: "Secret_Child_EmailID",
  },
];

function App() {
  const [fileData, setFileData] = useState({});
  const [excelData, setExcelData] = useState([]);
  const [dataSource, setDataSource] = useState([]);
  const [dataSourceName, setDataSourceName] = useState('');

  //  Fetch compare excel file
  useEffect(() => {
    fetchExcelFile();
  }, []);

  useEffect(() => {
    if (excelData.length === 0) return;

    let assignments = [];
    let availableChildren = excelData.map(
      (employee) => employee.Employee_EmailID
    );

    // Shuffle the array to ensure randomness
    function shuffleArray(array) {
      for (let i = array.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [array[i], array[j]] = [array[j], array[i]];
      }
    }

    shuffleArray(availableChildren);

    excelData.forEach((employee) => {
      let validChildren = availableChildren.filter((childEmail) => {
        // Ensure the child isn't the employee themselves and wasn't assigned last year
        return (
          childEmail !== employee.Employee_EmailID &&
          !fileData.some(
            (lastAssignment) =>
              lastAssignment.Employee_EmailID === employee.Employee_EmailID &&
              lastAssignment.Secret_Child_EmailID === childEmail
          )
        );
      });

      if (validChildren.length === 0) {
        throw new Error(
          "No valid assignments found. Consider reshuffling or revising constraints."
        );
      }

      // Assign the first valid child and remove them from the pool
      let assignedChildEmail = validChildren[0];
      availableChildren = availableChildren.filter(
        (email) => email !== assignedChildEmail
      );

      let assignedChild = excelData.find(
        (emp) => emp.Employee_EmailID === assignedChildEmail
      );

      assignments.push({
        Employee_Name: employee.Employee_Name,
        Employee_EmailID: employee.Employee_EmailID,
        Secret_Child_Name: assignedChild.Employee_Name,
        Secret_Child_EmailID: assignedChild.Employee_EmailID,
      });
    });

    handleFileExport(assignments);
  }, [excelData, fileData]);

  const readFile = (arrayBuffer, type) => {
    const workbook = XLSX.read(arrayBuffer, { type });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet);

    return jsonData;
  };

  const fetchExcelFile = async () => {
    const response = await fetch("/Secret-Santa-Game-Result-2023.xlsx");
    const arrayBuffer = await response.arrayBuffer();

    const jsonData = readFile(arrayBuffer, "array");
    setFileData(jsonData);
  };

  const handleFileExport = (assignments) => {
    const fileName = `Secret-Santa-Game-Result-${new Date().getUTCFullYear()}.xlsx`;

    const worksheet = XLSX.utils.json_to_sheet(assignments);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet);

    XLSX.utils.sheet_add_aoa(
      worksheet,
      [
        [
          "Employee_Name",
          "Employee_EmailID",
          "Secret_Child_Name",
          "Secret_Child_EmailID",
        ],
      ],
      { origin: "A1" }
    );

    XLSX.writeFile(workbook, fileName, { compression: true });

    setDataSource(assignments);
    setDataSourceName(fileName);
  };

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    const reader = new FileReader();

    reader.onload = (event) => {
      const sheetData = readFile(event.target.result, "binary");
      setExcelData(sheetData);
    };

    reader.readAsBinaryString(file);
  };

  return (
    <>
      <div className="card">
        <div>
          <input type="file" onChange={handleFileUpload} />
          {excelData && excelData.length ? (
            <div style={{ textAlign: "center" }}>
              <h2>Exported Data : {dataSourceName}</h2>
              {dataSource && dataSource.length ? (
                <Table
                  dataSource={dataSource}
                  columns={columns}
                  pagination={{ pageSize: 20 }}
                />
              ) : (
                ""
              )}
            </div>
          ) : null}
        </div>
      </div>
    </>
  );
}

export default App;
