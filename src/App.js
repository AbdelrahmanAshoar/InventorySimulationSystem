import React, { useEffect, useState } from "react";
import * as XLSX from "xlsx";
import "./App.css";
import { tab } from "@testing-library/user-event/dist/tab";
import { Button } from "@mui/material";
import Table from "@mui/material/Table";
import TableBody from "@mui/material/TableBody";
import TableCell from "@mui/material/TableCell";
import TableContainer from "@mui/material/TableContainer";
import TableHead from "@mui/material/TableHead";
import TableRow from "@mui/material/TableRow";
import Paper from "@mui/material/Paper";
import { saveAs } from "file-saver";
import SaveIcon from "@mui/icons-material/Save";

const App = () => {
  const [tablesData, setTablesData] = useState([]);
  const [dataSystem, setDataSystem] = useState([]);
  const [prob, setprob] = useState([
    { type: "", prob: 0, cumulative: 0, from: 0, to: 0 },
  ]);
  const [demand, seDemand] = useState([
    {
      demand: 0,
      probGood: 0,
      probFair: 0,
      probPoor: 0,
      cumulativeGood: 0,
      cumulativeFair: 0,
      cumulativePoor: 0,
      fromGood: 0,
      toGood: 0,
      fromFair: 0,
      toFair: 0,
      fromPoor: 0,
      toPoor: 0,
    },
  ]);
  const [simulatedData, setSimulatedData] = useState([
    {
      day: 0,
      RD_Type: 0,
      type: "",
      RD_Demand: 0,
      demand: 0,
      revenueFromSales: 0,
      excessDemand: 0,
      lostProfitExcessDemand: 0,
      numberOfScrap: 0,
      salvage: 0,
      dailyProfit: 0,
    },
  ]);

  const [fileUploaded, setFileUploaded] = useState(false);

  const handleFileUpload = (event) => {
    const file = event.target.files[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = (e) => {
        const binaryStr = e.target.result;
        const workbook = XLSX.read(binaryStr, { type: "binary" });
        const tables = [];

        workbook.SheetNames.forEach((sheetName) => {
          const worksheet = workbook.Sheets[sheetName];
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

          let table = [];
          jsonData.forEach((row) => {
            if (row.length === 0 && table.length > 0) {
              tables.push(table);
              table = [];
            } else if (row.length > 0) {
              table.push(row);
            }
          });
          if (table.length > 0) {
            tables.push(table);
          }
        });

        // Convert each table's array to array of objects using the first row as headers
        const tablesWithObjects = tables.filter((table) => {
          const headers = table[0];

          return table.slice(1).map((row) => {
            const obj = {};
            headers.forEach((header, index) => {
              obj[header] = row[index] || "";
            });
            return obj;
          });
        });

        // --------------------
        setFileUploaded(true);
        setTablesData(tablesWithObjects.filter((table) => table.length > 1));
        const [, ...arr] = tablesWithObjects[1];

        let prevCumulative = 0;
        let _from = 0,
          _to = null;
        setprob(
          arr.map((news) => {
            prevCumulative += news[1];
            _from = _to == null ? 0 : _to + 1;
            _to = +(prevCumulative * 100).toFixed(0);
            return {
              type: news[0],
              prob: news[1],
              cumulative: prevCumulative,
              from: _from,
              to: _to,
            };
          })
        );

        const [, , ...arr1] = tablesWithObjects[3];
        let prevCumulativeGood = 0,
          _fromGood = 0,
          _toGood = null;
        let prevCumulativeFair = 0,
          _fromFair = 0,
          _toFair = null;
        let prevCumulativePoor = 0,
          _fromPoor = 0,
          _toPoor = null;

        seDemand(
          arr1.map((d) => {
            prevCumulativeGood += d[1];
            prevCumulativeFair += d[2];
            prevCumulativePoor += d[3];
            _fromGood = _toGood == null ? 0 : _toGood + 1;
            _toGood = +(prevCumulativeGood * 100).toFixed(0);
            _fromFair = _toFair == null ? 0 : _toFair + 1;
            _toFair = +(prevCumulativeFair * 100).toFixed(0);
            _fromPoor = _toPoor == null ? 0 : _toPoor + 1;
            _toPoor = +(prevCumulativePoor * 100).toFixed(0);
            return {
              demand: d[0],
              probGood: d[1],
              probFair: d[2],
              probPoor: d[3],
              cumulativeGood: prevCumulativeGood,
              cumulativeFair: prevCumulativeFair,
              cumulativePoor: prevCumulativePoor,
              fromGood: _fromGood,
              toGood: _toGood,
              fromFair: _fromFair,
              toFair: _toFair,
              fromPoor: _fromPoor,
              toPoor: _toPoor,
            };
          })
        );

        const [, ...arr2] = tablesWithObjects[2];
        setDataSystem(arr2);
      };
      reader.readAsBinaryString(file);
    }
  };

  const findType = (RDNumber) => {
    if (RDNumber >= prob[0].from && RDNumber <= prob[0].to) {
      return prob[0];
    } else if (RDNumber >= prob[1].from && RDNumber <= prob[1].to) {
      return prob[1];
    } else return prob[2];
  };

  const findDemand = (RDNumber, type) => {
    if (type === "Good") {
      for (let i = 0; i < demand.length; i++) {
        if (RDNumber >= demand[i].fromGood && RDNumber <= demand[i].toGood) {
          return demand[i];
        }
      }
    } else if (type === "Fair") {
      for (let i = 0; i < demand.length - 1; i++) {
        if (RDNumber >= demand[i].fromFair && RDNumber <= demand[i].toFair) {
          return demand[i];
        }
      }
    } else {
      for (let i = 0; i < demand.length - 2; i++) {
        if (RDNumber >= demand[i].fromPoor && RDNumber <= demand[i].toPoor) {
          return demand[i];
        }
      }
    }
  };

  let randomNumberForType, randomNumberForDemand;
  const arr = [];

  const simulteData = () => {
    for (let i = 1; i <= 5; i++) {
      randomNumberForType = Math.floor(Math.random() * 101);
      randomNumberForDemand = Math.floor(Math.random() * 101);
      let _type = findType(randomNumberForType).type;
      let _demand = findDemand(randomNumberForDemand, _type).demand;

      let _revenueFromSales =
        _demand >= dataSystem[0][1]
          ? (dataSystem[0][1] * dataSystem[2][1]).toFixed(2)
          : (_demand * dataSystem[2][1]).toFixed(2);
      let _excessDemand =
        _demand >= dataSystem[0][1] ? _demand - dataSystem[0][1] : 0;
      let _lostProfitExcessDemand = (_excessDemand * dataSystem[3][1]).toFixed(
        2
      );
      let _numberOfScrap =
        dataSystem[0][1] >= _demand ? dataSystem[0][1] - _demand : 0;
      let _dailyProfit = (
        +_revenueFromSales +
        +_lostProfitExcessDemand -
        +dataSystem[0][1] * +dataSystem[1][1] -
        +_lostProfitExcessDemand
      ).toFixed(2);
      arr.push({
        day: i,
        RD_Type: randomNumberForType,
        type: _type,
        RD_Demand: randomNumberForDemand,
        demand: _demand,
        revenueFromSales: _revenueFromSales,
        excessDemand: _excessDemand,
        lostProfitExcessDemand: _lostProfitExcessDemand,
        numberOfScrap: _numberOfScrap,
        salvage: _numberOfScrap * dataSystem[4][1],
        dailyProfit: _dailyProfit,
      });
    }

    setSimulatedData([...simulatedData, ...arr]);
  };

  const exportToExcel = () => {
    const worksheet = XLSX.utils.json_to_sheet(simulatedData);

    // Define column widths
    let cellsWidth = [];
    for (let i = 0; i < 11; i++) {
      cellsWidth.push({ wch: 15 });
    }
    worksheet["!cols"] = cellsWidth;

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
    const excelBuffer = XLSX.write(workbook, {
      bookType: "xlsx",
      type: "array",
      
    });
    const dataBlob = new Blob([excelBuffer], {
      type: "application/octet-stream",
    });
    saveAs(dataBlob, "data.xlsx");
  };


  return (
    <div className={"app"}>
      <header>
        <h1>Newspaper Seller's System</h1>
        <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} />
      </header>

      {tablesData.length > 0 &&
        tablesData.map((table, tableIndex) => (
          <div key={tableIndex}>
            <h2>Table {tableIndex + 1}</h2>
            {table.length > 0 && (
              <table>
                <thead>
                  <tr>
                    {Object.keys(table[0]).map((header, index) => (
                      <th key={index}>{header}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {table.map((row, rowIndex) => (
                    <tr key={rowIndex}>
                      {Object.values(row).map((value, index) => (
                        <td key={index}>{value}</td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
        ))}

      <Button
        sx={{ m: 2 }}
        onClick={() => {
          simulteData();
        }}
        disabled={fileUploaded ? false : true}
      >
        Generate Simulation Table
      </Button>

      {/* simulation TAble */}

      <TableContainer component={Paper}>
        <Button sx={{ float: "left", m: 2 }} onClick={exportToExcel}>
          <SaveIcon />
        </Button>
        <Table
          sx={{ minWidth: 900, textAlign: "center", border: "1px solid #ccc" }}
          aria-label="simple table"
        >
          <TableHead>
            <TableRow>
              <TableCell>Day</TableCell>
              <TableCell>R.D. for Type of Newsday</TableCell>
              <TableCell>Type of Newsday</TableCell>
              <TableCell>R.D. for Demand</TableCell>
              <TableCell>Demand</TableCell>
              <TableCell>Revenue from Sales</TableCell>
              <TableCell>Excess Demand</TableCell>
              <TableCell>Lost Profit from Excess Demand</TableCell>
              <TableCell>Number of Scrap Papers</TableCell>
              <TableCell>Salvage from Sale of Scrap</TableCell>
              <TableCell>Daily Profit</TableCell>
            </TableRow>
          </TableHead>
          <TableBody>
            {simulatedData.map((row) => (
              <TableRow
                key={row.day}
                sx={{ "&:last-child td, &:last-child th": { border: 0 } }}
              >
                <TableCell>{row.day}</TableCell>
                <TableCell>{row.RD_Type}</TableCell>
                <TableCell>{row.type}</TableCell>
                <TableCell>{row.RD_Demand}</TableCell>
                <TableCell>{row.demand}</TableCell>
                <TableCell>{row.revenueFromSales}</TableCell>
                <TableCell>{row.excessDemand}</TableCell>
                <TableCell>{row.lostProfitExcessDemand}</TableCell>
                <TableCell>{row.numberOfScrap}</TableCell>
                <TableCell>{row.salvage}</TableCell>
                <TableCell>{row.dailyProfit}</TableCell>
              </TableRow>
            ))}
          </TableBody>
        </Table>
      </TableContainer>
    </div>
  );
};

export default App;
