import React, { useState } from "react";
import styled from "styled-components";
import { read, utils, writeFile } from "xlsx";

const FilePicker = () => {
  const [file, setFile] = useState(null);
  const [input1, setInput1] = useState("");
  const [input2, setInput2] = useState("");
  const [excelData, setExcelData] = useState([]);
  const [resArr, setResArr] = useState([]);

  const handleFileChange = (e) => {
    const selectedFile = e.target.files[0];
    if (
      selectedFile &&
      selectedFile.type ===
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    ) {
      setFile(selectedFile);
      const reader = new FileReader();

      reader.onload = (event) => {
        const arrayBuffer = event.target.result;
        const workbook = read(arrayBuffer, { type: "array" });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = utils.sheet_to_json(worksheet, { header: 1 });
        setExcelData(jsonData);
      };

      reader.readAsArrayBuffer(selectedFile);
    } else {
      alert("Please select an .xlsx file");
      e.target.value = null;
    }
  };

  const handleAdd = (e) => {
    e.preventDefault();

    for (let x of excelData) {
      if (Number(input1) === x[0]) {
        setResArr((prevResArr) => [
          ...prevResArr,
          {
            Serial_No: x[0],
            Service_Number: x[1],
            Description: x[2],
            Quantity: input2,
            Unit_Price: x[6],
            Total_Price: input2 * x[6],
          },
        ]);
      }
    }
    setInput1("");
    setInput2("");
  };

  const handleSubmit = (e) => {
    e.preventDefault();
    const worksheet = utils.json_to_sheet(resArr);
    const workbook = utils.book_new();
    utils.book_append_sheet(workbook, worksheet, "Valuation");
    writeFile(workbook, "Valuation.xlsx", { compression: true });
  };

  return (
    <Container>
      <Form>
        <Label>Pick an Excel File (.xlsx)</Label>
        <FileInput type="file" accept=".xlsx" onChange={handleFileChange} />

        <Label>Serial Number</Label>
        <Input
          type="text"
          value={input1}
          onChange={(e) => setInput1(e.target.value)}
          placeholder="Enter Serial Number"
        />

        <Label>Quantity</Label>
        <Input
          type="text"
          value={input2}
          onChange={(e) => setInput2(e.target.value)}
          placeholder="Enter quantity"
        />
        <Button type="submit" onClick={handleAdd}>
          Add
        </Button>
        <Button type="submit" onClick={handleSubmit}>
          Submit
        </Button>
      </Form>
    </Container>
  );
};

const Container = styled.div`
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  height: 100vh;
  background-color: #f0f4f8;
`;

const Form = styled.form`
  background-color: #fff;
  padding: 20px 20px;
  border-radius: 8px;
  box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
  display: flex;
  flex-direction: column;
  gap: 15px;
  width: 300px;
`;

const Label = styled.label`
  font-size: 16px;
  color: #333;
`;

const Input = styled.input`
  padding: 10px 10px;
  border: 1px solid #ccc;
  border-radius: 4px;
  font-size: 14px;
`;

const FileInput = styled.input`
  padding: 10px 10px;
  border: 1px solid #ccc;
  border-radius: 4px;
  font-size: 14px;
`;

const Button = styled.button`
  background-color: #007bff;
  color: #fff;
  border: none;
  padding: 10px;
  border-radius: 4px;
  font-size: 16px;
  cursor: pointer;
  &:hover {
    background-color: #0056b3;
  }
`;

export default FilePicker;
