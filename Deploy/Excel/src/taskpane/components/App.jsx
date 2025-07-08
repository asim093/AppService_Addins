import React, { useEffect } from "react";
import { MemoryRouter as Router, Routes, Route } from "react-router-dom";
import HomeScreen from "./Home/HomeScreen";
import "./App.css";
import ExcelImport from "./Export/Export";
import ExportExcel from "./Export/ExportExcel";
import SignOutUser from "./SignOut/SignOutUser";


function App() {
  return (
    <Router>
      <Routes>
        <Route path="/" element={<HomeScreen />} />
        <Route path="/export" element={<ExcelImport />} />
        <Route path="/exportExcel" element={<ExportExcel />} />
        <Route path="/SignOutUser" element={<SignOutUser />} />
      </Routes>
    </Router>
  );
}

export default App;
