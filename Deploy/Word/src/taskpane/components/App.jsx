import React, { useEffect } from "react";
import { MemoryRouter as Router, Routes, Route } from "react-router-dom";
import HomeScreen from "./HomeScreen/HomeScreen";
import Home from "./Home/Home";



function App() {
  return (
    <Router>
      <Routes>
        <Route path="/" element={<HomeScreen />} />
        <Route path="/Home/:email" element={<Home />} />

      </Routes>
    </Router>
  );
}

export default App;
