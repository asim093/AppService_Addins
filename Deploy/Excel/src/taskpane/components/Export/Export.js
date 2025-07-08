import React, { useState } from "react";
import "./ExcelImport.scss";

import DismissIcon from "../assets/images/Dismiss.png";
import IconCheck from "../assets/images/IconCheck.png";
import UserIcon from "../assets/images/user.png";
import NeedHelpIcon from "../assets/images/needHelp.png";
import ArrowRightIcon from "../assets/images/ArrowRight.png";
import Arrow_white_Import from "../assets/images/Arrow_white_Import.png";
import UpdateIcon from "../assets/images/update.png";
import MultiSelectDropdown from "../MultiSelectDropdown/MultiSelectDropdown";
import CustomTabs from "../Tabs/CustomTabs";
import ExpandableTable from "../ExpandableTable/ExpandableTable";
import InsertIcon from "../assets/images/InsertIcon.png";
import Checkactive from "../assets/images/checkactive.png";
import Checkhover from "../assets/images/checkhover.png";
import { useNavigate } from "react-router-dom";

const ExcelImport = () => {
    const [expandedRow, setExpandedRow] = useState(null);
    const [checked, setChecked] = useState(false);
    const [loading, setLoading] = useState(false);
    const [file, setfile] = useState('true');
    const navigate = useNavigate();
  const data = [
    {
      type: "Table",
      name: "Name 2",
      sheet: "Sheet 1",
      workspace: "file_name_1",
      lastUpdate: "14/02/2025, 15:29:31",
    },
    {
      type: "Table",
      name: "Name 3",
      sheet: "Sheet 2",
      workspace: "file_name_2",
      lastUpdate: "13/02/2025, 12:45:00",
    },
  ];
  
  const handleChange = () => {
    setLoading(true);
    setTimeout(() => {
      setLoading(false);
      navigate("/exportExcel");
    }, 2000);
  };
const HandleFile = () => {
  setfile(!file)
}
  const toggleRow = (index) => {
    setExpandedRow(expandedRow === index ? null : index);
  };

  return (
    <div className="excel-import-container">
      {loading ? (
        <div className="main-loading-container" >
        <div className="loading-container">
          <div className="spinner"></div>
          <p className="loading-text">Loading...</p>
        </div>
        </div>
      ) : (
        <>
      <div className="success-message">
        <img src={IconCheck} alt="Success" />
        <span className="LoggedSuccessfully" >Logged in successfully</span>
        <img src={DismissIcon} alt="Dismiss" className="dismiss-icon" />
      </div>

      {/* User Info */}
      <div className="user-info">
        <img src={UserIcon} alt="User" className="user-icon" />
        <div className="user-details">
          <p className="user-name m-minus-t">Arlene McCoy</p>
          <p className="user-role">Manage Account</p>
        </div>
        <img src={NeedHelpIcon} alt="Help" className="help-icon" />
      </div>

      {/* Import Excel Content */}
      <div className="import-section">
        <h2>Get Excel Content</h2>
        <p>First submit content on Excel in order to successfully import it to Word</p>
        <button className="update-btn">
          <img src={UpdateIcon} alt="Update" />
          Update
        </button>
        <p className="last-update">Last update: 14/02/2025, 15:29:31</p>
      </div>

      <div className="imported-items">
        <h2>Select Imported Items</h2>
      </div>
      <MultiSelectDropdown />
      <CustomTabs  files={file} />
      <div className="button-container">
      <label className="checkbox-container">
        <input 
          type="checkbox" 
          checked={checked} 
          onChange={() => setChecked(!checked)} 
        />
        <img 
          src={checked ? Checkactive : Checkhover} 
          alt="checkbox" 
          className="checkbox-icon"
        />
        Use destination formatting
      </label>

      <button className="insert-button" onClick={HandleFile} >
        <img src={InsertIcon} alt="Insert Icon" />
        Insert
      </button>
    </div>
      <ExpandableTable source={UpdateIcon} headingText='Update All' headingfirst='Linked fields ' className="" />
      </>
    )}
    </div>
  );
};

export default ExcelImport;
