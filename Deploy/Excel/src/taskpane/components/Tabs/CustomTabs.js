import React, { useState, useEffect, useRef } from "react";
import "./CustomTabs.scss";
import SearchIcon from "../assets/images/searchIcon.png"; // Update path accordingly
import RadioBase from "../assets/images/RadioBase.png";
import RadioColor from "../assets/images/radioColor.png";
import BottomIcon from "../assets/images/bottomicon.png";
import NeedHelp from "../assets/images/needHelp.png";
const tabs = ["All", "Tables", "Text", "Images"];
const options = [
    { id: 1, name: "Name 1", info: "Excel range size: 3 rows and 2 columns." },
    { id: 2, name: "Name 2", info: "Excel range size: 5 rows and 4 columns." },
    { id: 3, name: "Name 3", info: "Excel range size: 2 rows and 3 columns." },
    { id: 4, name: "Name 4", info: "Excel range size: 6 rows and 5 columns." },
  ];
  
  
const tabData = {
  All: "This is the All tab. It shows everything.",
  Tables: "Tables data: Table 1, Table 2, Table 3...",
  Text: "Text data: This is some sample text content.",
  Images: "Images: Image1.png, Image2.png, Image3.png...",
};

const CustomTabs = (files) => {
  console.log(files , "files")
  const [activeTab, setActiveTab] = useState("All");
  const [isOpen, setIsOpen] = useState(false);
  const [selectedOptions, setSelectedOptions] = useState([]);
const [hoveredItem, setHoveredItem] = useState(null);
  const dropdownRef = useRef(null);

  const toggleDropdown = () => setIsOpen(!isOpen);

  const handleSelect = (id) => {
    setSelectedOptions((prevSelected) =>
      prevSelected.includes(id)
        ? prevSelected.filter((item) => item !== id)
        : [...prevSelected, id]
    );
  };

  const toggleTooltip = (id) => {
    setHoveredItem(hoveredItem === id ? null : id);
  };

  // Close dropdown and tooltip when clicking outside
  useEffect(() => {
    const handleClickOutside = (event) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target)) {
        setHoveredItem(null);
        setIsOpen(false);
      }
    };

    document.addEventListener("mousedown", handleClickOutside);
    return () => {
      document.removeEventListener("mousedown", handleClickOutside);
    };
  }, []);
  return (
    <div className="custom-tabs">
      <div className="tabs">
        {tabs.map((tab) => (
          <div
            key={tab}
            className={`tab ${activeTab === tab ? "active" : ""}`}
            onClick={() => setActiveTab(tab)}
          >
            {tab}
          </div>
        ))}
      </div>
      {files.files ?
      <div className="search-bar">
        <img src={SearchIcon} alt="Search" className="search-icon" />
       
        <input type="text" placeholder="Search" />
        
        
        
      </div>
      :
      <div className="mt-tr" >
      <input type="file" placeholder="Search " /> 
      </div>
      }

      <div className="tab-content">{tabData[activeTab]}</div>
      <div className="dropdown-container z_index" >
       <div className="dropdown-list bg_background_color">
                <p className="dropdown-action item_margin">4 items found</p>
                {options.map((option) => (
                  <div key={option.id} className="dropdown-item">
                    <img
                      src={selectedOptions.includes(option.id) ? RadioColor : RadioBase}
                      alt="Radio"
                      onClick={() => handleSelect(option.id)}
                    />
                    <span onClick={() => handleSelect(option.id)}>{option.name}</span>
      
                    <img
                      src={NeedHelp}
                      alt="Info"
                      className="help-icon"
                      onClick={() => toggleTooltip(option.id)}
                    />
      
                    {hoveredItem === option.id && (
                      <div className="tooltip">
                        <div className="tooltip-arrow"></div>
                        <div className="tooltip-content">
                          <p><strong>Table:</strong> {option.name}</p>
                          <p><strong>Type:</strong> Table (Word-formatted)</p>
                          <p>{option.info}</p>
                        </div>
                      </div>
                    )}
                  </div>
                ))}
              </div>
              </div>
    </div>
  );
};

export default CustomTabs;
