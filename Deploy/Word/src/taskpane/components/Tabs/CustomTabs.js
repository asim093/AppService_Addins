import React, { useState, useEffect, useRef } from "react";
import "./CustomTabs.scss";
import RadioBase from "../../../../assets/RadioBase.png";
import RadioColor from "../../../../assets/radioColor.png";
import NeedHelp from "../../../../assets/needHelp.png";

const tabs = ["All", "Tables", "Text", "Images"];

const CustomTabs = ({ files, Setrangename, Setselected }) => {
  const [activeTab, setActiveTab] = useState("All");
  const [isOpen, setIsOpen] = useState(false);
  const [selectedOption, setSelectedOption] = useState(null);
  const [hoveredItem, setHoveredItem] = useState(null);
  const dropdownRef = useRef(null);

  const [rangeData, setRangeData] = useState(JSON.parse(localStorage.getItem("ExcelData")));

 

  useEffect(() => {
    const handleClickOutside = (event) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target)) {
        setHoveredItem(null);
        setIsOpen(false);
      }
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  const handleSelect = (index) => {
    if (selectedOption === index) {
      setSelectedOption(null);
      Setselected(false);
    } else {
      Setselected(true);
      setSelectedOption(index);
      Setrangename(filteredData[index]?.rangeName || "");
    }
  };

  const toggleTooltip = (id) => {
    setHoveredItem(hoveredItem === id ? null : id);
  };

  const filteredData = rangeData?.filter((item) => {
    switch (activeTab) {
      case "Text":
        return item.type === "text";
      case "Images":
        return item.type === "image";
      case "Tables":
        return item.type === "table";
      default:
        return true;
    }
  });

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

      {/* Dropdown */}
      <div className="dropdown-container z_index" ref={dropdownRef}>
        <div className="dropdown-list bg_background_color">
          <p className="dropdown-action item_margin">
            {filteredData?.length ?? "0"} items Found
          </p>

          {filteredData?.map((range, index) => (
            <div key={index} className="dropdown-item">
              <img
                src={selectedOption === index ? RadioColor : RadioBase}
                alt="Radio"
                onClick={() => handleSelect(index)}
              />
              <span onClick={() => handleSelect(index)}>
                {range.rangeName}
              </span>

              {hoveredItem === index && (
                <div className="tooltip">
                  <div className="tooltip-arrow"></div>
                  <div className="tooltip-content">
                    <p><strong>Table:</strong> {range?.Name || "Unknown"}</p>
                    <p><strong>Type:</strong> {range?.type || "N/A"}</p>
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
