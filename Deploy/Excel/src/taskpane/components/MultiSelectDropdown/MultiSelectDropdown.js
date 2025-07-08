import React, { useState, useEffect, useRef } from "react";
import "./MultiSelectDropdown.scss";
import RadioBase from "../assets/images/RadioBase.png";
import RadioColor from "../assets/images/radioColor.png";
import BottomIcon from "../assets/images/bottomicon.png";
import NeedHelp from "../assets/images/needHelp.png";

const options = [
  { id: 1, name: "Name 1", info: "Excel range size: 3 rows and 2 columns." },
  { id: 2, name: "Name 2", info: "Excel range size: 5 rows and 4 columns." },
  { id: 3, name: "Name 3", info: "Excel range size: 2 rows and 3 columns." },
  { id: 4, name: "Name 4", info: "Excel range size: 6 rows and 5 columns." },
];

const MultiSelectDropdown = () => {
  const [selectedOptions, setSelectedOptions] = useState([]);
  const [isOpen, setIsOpen] = useState(false);
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
    <div className="dropdown-container" ref={dropdownRef}>
      <div className="dropdown-header" onClick={toggleDropdown}>
        <span>Sheet 1</span>
        <img src={BottomIcon} alt="Dropdown Icon" />
      </div>

      {isOpen && (
        <div className="dropdown-list z-index_rt">
          <p className="dropdown-action">Action</p>
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
      )}
    </div>
  );
};

export default MultiSelectDropdown;
