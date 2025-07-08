import React, { useState, useEffect, useRef } from "react";
import "./MultiSelectDropdown.scss";
import RadioBase from "../../../../assets/RadioBase.png";
import RadioColor from "../../../../assets/radioColor.png";
import BottomIcon from "../../../../assets/bottomicon.png";
import NeedHelp from "../../../../assets/needHelp.png";

const MultiSelectDropdown = ({ data, setSelectedFileData }) => {
  const [selectedFileName, setSelectedFileName] = useState("");
  const [isOpen, setIsOpen] = useState(false);
  const [hoveredItem, setHoveredItem] = useState(null);
  const [selectedFile, setSelectedFile] = useState(null);
  const [searchTerm, setSearchTerm] = useState("");
  const dropdownRef = useRef(null);

  const toggleDropdown = () => setIsOpen(!isOpen);

  const handleSelect = (fileObj, index) => {
    if (selectedFile && selectedFile.name === fileObj.Name && selectedFile.index === index) {
      setSelectedFile(null);
      setSelectedFileData("");
    } else {
      setSelectedFile({ name: fileObj.Name, index });
      setSelectedFileData(fileObj.Name);
    }
  };

  const toggleTooltip = (index) => {
    setHoveredItem(hoveredItem === index ? null : index);
  };

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

  const uniqueData = Array.from(
    new Map(
      (data ?? [])
        .filter((file) =>
          file.Name.toLowerCase().includes(searchTerm.toLowerCase())
        )
        .map((file) => [file.Name, file])
    ).values()
  );


  return (
    <div className="dropdown-container" ref={dropdownRef}>
      <div className="dropdown-header" onClick={toggleDropdown}>
        <span>{selectedFile?.name ? selectedFile.name.slice(0, 25) + "..." : "Select File"}</span>
        <img src={BottomIcon} alt="Dropdown Icon" />
      </div>

      {isOpen && (
        <div className="dropdown-list z-index_rt">
          <div className="dropdown-div">
            <input
              type="text"
              placeholder="Search by filename..."
              className="dropdown-search"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
            />
          </div>

          {uniqueData.length === 0 ? (
            <p className="no-results">No matching files</p>
          ) : (
            uniqueData.map((file, index) => {
              const isSelected =
                selectedFile &&
                selectedFile.name === file.Name &&
                selectedFile.index === index;

              return (
                <div key={index} className="dropdown-item">
                  <img
                    src={isSelected ? RadioColor : RadioBase}
                    alt="Radio"
                    onClick={() => handleSelect(file, index)}
                  />
                  <span onClick={() => handleSelect(file, index)}>{file.Name}</span>

                  <img
                    src={NeedHelp}
                    alt="Info"
                    className="help-icon"
                    onClick={() => toggleTooltip(index)}
                  />

                  {hoveredItem === index && (
                    <div className="tooltip">
                      <div className="tooltip-arrow"></div>
                      <div className="tooltip-content">
                        <p><strong>Table:</strong> {file.Name}</p>
                        <p><strong>Type:</strong> Table (Word-formatted)</p>
                      </div>
                    </div>
                  )}
                </div>
              );
            })
          )}
        </div>
      )}
    </div>
  );
};

export default MultiSelectDropdown;
