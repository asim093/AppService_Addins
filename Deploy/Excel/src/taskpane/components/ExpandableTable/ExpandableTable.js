import React, { useEffect, useState } from "react";
import "./ExpandableTable.scss";
import updateIcon from "../assets/images/update.png";
import documentSearchIcon from "../assets/images/DocumentSearch.png";
import bottomIcon from "../assets/images/bottomicon.png";
import tableIcon from "../assets/images/Table.png";



const ExpandableTable = ({ source, headingText, Rangedata, headingfirst, ClassName }) => {

  const [data, setData] = useState();
  useEffect(() => {
    setData(Rangedata)
    console.log(Rangedata)
  }, [Rangedata])

  const toggleExpand = (id) => {
    setData((prevData) =>
      prevData.map((item) =>
        item.id === id ? { ...item, expanded: !item.expanded } : item
      )
    );
  };



  return (
    <div className="">
      <div className="containerlink">
        <span className="text">{headingfirst}</span>
        
      </div>

      <div className="header">
        <span className="headerText">Type</span>
        <span className="headerText">Name</span>
        <span className="headerText"></span>
      </div>

      {data?.map((item) => {
        const isImageType = item.RangeName?.trim().split(/\s+/).slice(-3).every(word => word.endsWith("_img"));

        return (
          <div key={item.id}>
            <div className="row">
              <span className="cell">{item?.type}</span>
              <span className="cell">{item.RangeName}</span>
              <div className="iconContainer">
                <button onClick={() => toggleExpand(item.id)}>
                </button>
                <button onClick={() => toggleExpand(item.id)}>
                  <img
                    src={bottomIcon}
                    alt="Toggle"
                    className={`bottomIcon ${item.expanded ? "arrowRotated" : ""}`}
                  />
                </button>
              </div>
            </div>

            {item.expanded && (
              <div className="expandedContainer">
                <img src={tableIcon} alt="Table" className="tableIcon" />
                <p className="infoText"><span>Type:</span> {item.type}</p>
                <p className="infoText"><span>Range:</span> {item.Sheet}</p>
                <p className="infoText"><span>Name:</span> {item.RangeName}</p>
              </div>
            )}
          </div>
        );
      })}

    </div>
  );
};

export default ExpandableTable;
