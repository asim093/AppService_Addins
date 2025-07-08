import React, { useEffect, useState } from "react";
import "./ExpandableTable.scss";
import updateIcon from "../../../../assets/update.png";
import documentSearchIcon from "../../../../assets/DocumentSearch.png"
import bottomIcon from "../../../../assets/bottomicon.png";
import tableIcon from "../../../../assets/Table.png";

const ExpandableTable = ({ source, headingText, headingfirst, ClassName }) => {
  const [documentRanges, SetdocumentRanges] = useState([])

  const toggleExpand = (index) => {
    SetdocumentRanges((prevData) =>
      prevData.map((item, i) =>
        i === index ? { ...item, expanded: !item.expanded } : item
      )
    );
  
  };
  


  useEffect(() => {
    async function getrangename() {
      const allDocumentRanges = await getAllRangeNames();

      if (allDocumentRanges) {
        SetdocumentRanges((prev) =>
          allDocumentRanges.map((newItem) => {
            const prevItem = prev.find((p) => p.name === newItem.name);
            return {
              ...newItem,
              expanded: prevItem?.expanded || false,
            };
          })
        );
      }
    }
    getrangename();
  }, [documentRanges]);




  const getAllRangeNames = async () => {
    try {
      return await Word.run(async (context) => {
        const contentControls = context.document.contentControls;
        contentControls.load("items/tag , items/name");
        await context.sync();
        const tags = contentControls.items.map(cc => cc.tag).filter(tag => tag);
        const uniqueTags = Array.from(new Set(tags));
        const tagObjects = uniqueTags.map(tag => ({
          name: tag,
          type: "Table",
          sheet: "Sheet 1",
          workspace: "Workspace 1",
          lastUpdate: 10 / 12 / 2024,
          expanded: false
        }));

        return tagObjects;
      });
    } catch (error) {
      console.error("Failed to get range names:", error);
      return [];
    }
  };

  return (
    <div className="">
      <div className="containerlink">
        <span className="text">{headingfirst}</span>
        <button className="updateButton" >
          <img src={source} alt="Update" className={`icon  ${ClassName}`} />
          <span className="updateText">{headingText}</span>
        </button>
      </div>

      <div className="header">
        <span className="headerText">Type</span>
        <span className="headerText">Name</span>
        <span className="headerText"></span>
      </div>


      {documentRanges?.map((item, index) => (
        <div key={index}>
          <div className="row">
            <span className="cell">text</span>
            <span className="cell">{item?.name}</span>
            <div className="iconContainer">
              <button onClick={() => toggleExpand(index)}>
             
              </button>
              <button onClick={() => toggleExpand(index)}>
                <img
                  src={bottomIcon}
                  alt="Toggle"
                  className={`bottomIcon ${item?.expanded ? "arrowRotated" : ""}`}
                />
              </button>
            </div>
          </div>

          {item?.expanded && (
            <div className="expandedContainer">
              <img src={tableIcon} alt="Table" className="tableIcon" />
              <p className="infoText"><span></span>{item.type}t</p>
              <p className="infoText"><span>Sheet:</span> {item.sheet}</p>
              <p className="infoText"><span>Name:</span> {item.name}</p>
            </div>
          )}
        </div>
      ))}
    </div>
  );
};

export default ExpandableTable;
