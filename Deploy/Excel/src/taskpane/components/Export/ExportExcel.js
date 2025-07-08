import React, { useEffect, useState } from "react";
import "./ExcelImport.scss";
import DismissIcon from "../assets/images/Dismiss.png";
import { jwtDecode } from "jwt-decode";
import IconCheck from "../assets/images/IconCheck.png";
import UserIcon from "../assets/images/user.png";
import NeedHelpIcon from "../assets/images/needHelp.png";
import Arrow_white_Import from "../assets/images/Arrow_white_Import.png";
import UpdateIcon from "../assets/images/update.png";
import TableSearch from "../assets/images/TableSearch.png";
import MultiSelectDropdown from "../MultiSelectDropdown/MultiSelectDropdown";
import CustomTabs from "../Tabs/CustomTabs";
import ExpandableTable from "../ExpandableTable/ExpandableTable";
import InsertIcon from "../assets/images/InsertIcon.png";
import ArrowExport from "../assets/images/ArrowExport.png";
import Checkactive from "../assets/images/checkactive.png";
import Checkhover from "../assets/images/checkhover.png";
import { useNavigate } from "react-router-dom";

const ExportExcel = () => {
  const [expandedRow, setExpandedRow] = useState(null);
  const [checked, setChecked] = useState(false);
  const [loading, setLoading] = useState(false);
  const [detectingRanges, setDetectingRanges] = useState(false); 
  const [Rangedata, SetRangedata] = useState([]);
  const [filteredRangedata, setFilteredRangedata] = useState([]);
  const [showpopup, SetShowpopup] = useState(false);
  const [isUpdating, setIsUpdating] = useState(false);
  const [prefixInput, setPrefixInput] = useState("");
  const navigate = useNavigate();

  const handleChange = () => {
    setLoading(true);
    setTimeout(() => {
      setLoading(false);
      navigate("/SignOutUser");
    }, 2000);
  };

  useEffect(() => {
    if (showpopup) {
      const timer = setTimeout(() => {
        SetShowpopup(false);
      }, 2000);
      
      return () => clearTimeout(timer);
    }
  }, [showpopup]);

  const filterRangesByPrefix = (ranges, prefix) => {
    if (!prefix.trim()) {
      return ranges;
    }

    const trimmedPrefix = prefix.trim().toLowerCase();

    return ranges.filter((range) => {
      const rangeName = range.RangeName.toLowerCase();
      return rangeName.includes(trimmedPrefix);
    });
  };

  const handlePrefixChange = (e) => {
    const value = e.target.value;
    setPrefixInput(value);

    const filtered = filterRangesByPrefix(Rangedata, value);
    setFilteredRangedata(filtered);
  };

  const extractNamedRangeAsHtml = async (rangeName, context) => {
    try {
      const namedRange = context.workbook.names.getItem(rangeName).getRange();
      namedRange.load(["text", "rowCount", "columnCount", "values"]);
      await context.sync();

      const rowCount = namedRange.rowCount;
      const columnCount = namedRange.columnCount;
      const textValues = namedRange.text;
      const values = namedRange.values;

      let html = `<table style="border-collapse: collapse;">`;

      for (let i = 0; i < rowCount; i++) {
        html += "<tr>";
        for (let j = 0; j < columnCount; j++) {
          const cell = namedRange.getCell(i, j);
          cell.format.load(["fill/color", "font/color", "font/bold", "font/size", "font/name"]);
          await context.sync();

          const value = textValues[i][j];
          const bgColor = cell.format.fill.color || "#ffffff";
          const fontColor = cell.format.font.color || "#000000";
          const fontSize = cell.format.font.size || 12;
          const fontName = cell.format.font.name || "Arial";
          const bold = cell.format.font.bold ? "bold" : "normal";

          html += `<td style="
            border: 1px solid #000;
            padding: 6px;
            background-color: ${bgColor};
            color: ${fontColor};
            font-size: ${fontSize}px;
            font-family: ${fontName};
            font-weight: ${bold};
          ">${value}</td>`;
        }
        html += "</tr>";
      }

      html += "</table>";

      return {
        html: html,
        value: values.flat().join(", "),
        rowCount: rowCount,
        columnCount: columnCount,
      };
    } catch (error) {
      console.error(`Error extracting HTML for range ${rangeName}:`, error);
      return null;
    }
  };

  const getRanges = async () => {
    try {
      setLoading(true); 
      Office.context.auth.getAccessTokenAsync(
        {
          allowConsentPrompt: true,
          allowSignInPrompt: true,
          forMSGraphAccess: true,
        },
        async (result) => {
          console.log("Token callback result:", result);

          if (result.status === "succeeded" && result.value) {
            const decodedToken = jwtDecode(result.value);
            const email = decodedToken.preferred_username;
            console.log("Email:", email);

            await Excel.run(async (context) => {
              const workbook = context.workbook;
              workbook.load("name");
              const sheets = workbook.worksheets;
              sheets.load("items/name");
              await context.sync();

              const fileName = workbook.name;

              const names = workbook.names;
              names.load("items/name,items/value");
              await context.sync();

              for (const sheet of sheets.items) {
                const worksheet = workbook.worksheets.getItem(sheet.name);

                for (const namedItem of names.items) {
                  try {
                    if (namedItem.name.startsWith(prefixInput)) {
                      let payload = {};

                      if (namedItem.name.endsWith("_img")) {
                        const range = worksheet.getRange(namedItem.value);
                        await context.sync();
                        const image = range.getImage();
                        await context.sync();

                        payload = {
                          user: email,
                          rangeName: namedItem.name,
                          value: image.value,
                          type: "image",
                          fileName,
                          sheetName: sheet.name,
                        };
                      } else {
                        const range = worksheet.getRange(namedItem.value);
                        range.load(["values", "rowCount", "columnCount"]);
                        await context.sync();

                        const isTable = range.rowCount > 1 || range.columnCount > 1;

                        if (isTable) {
                          const htmlData = await extractNamedRangeAsHtml(namedItem.name, context);

                          if (htmlData) {
                            payload = {
                              user: email,
                              rangeName: namedItem.name,
                              value: htmlData.value,
                              html: htmlData.html,
                              type: "table",
                              fileName,
                              sheetName: sheet.name,
                              rowCount: htmlData.rowCount,
                              columnCount: htmlData.columnCount,
                            };
                          } else {
                            const value = range.values.flat().join(", ");
                            payload = {
                              user: email,
                              rangeName: namedItem.name,
                              value: value,
                              type: "table",
                              fileName,
                              sheetName: sheet.name,
                              rowCount: range.rowCount,
                              columnCount: range.columnCount,
                            };
                          }
                        } else {
                          const value = range.values.flat().join(", ");
                          payload = {
                            user: email,
                            rangeName: namedItem.name,
                            value: value,
                            type: "text",
                            fileName,
                            sheetName: sheet.name,
                            rowCount: range.rowCount,
                            columnCount: range.columnCount,
                          };
                        }
                      }

                      await uploadData(payload);
                    }
                  } catch (error) {
                    console.error(`❌ Error processing named range '${namedItem.name}' in sheet '${sheet.name}':`, error);
                  }
                }

                const shapes = worksheet.shapes;
                shapes.load("items/name,type");
                await context.sync();

                const imageShapes = shapes.items.filter((shape) => {
                  if (shape.type !== "Image") return false;

                  if (!prefixInput.trim()) return true;

                  const trimmedPrefix = prefixInput.trim().toLowerCase();
                  const shapeName = shape.name.toLowerCase();
                  return shapeName.includes(trimmedPrefix);
                });

                console.log(`Exporting ${imageShapes.length} image shapes from sheet '${sheet.name}' matching prefix "${prefixInput}"`);

                for (const shape of imageShapes) {
                  try {
                    const image = shape.getAsImage(Excel.PictureFormat.png);
                    await context.sync();

                    const imagePayload = {
                      user: email,
                      rangeName: shape.name,
                      value: image.value,
                      type: "image",
                      fileName,
                      sheetName: sheet.name,
                    };

                    await uploadData(imagePayload);
                  } catch (error) {
                    console.error(`❌ Error extracting image '${shape.name}' in sheet '${sheet.name}':`, error);
                  }
                }
              }
              SetShowpopup(true);
              setLoading(false);
            });
          }
        }
      );
    } catch (error) {
      console.error("Error during Get Ranges:", error);
      setLoading(false); 
    }
  };

  const uploadData = async (payload) => {
    try {
      const existingData = JSON.parse(localStorage.getItem("ExcelData")) || [];

      const updatedData = existingData.map((item) => {
        if (item.rangeName === payload.rangeName && item.fileName === payload.fileName) {
          return { ...item, ...payload };
        }
        return item;
      });

      const isNew = !existingData.some(
        (item) => item.rangeName === payload.rangeName && item.fileName === payload.fileName
      );

      if (isNew) {
        updatedData.push(payload);
      }

      localStorage.setItem("ExcelData", JSON.stringify(updatedData));
    } catch (error) {
      console.error("Failed to store/update payload in localStorage:", error);
    }
  };

  const getDetectRange = async () => {
    try {
      setDetectingRanges(true); 
      await Excel.run(async (context) => {
        const workbook = context.workbook;
        workbook.load("name");
        await context.sync();

        const fileName = workbook.name;

        const sheets = workbook.worksheets;
        sheets.load("items/name");
        await context.sync();

        const names = workbook.names;
        names.load("items/name,items/value");
        await context.sync();

        const payload = [];

        for (const sheet of sheets.items) {
          const worksheet = workbook.worksheets.getItem(sheet.name);

          for (const namedItem of names.items) {
            try {
              if (namedItem.name.startsWith(prefixInput)) {
                const range = worksheet.getRange(namedItem.value);
                range.load(["values", "rowCount", "columnCount"]);
                await context.sync();

                const value = range.values.flat().join(", ");
                payload.push({
                  id: namedItem.name,
                  Sheet: sheet.name,
                  RangeName: namedItem.name,
                  value: value,
                  filename: fileName,
                  type: namedItem.name.endsWith("_img")
                    ? "Image"
                    : range.rowCount > 1 || range.columnCount > 1
                      ? "Table"
                      : "Text",
                });
              }
            } catch (error) {
              console.error(`❌ Skipping invalid or deleted range '${namedItem.name}' in sheet '${sheet.name}':`, error);
            }
          }

          const shapes = worksheet.shapes;
          shapes.load("items");
          await context.sync();

          const imageShapes = shapes.items.filter((shape) => shape.type === "Image");

          for (const shape of imageShapes) {
            try {
              shape.load("name");
              const image = shape.getAsImage(Excel.PictureFormat.png);
              await context.sync();

              payload.push({
                id: shape.name,
                Sheet: sheet.name,
                RangeName: shape.name,
                value: image.value,
                filename: fileName,
                type: "Image",
              });
            } catch (error) {
              console.error(`❌ Error extracting image '${shape.name}' in sheet '${sheet.name}':`, error);
            }
          }
        }

        SetRangedata(payload);
        const filtered = filterRangesByPrefix(payload, prefixInput);
        setFilteredRangedata(filtered);
        setDetectingRanges(false); 
      });
    } catch (error) {
      console.error("Error in getDetectRange:", error);
      setDetectingRanges(false); 
    }
  };

  const toggleRow = (index) => {
    setExpandedRow(expandedRow === index ? null : index);
  };

  return (
    <div className="excel-import-container">
      {(loading || detectingRanges) ? (
        <div className="main-loading-container">
          <div className="loading-container">
            <div className="spinner"></div>
            <p className="loading-text">
              {loading ? "Exporting data..." : "Detecting ranges..."}
            </p>
          </div>
        </div>
      ) : (
        <>
          {showpopup && (
            <div className="success-message">
              <img src={IconCheck} alt="Success" />
              <span className="LoggedSuccessfully">Data Exported Successfully</span>
              <img
                src={DismissIcon}
                alt="Dismiss"
                className="dismiss-icon"
                onClick={() => SetShowpopup(false)}
              />
            </div>
          )}

          <div className="import-section">
            <h2>Choose content to export</h2>
            <div className="prefix-main-div">
              <div className="prefix-div">
                <p className="prefix-div-p">Items Name Prefix</p>
                <div className="prefix-input-div">
                  <input
                    type="text"
                    className="prefix-input"
                    value={prefixInput}
                    onChange={handlePrefixChange}
                    placeholder="Enter Prefix"
                  />
                </div>
              </div>
              <div>
                {prefixInput && (
                  <p className="filter-info">
                    Showing {filteredRangedata.length} ranges containing "{prefixInput}"
                  </p>
                )}
              </div>
            </div>

            <button 
              className="update-btn" 
              onClick={getDetectRange}
              disabled={detectingRanges}
            >
              <img src={TableSearch} alt="Update" />
              {detectingRanges ? "Detecting..." : "Detect Ranges"}
            </button>
          </div>

          {filteredRangedata.length > 0 && (
            <ExpandableTable
              className="disabled_image"
              Rangedata={filteredRangedata}
              source=""
              headingfirst="Export List"
            />
          )}

          <button 
            className="insert-button" 
            onClick={getRanges}
            disabled={loading}
          >
            <img src={ArrowExport} alt="Insert Icon" />
            Export ({filteredRangedata.length} ranges)
          </button>
        </>
      )}
    </div>
  );
};

export default ExportExcel;