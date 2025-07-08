import React, { useEffect, useState } from "react";
import "./HomeScreen.scss";
import DismissIcon from "../../../../assets/Dismiss.png";
import IconCheck from "../../../../assets/IconCheck.png";
import UserIcon from "../../../../assets/user.png";
import NeedHelpIcon from "../../../../assets/needHelp.png";
import Arrow_white_Import from "../../../../assets/Arrow_white_Import.png";
import UpdateIcon from "../../../../assets/update.png";
import MultiSelectDropdown from "../MultiSelectDropdown/MultiSelectDropdown";
import CustomTabs from "../Tabs/CustomTabs";
import ExpandableTable from "../ExpandableTable/ExpandableTable";
import InsertIcon from "../../../../assets/InsertIcon.png";
import Checkactive from "../../../../assets/checkactive.png";
import Checkhover from "../../../../assets/checkhover.png";
import { useNavigate, useParams } from "react-router-dom";

const Home = () => {
  const [expandedRow, setExpandedRow] = useState(null);
  const [checked, setChecked] = useState(false);
  const [loading, setLoading] = useState(false);
  const [file, setfile] = useState("true");
  const [Data, Setdata] = useState([]);
  const [selectedFileData, setSelectedFileData] = useState(null);
  const [rangeData, setRangeData] = useState(null);
  const [rangename, Setrangename] = useState("");
  const [valuesData, SetvaluesData] = useState([]);
  const [isUpdating, setIsUpdating] = useState(false);
  const [Selected, Setselected] = useState(false);
  const [UpdateTime, SetUpdateTime] = useState("");
  const [removeLoginPopup, Setremoveloginpopup] = useState(true);

  const { email } = useParams();
  const navigate = useNavigate();
  const currentdate = new Date();
  const datetime = currentdate.toLocaleString();

  useEffect(() => {
    getUserdata();

    setTimeout(() => {
      Setremoveloginpopup(false);
    }, 2000);
  }, []);

  useEffect(() => {
    SetUpdateTime(datetime);
  }, [datetime]);

  useEffect(() => {
    async function getdata() {
      const data = await fetchRangevalues(rangename);
      SetvaluesData(data);
    }

    getdata();
  }, [rangename]);

  const getUserdata = async () => {
    try {
      const res = await fetch(
        `https://us-central1-bbca-be.cloudfunctions.net/api/getFilename?email=${email}`
      );
      const data = await res.json();
      Setdata(data.filenames);
    } catch (error) {
      console.error("Error fetching user data:", error);
    }
  };

  const fetchRangevalues = (rangename) => {
    const storedData = localStorage.getItem("ExcelData");
    let data = [];
    if (storedData) {
      try {
        data = JSON.parse(storedData);
      } catch (error) {
        console.error("Invalid JSON in localStorage:", error);
      }
    }
    const result = data?.filter((d) => d.rangeName === rangename);
    return result || [];
  };

  const insertIntoWord = async () => {
    try {
      const freshData = await fetchRangevalues(rangename);
      console.log("values data", freshData);

      await Word.run(async (context) => {
        if (!freshData || freshData.length === 0 || !rangename) {
          console.warn("No data or rangename provided.");
          return;
        }

        const data = freshData[0];
        const body = context.document.body;

        if (data.type === "table" && data.html) {
          let html = data.html;

          if (checked) {
            html = html.replace(/style="([^"]*)"/g, (match, styleContent) => {
              const allowedStyles = styleContent
                .split(";")
                .map((s) => s.trim())
                .filter((s) => s.startsWith("border"))
                .join("; ");
              return allowedStyles ? `style="${allowedStyles}"` : "";
            });
          }

          const range = context.document.getSelection();
          range.insertHtml(html, Word.InsertLocation.replace);
          const cc = range.insertContentControl();
          cc.title = "Upload Table";
          cc.tag = rangename;

          await context.sync();
          return;
        }

        if (rangename.endsWith("_img") || data.type === "image") {
          const base64ImageString = data.value;
          const image = body.insertInlinePictureFromBase64(
            base64ImageString,
            Word.InsertLocation.end
          );
          image.select();
          const cc = image.insertContentControl();
          cc.title = "Uploaded Image";
          cc.tag = rangename;
          await context.sync();
          return;
        }

        if ((data.rowCount > 1 || data.columnCount > 1) && checked) {
          const values = data.value.split(",").map((s) => s.trim());
          const tblData = [];

          for (let i = 0; i < values.length; i += data.columnCount) {
            tblData.push(values.slice(i, i + data.columnCount));
          }

          const table = body.insertTable(
            data.rowCount,
            data.columnCount,
            Word.InsertLocation.end,
            tblData
          );
          const cc = table.insertContentControl();
          cc.tag = rangename;
          cc.title = `Range: ${rangename}`;
          await context.sync();
          return;
        }

        const range = context.document.getSelection();
        const contentControl = range.insertContentControl();
        contentControl.tag = rangename;
        contentControl.title = `Range: ${rangename}`;
        contentControl.insertText(data.value, Word.InsertLocation.replace);
        await context.sync();
      });
    } catch (error) {
      console.error("Failed to insert into Word:", error);
    }
  };

  const handleUpdate = async () => {
    setIsUpdating(true);
    try {
      const freshValues = await fetchRangevalues(rangename);
      SetvaluesData(freshValues);

      if (!freshValues || freshValues.length === 0) {
        console.warn("No data found for range:", rangename);
        return;
      }

      await Word.run(async (context) => {
        const contentControls = context.document.contentControls;
        contentControls.load("items/tag");
        await context.sync();

        let foundControl = false;

        for (let cc of contentControls.items) {
          if (cc.tag === rangename) {
            foundControl = true;
            const rangeValue = freshValues[0];
            if (rangeValue.rowCount > 1 || rangeValue.columnCount > 1) {
              const values = rangeValue.value.split(",").map((s) => s.trim());
              let tblData = [];

              console.log(rangeValue);
              for (let i = 0; i < values.length; i += rangeValue.columnCount) {
                tblData.push(values.slice(i, i + rangeValue.columnCount));
              }

              const ccRange = cc.getRange();
              const paragraph = ccRange.insertParagraph("", Word.InsertLocation.after);
              paragraph.load();
              await context.sync();

              cc.delete();
              await context.sync();

              const table = paragraph.insertTable(
                rangeValue.rowCount,
                rangeValue.columnCount,
                Word.InsertLocation.before,
                tblData
              );

              paragraph.delete();
              const newCC = table.insertContentControl();
              newCC.tag = rangename;
              newCC.title = `Range: ${rangename}`;

              console.log("Updated table content control with tag:", rangename);
            } else if (rangename.endsWith("_img")) {
              const ccRange = cc.getRange();
              const paragraph = ccRange.insertParagraph("", Word.InsertLocation.after);
              paragraph.load();
              await context.sync();

              cc.delete();
              await context.sync();

              const base64ImageString = rangeValue.value;
              console.log(base64ImageString);
              const image = paragraph.insertInlinePictureFromBase64(
                base64ImageString,
                Word.InsertLocation.before
              );

              paragraph.delete();
              const newCC = image.insertContentControl();
              newCC.tag = rangename;
              newCC.title = `Range: ${rangename}`;

              console.log("Updated image content control with tag:", rangename);
            } else {
              let newValue = rangeValue.value;
              cc.insertText(newValue, Word.InsertLocation.replace);
              console.log("Updated text content control with tag:", rangename);
            }
          }
        }

        if (!foundControl) {
          console.warn(`No content control with tag '${rangename}' found in the document.`);
        }

        await context.sync();
      });
    } catch (error) {
      console.error("Error updating Word document:", error);
    } finally {
      setIsUpdating(false);
    }
  };

  const handleChange = () => {
    setLoading(true);
    setTimeout(() => {
      setLoading(false);
      navigate("/exportExcel");
    }, 2000);
  };

  const toggleRow = (index) => {
    setExpandedRow(expandedRow === index ? null : index);
  };

  return (
    <div className="excel-import-container">
      {loading ? (
        <div className="main-loading-container">
          <div className="loading-container">
            <div className="spinner"></div>
            <p className="loading-text">Loading...</p>
          </div>
        </div>
      ) : (
        <>
          {removeLoginPopup && (
            <div className="success-message">
              <img src={IconCheck} alt="Success" />
              <span className="LoggedSuccessfully">Login Successfull</span>
              <img src={DismissIcon} alt="Dismiss" className="dismiss-icon" />
            </div>
          )}

          <div className="import-section">
            <h2>Get Excel Content</h2>
            <p>Select from the fields available below to Insert into the document</p>

            <button className="update-btn" onClick={handleUpdate} disabled={isUpdating}>
              {isUpdating ? (
                <div className="spinner small-spinner"></div>
              ) : (
                <>
                  <img src={UpdateIcon} alt="Update" />
                  Update
                </>
              )}
            </button>

            <p className="last-update">{UpdateTime}</p>
          </div>

          <div className="imported-items">
            <h2>Select Content</h2>
          </div>

          {/* <MultiSelectDropdown data={Data} setSelectedFileData={setSelectedFileData} /> */}
          <CustomTabs
            files={file}
            rangeData={rangeData}
            Setrangename={Setrangename}
            Setselected={Setselected}
          />

          <div className="button-container">
            <label className="checkbox-container">
              <input type="checkbox" checked={checked} onChange={() => setChecked(!checked)} />
              <img
                src={checked ? Checkactive : Checkhover}
                alt="checkbox"
                className="checkbox-icon"
              />
              Use destination formatting
            </label>

            <button
              className="insert-button"
              // disabled={valuesData.length === 0 || !Selected}
              onClick={insertIntoWord}
            >
              <img src={InsertIcon} alt="Insert Icon" />
              Insert
            </button>
          </div>

          <ExpandableTable
            source={UpdateIcon}
            headingText="Update All"
            headingfirst="Linked Fields "
            className=""
          />
        </>
      )}
    </div>
  );
};

export default Home;
