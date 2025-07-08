import React, { useState } from "react";
import "./index.scss";
import ArrowLeft from "../assets/images/ArrowLeft.png";
import UserIcon from "../assets/images/user.png";
import { useNavigate } from "react-router-dom";
import SignOutIcon from "../assets/images/SignOut.png";

const SignOutUser = () => {
const [loading, setLoading] = useState(false);
const navigate = useNavigate();
    const handleSignOut = () => {
        setLoading(true);
        setTimeout(() => {
          setLoading(false);
          navigate("/");
        }, 2000);
    }
    const handleSignBack = () => {
        setLoading(true);
        setTimeout(() => {
          setLoading(false);
          navigate("/exportExcel");
        }, 2000);
    }
  return (
    <div className="profile-container">
    {loading ? (
        <div className="main-loading-container" >
        <div className="loading-container">
          <div className="spinner"></div>
          <p className="loading-text">Loading...</p>
        </div>
        </div>
      ) : (
        <>
      <div className="back-button" onClick={handleSignBack}>
        <img src={ArrowLeft} alt="Back" />
        <span>Back</span>
      </div>

      <div className="profile-info">
        <img src={UserIcon} alt="User" className="profile-avatar" />
        <h2>Arlene McCoy</h2>
        <p>arlenemccoy123@gmail.com</p>
      </div>

      <div className="profile-details">
        <div className="detail-row">
          <span className="label">Imported from Excel</span>
          <span className="label_line" >12 times</span>
        </div>
        <div className="detail-row">
          <span className="label">Account creation date</span>
          <span className="label_line"  >2025-02-14</span>
        </div>
        <div className="detail-row">
          <span className="label">Previous Session</span>
          <span className="label_line"  >2025-02-16</span>
        </div>
      </div>

      <button className="signout-button" onClick={handleSignOut} >
        <img src={SignOutIcon} alt="Sign Out" />
        Sign Out
      </button>
      </>
    )}
    </div>
  );
};

export default SignOutUser;
