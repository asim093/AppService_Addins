import React, { useState, useEffect } from "react";
import { PublicClientApplication } from "@azure/msal-browser";
import { jwtDecode } from "jwt-decode";
import "./HomeScreen.scss";
import logo from "../../../../assets/logoword.png";
import logError from "../../../../assets/LogError.png";
import dismiss from "../../../../assets/Dismiss.png";
import microsoftTri from "../../../../assets/MicrosoftTri.png";
import needHelp from "../../../../assets/needHelp.png";
import { useNavigate } from "react-router-dom";

const checkEmail = async (email) => {
  try {
    const response = await fetch(
      `https://us-central1-bbca-be.cloudfunctions.net/api/check-email?email=${email}`
    );
    const data = await response.json();
    if (data.exists) {
      return true;
    } else {
      return false;
    }
  } catch (apiError) {
    console.error("API call failed:", apiError);
    return false;
  }
};

const HomeScreen = () => {
  const navigate = useNavigate();
  const [officeReady, setOfficeReady] = useState(false);
  const [error, setError] = useState("");
  const [showError, setShowError] = useState(false);
  const [loading, setLoading] = useState(false);
  const [email, Setemail] = useState("");

  useEffect(() => {
    const initializeOffice = async () => {
      try {
        await new Promise((resolve) => {
          Office.onReady(() => {
            console.log("Office.js fully initialized");
            setOfficeReady(true);
            resolve();
          });
        });
      } catch (initError) {
        console.error("Office init failed:", initError);
        setError("Failed to initialize Office context");
      }
    };

    if (window.Office) {
      initializeOffice();
    } else {
      setError("Office.js not loaded - check your manifest source URLs");
    }
  }, []);

  if (error) {
    return (
      <div style={{ color: "red", padding: "20px" }}>
        CRITICAL ERROR: {error}
        <button onClick={() => window.location.reload()}>Reload Add-in</button>
      </div>
    );
  }

  Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
      console.log("Office is ready");
    }
  });

  const CLIENT_ID = "ab1349c6-78b8-4824-800b-066ea1c49997";
  const AUTHORITY = "https://login.microsoftonline.com/common";

  const msalInstance = new PublicClientApplication({
    auth: {
      clientId: CLIENT_ID,
      authority: AUTHORITY,
      redirectUri: window.location.origin,
    },
  });

  useEffect(() => {
    const login = async () => {
      try {
        const response = await msalInstance.loginPopup({
          scopes: ["openid", "email", "profile"],
        });

        console.log("ID Token:", response.idToken);

        // Decode the token to extract user email
        const decodedToken = jwtDecode(response.idToken);
        console.log("Decoded Token:", decodedToken);

        // setEmail(decodedToken.email || decodedToken.upn);
      } catch (error) {
        console.error("Login failed:", error);
      }
    };

    login();
  }, []);

  const handleLogin = () => {
    console.log(`Retrieving User Email`);
    if (!officeReady) return;
    setLoading(true);

    try {
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
            console.log("Decoded token:", decodedToken);
            console.log("Email:", decodedToken.preferred_username);
            Setemail(decodedToken.preferred_username);
            const emailCheck = await checkEmail(decodedToken.preferred_username);
            console.log("Email Check:", emailCheck);
            setLoading(false);
            if (emailCheck) {
              setShowError(false);
              console.log(decodedToken.preferred_username);
              navigate(`/Home/${decodedToken.preferred_username}`);
            } else {
              setShowError(true);
            }
          } else {
            console.error("Failed to get token:", result);
            setError(`Token retrieval failed: ${result.error.message}`);
          }
        }
      );
    } catch (authError) {
      console.error("Auth error details:", JSON.stringify(authError, null, 2));
      setError(`Authentication failed: ${authError.message}`);
    }
  };

  return (
    <div className="container">
      {loading ? (
        <div className="loading-container">
          <div className="spinner"></div>
          <p className="loading-text">Loading...</p>
        </div>
      ) : (
        <>
          {showError && (
            <div className="error-container">
              <div className="error-content">
                <img src={logError} alt="error icon" className="error-icon" />
                <div className="error-text-container">
                  <p className="error-title">Log In Error</p>
                  <p className="error-message">
                    This account is not in our system. For more questions contact the administrator
                    via the link below.
                  </p>
                </div>
                <button className="dismiss-button" onClick={() => setShowError(false)}>
                  <img src={dismiss} alt="dismiss icon" className="dismiss-icon" />
                </button>
              </div>
            </div>
          )}

          <div className="main-content">
            <img src={logo} alt="logo" className="logo" />
            <h1 className="welcome-text">Welcome!</h1>
            <p className="sub-text">Export from Excel to Word with ease.</p>
            <button className="login-button" onClick={handleLogin}>
              <img src={microsoftTri} alt="Microsoft icon" className="microsoft-icon" />
              <span>Sign In With Microsoft</span>
            </button>
          </div>

          <a
            href="mailto:markp@bbcre.ca"
            className="help-container"
            style={{ display: "flex", alignItems: "center", textDecoration: "none" }}
          >
            <img src={needHelp} alt="help icon" className="help-icon" />
            <div>
              <p className="help-title">Need Help?</p>
              <p className="help-text">Contact Administrator</p>
            </div>
          </a>
        </>
      )}
    </div>
  );
};

export default HomeScreen;
