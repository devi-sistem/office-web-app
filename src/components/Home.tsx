import React from 'react';
import { Link } from 'react-router-dom';
import { useMsal } from "@azure/msal-react";
import { loginRequest } from "../authConfig";

const Home: React.FC = () => {
  const { instance, accounts } = useMsal();

  const handleLogin = async () => {
    try {
      await instance.loginPopup(loginRequest);
    } catch (error) {
      console.error("Login failed", error);
    }
  };

  return (
    <div className="p-4">
      <h1 className="text-2xl font-bold mb-4">Local Office Web App</h1>
      {accounts.length === 0 ? (
        <button onClick={handleLogin} className="bg-blue-500 hover:bg-blue-700 text-white font-bold py-2 px-4 rounded">
          Sign In
        </button>
      ) : (
        <div>
          <p className="mb-4">Welcome, {accounts[0].name}!</p>
          <Link to="/files" className="bg-green-500 hover:bg-green-700 text-white font-bold py-2 px-4 rounded">
            Select and View Local Files
          </Link>
        </div>
      )}
    </div>
  );
};

export default Home;