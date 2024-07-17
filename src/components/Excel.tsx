import React, { useState } from 'react';
import { useMsal } from "@azure/msal-react";
import { InteractionRequiredAuthError } from "@azure/msal-browser";
import { Client } from "@microsoft/microsoft-graph-client";
import { AuthCodeMSALBrowserAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/authCodeMsalBrowser";

const Excel: React.FC = () => {
  const { instance, accounts } = useMsal();
  const [fileName, setFileName] = useState('');
  const [spreadsheets, setSpreadsheets] = useState<string[]>([]); // For simulation

  const createSpreadsheet = async () => {
    if (!fileName) {
      alert('Please enter a file name');
      return;
    }

    if (accounts.length === 0) {
      alert('No user signed in. Please sign in first.');
      return;
    }

    try {
      // Try to acquire the token silently first
      const silentRequest = {
        scopes: ["Files.ReadWrite.All"],
        account: accounts[0]
      };

      let tokenResponse;
      try {
        tokenResponse = await instance.acquireTokenSilent(silentRequest);
      } catch (error) {
        if (error instanceof InteractionRequiredAuthError) {
          // fallback to interaction when silent call fails
          tokenResponse = await instance.acquireTokenPopup(silentRequest);
        } else {
          throw error;
        }
      }

      const authProvider = new AuthCodeMSALBrowserAuthenticationProvider(instance, {
        account: accounts[0],
        scopes: ["Files.ReadWrite.All"]
      });

      const graphClient = Client.initWithMiddleware({ authProvider });

      try {
        await graphClient.api('/me/drive/root:/' + fileName + '.xlsx:/content')
          .put(''); // Empty content creates a blank Excel file
        alert('Spreadsheet created successfully!');
      } catch (apiError: any) {
        if (apiError.message && apiError.message.includes("Tenant does not have a SPO license")) {
          // Simulate spreadsheet creation if SPO license is not available
          setSpreadsheets(prev => [...prev, `${fileName}.xlsx`]);
          alert('Spreadsheet created successfully (simulated)!');
        } else {
          throw apiError;
        }
      }
    } catch (error) {
      console.error('Error creating spreadsheet:', error);
      alert('Error creating spreadsheet. Please try again.');
    }
  };

  return (
    <div className="p-4">
      <h1 className="text-2xl font-bold mb-4">Excel Web App</h1>
      <input 
        type="text" 
        value={fileName} 
        onChange={(e) => setFileName(e.target.value)} 
        placeholder="Enter file name" 
        className="border p-2 mb-4 w-full"
      />
      <button 
        onClick={createSpreadsheet} 
        className="bg-green-500 hover:bg-green-700 text-white font-bold py-2 px-4 rounded"
      >
        Create Spreadsheet
      </button>
      {spreadsheets.length > 0 && (
        <div className="mt-4">
          <h2 className="text-xl font-bold">Created Spreadsheets (Simulated):</h2>
          <ul>
            {spreadsheets.map((sheet, index) => (
              <li key={index}>{sheet}</li>
            ))}
          </ul>
        </div>
      )}
    </div>
  );
};

export default Excel;