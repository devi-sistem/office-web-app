import React, { useState } from 'react';
import { useMsal } from "@azure/msal-react";
import { Client } from "@microsoft/microsoft-graph-client";
import { AuthCodeMSALBrowserAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/authCodeMsalBrowser";

const PowerPoint: React.FC = () => {
  const { instance, accounts } = useMsal();
  const [fileName, setFileName] = useState('');

  const createPresentation = async () => {
    if (!fileName) {
      alert('Please enter a file name');
      return;
    }

    const authProvider = new AuthCodeMSALBrowserAuthenticationProvider(instance, { account: accounts[0], scopes: ["Files.ReadWrite"] });
    const graphClient = Client.initWithMiddleware({ authProvider });

    try {
      await graphClient.api('/me/drive/root:/' + fileName + '.pptx:/content')
        .put(''); // Empty content creates a blank PowerPoint file
      alert('Presentation created successfully!');
    } catch (error) {
      console.error('Error creating presentation:', error);
      alert('Error creating presentation. Please try again.');
    }
  };

  return (
    <div className="p-4">
      <h1 className="text-2xl font-bold mb-4">PowerPoint Web App</h1>
      <input 
        type="text" 
        value={fileName} 
        onChange={(e) => setFileName(e.target.value)} 
        placeholder="Enter file name" 
        className="border p-2 mb-4 w-full"
      />
      <button 
        onClick={createPresentation} 
        className="bg-red-500 hover:bg-red-700 text-white font-bold py-2 px-4 rounded"
      >
        Create Presentation
      </button>
    </div>
  );
};

export default PowerPoint;