import React, { useState } from 'react';
import { useMsal } from "@azure/msal-react";

const Word: React.FC = () => {
  const { accounts } = useMsal();
  const [content, setContent] = useState('');
  const [fileName, setFileName] = useState('');
  const [documents, setDocuments] = useState<string[]>([]);

  const createDocument = async () => {
    if (!fileName) {
      alert('Please enter a file name');
      return;
    }

    if (accounts.length === 0) {
      alert('No user signed in. Please sign in first.');
      return;
    }

    try {
      // Simulate document creation
      setDocuments(prev => [...prev, `${fileName}.docx`]);
      alert('Document created successfully (simulated)!');
    } catch (error) {
      console.error('Error creating document:', error);
      alert('Error creating document. Please try again.');
    }
  };

  return (
    <div className="p-4">
      <h1 className="text-2xl font-bold mb-4">Word Web App (Simulated)</h1>
      <input 
        type="text" 
        value={fileName} 
        onChange={(e) => setFileName(e.target.value)} 
        placeholder="Enter file name" 
        className="border p-2 mb-4 w-full"
      />
      <textarea 
        value={content} 
        onChange={(e) => setContent(e.target.value)} 
        placeholder="Enter your document content here" 
        className="border p-2 w-full h-64 mb-4"
      />
      <button 
        onClick={createDocument} 
        className="bg-blue-500 hover:bg-blue-700 text-white font-bold py-2 px-4 rounded"
      >
        Create Document
      </button>
      <div className="mt-4">
        <h2 className="text-xl font-bold">Created Documents:</h2>
        <ul>
          {documents.map((doc, index) => (
            <li key={index}>{doc}</li>
          ))}
        </ul>
      </div>
    </div>
  );
};

export default Word;