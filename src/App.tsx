import React from 'react';
import { BrowserRouter as Router, Route, Routes, Link } from 'react-router-dom';
import { MsalProvider } from "@azure/msal-react";
import { PublicClientApplication } from "@azure/msal-browser";
import { msalConfig } from "./authConfig";
import Home from './components/Home';
import Word from './components/Word';
import Excel from './components/Excel';
import PowerPoint from './components/PowerPoint';
import FileList from './components/FileList';

const msalInstance = new PublicClientApplication(msalConfig);

const App: React.FC = () => {
  return (
    <MsalProvider instance={msalInstance}>
      <Router>
        <div className="p-4">
          <nav className="mb-4">
            <ul className="flex space-x-4">
              <li><Link to="/" className="text-blue-500 hover:underline">Home</Link></li>
              <li><Link to="/files" className="text-blue-500 hover:underline">My Files</Link></li>
              <li><Link to="/word" className="text-blue-500 hover:underline">Word</Link></li>
              <li><Link to="/excel" className="text-blue-500 hover:underline">Excel</Link></li>
              <li><Link to="/powerpoint" className="text-blue-500 hover:underline">PowerPoint</Link></li>
            </ul>
          </nav>
          <Routes>
            <Route path="/" element={<Home />} />
            <Route path="/files" element={<FileList />} />
            <Route path="/word" element={<Word />} />
            <Route path="/excel" element={<Excel />} />
            <Route path="/powerpoint" element={<PowerPoint />} />
          </Routes>
        </div>
      </Router>
    </MsalProvider>
  );
};

export default App;