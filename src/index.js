import React from 'react';
import ReactDOM from 'react-dom/client';
import './index.css';
import App from './App';

// Mock the file system API for GitHub Pages (since it doesn't have fs access)
// This allows for testing without the actual file
window.fs = {
  readFile: async (filename) => {
    // When deployed, you'll need to include your Excel file in your public folder
    // and access it via fetch instead
    if (filename === 'reportstellarsoiree.xlsx') {
      try {
        const response = await fetch(`${process.env.PUBLIC_URL}/reportstellarsoiree.xlsx`);
        const arrayBuffer = await response.arrayBuffer();
        return new Uint8Array(arrayBuffer);
      } catch (error) {
        console.error('Error loading file:', error);
        throw error;
      }
    }
    throw new Error(`File ${filename} not found`);
  }
};

const root = ReactDOM.createRoot(document.getElementById('root'));
root.render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);