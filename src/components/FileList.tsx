import React, { useState, useEffect } from 'react';
import { useMsal } from "@azure/msal-react";
import FileViewer from './FileViewer';

interface LocalFile {
  id: string;
  name: string;
  type: 'word' | 'excel' | 'powerpoint' | 'unknown';
  handle: FileSystemFileHandle;
}

const FileList: React.FC = () => {
  const { accounts } = useMsal();
  const [files, setFiles] = useState<LocalFile[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [viewingFile, setViewingFile] = useState<LocalFile | null>(null);

  const openFilePicker = async () => {
    if ('showOpenFilePicker' in window) {
      try {
        setLoading(true);
        const fileHandles = await (window as any).showOpenFilePicker({
          multiple: true,
          types: [
            {
              description: 'Office Documents',
              accept: {
                'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx'],
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
                'application/vnd.openxmlformats-officedocument.presentationml.presentation': ['.pptx'],
              },
            },
          ],
        });

        const newFiles = await Promise.all(fileHandles.map(async (handle: FileSystemFileHandle, index: number) => {
          const file = await handle.getFile();
          let type: 'word' | 'excel' | 'powerpoint' | 'unknown' = 'unknown';
          if (file.name.endsWith('.docx')) type = 'word';
          else if (file.name.endsWith('.xlsx')) type = 'excel';
          else if (file.name.endsWith('.pptx')) type = 'powerpoint';

          return {
            id: index.toString(),
            name: file.name,
            type,
            handle,
          };
        }));

        setFiles(prevFiles => [...prevFiles, ...newFiles]);
      } catch (err) {
        console.error('Error opening file picker:', err);
        setError('Failed to open file picker. Please try again.');
      } finally {
        setLoading(false);
      }
    } else {
      setError('File System Access API is not supported in your browser.');
    }
  };

  const openFile = (file: LocalFile) => {
    setViewingFile(file);
  };

  if (loading) {
    return <div>Loading files...</div>;
  }

  if (error) {
    return <div>Error: {error}</div>;
  }

  return (
    <div className="p-4">
      <h2 className="text-xl font-bold mb-4">Your Local Files</h2>
      <button 
        onClick={openFilePicker}
        className="bg-blue-500 hover:bg-blue-700 text-white font-bold py-2 px-4 rounded mb-4"
      >
        Select Office Files
      </button>
      {files.length === 0 ? (
        <p>No files selected. Click the button above to select files.</p>
      ) : (
        <ul className="list-disc pl-5">
          {files.map((file) => (
            <li key={file.id} className="mb-2">
              <span className={`${file.type === 'word' ? 'text-blue-500' : file.type === 'excel' ? 'text-green-500' : file.type === 'powerpoint' ? 'text-red-500' : 'text-gray-500'}`}>
                {file.name}
              </span>
              {' - '}
              <button 
                className="text-gray-500 hover:text-gray-700 underline"
                onClick={() => openFile(file)}
              >
                Open
              </button>
            </li>
          ))}
        </ul>
      )}
      {viewingFile && (
        <FileViewer 
          file={viewingFile} 
          onClose={() => setViewingFile(null)} 
        />
      )}
    </div>
  );
};

export default FileList;