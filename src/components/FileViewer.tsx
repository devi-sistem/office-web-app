import React, { useState, useEffect } from 'react';
import mammoth from 'mammoth';

interface FileViewerProps {
  file: {
    name: string;
    type: 'word' | 'excel' | 'powerpoint' | 'unknown';
    handle: FileSystemFileHandle;
  };
  onClose: () => void;
}

const FileViewer: React.FC<FileViewerProps> = ({ file, onClose }) => {
  const [content, setContent] = useState<string>('');
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    const loadFileContent = async () => {
      try {
        if (file.type === 'word') {
          const fileData = await file.handle.getFile();
          const arrayBuffer = await fileData.arrayBuffer();
          const result = await mammoth.extractRawText({ arrayBuffer });
          setContent(result.value);
        } else {
          setContent('This file type is not supported for viewing yet.');
        }
      } catch (err) {
        console.error('Error reading file:', err);
        setError('Failed to read file content. Please try again.');
      } finally {
        setLoading(false);
      }
    };

    loadFileContent();
  }, [file]);

  if (loading) {
    return <div>Loading file content...</div>;
  }

  if (error) {
    return <div>Error: {error}</div>;
  }

  return (
    <div className="fixed inset-0 bg-gray-600 bg-opacity-50 overflow-y-auto h-full w-full">
      <div className="relative top-20 mx-auto p-5 border w-11/12 shadow-lg rounded-md bg-white">
        <div className="mt-3">
          <h3 className="text-lg leading-6 font-medium text-gray-900 mb-2">
            {file.name}
          </h3>
          <div className="mt-2 px-7 py-3 max-h-96 overflow-y-auto">
            {file.type === 'word' ? (
              <pre className="whitespace-pre-wrap font-sans text-sm text-gray-700">
                {content}
              </pre>
            ) : (
              <p className="text-sm text-gray-500">{content}</p>
            )}
          </div>
          <div className="items-center px-4 py-3">
            <button
              className="px-4 py-2 bg-blue-500 text-white text-base font-medium rounded-md w-full shadow-sm hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-300"
              onClick={onClose}
            >
              Close
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};

export default FileViewer;