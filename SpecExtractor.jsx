import React, { useState } from 'react';
import { Upload, CheckCircle, AlertCircle, Loader, Download, X } from 'lucide-react';

const BACKEND_URL = process.env.REACT_APP_BACKEND_URL || 'http://localhost:8000';

export default function SpecExtractor() {
  const [file, setFile] = useState(null);
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState(null);
  const [error, setError] = useState(null);
  const [dragActive, setDragActive] = useState(false);

  const handleDrag = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(e.type === 'dragenter' || e.type === 'dragover');
  };

  const handleDrop = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      setFile(e.dataTransfer.files[0]);
    }
  };

  const handleFileChange = (e) => {
    if (e.target.files && e.target.files[0]) {
      setFile(e.target.files[0]);
    }
  };

  const handleExtract = async () => {
    if (!file) {
      setError('Select an image first');
      return;
    }

    setLoading(true);
    setError(null);
    setResult(null);

    const formData = new FormData();
    formData.append('file', file);

    try {
      const response = await fetch(`${BACKEND_URL}/extract`, {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.detail || 'Extraction failed');
      }

      const data = await response.json();
      setResult(data);
    } catch (err) {
      setError(err.message || 'Network error');
    } finally {
      setLoading(false);
    }
  };

  const handleDownload = () => {
    if (result?.excel_url) {
      const url = `${BACKEND_URL}${result.excel_url}`;
      const a = document.createElement('a');
      a.href = url;
      a.download = result.excel_filename || 'spec.xlsx';
      a.click();
    }
  };

  const handleReset = () => {
    setFile(null);
    setResult(null);
    setError(null);
    setLoading(false);
  };

  return (
    <div className="min-h-screen bg-gradient-to-b from-slate-50 to-slate-100">
      {/* Header */}
      <div className="bg-slate-900 text-white py-8 px-4 shadow-lg">
        <div className="max-w-2xl mx-auto">
          <h1 className="text-3xl font-black tracking-tight mb-2">
            Garment Spec OCR
          </h1>
          <p className="text-slate-300 text-sm">
            Extract measurement specs from photos. Powered by Mistral Vision.
          </p>
        </div>
      </div>

      {/* Main Content */}
      <div className="max-w-2xl mx-auto px-4 py-12">
        
        {/* Upload Section */}
        {!result && (
          <div className="space-y-6">
            <div
              onDragEnter={handleDrag}
              onDragLeave={handleDrag}
              onDragOver={handleDrag}
              onDrop={handleDrop}
              className={`relative border-2 border-dashed rounded-xl p-8 transition-all ${
                dragActive
                  ? 'border-blue-500 bg-blue-50'
                  : 'border-slate-300 bg-white hover:border-slate-400'
              }`}
            >
              <input
                type="file"
                id="file-input"
                accept="image/*"
                onChange={handleFileChange}
                className="hidden"
              />

              <label
                htmlFor="file-input"
                className="flex flex-col items-center justify-center cursor-pointer gap-3"
              >
                {file ? (
                  <>
                    <CheckCircle size={48} className="text-green-600" strokeWidth={1.5} />
                    <p className="font-semibold text-slate-900">
                      {file.name}
                    </p>
                    <p className="text-sm text-slate-500">
                      {(file.size / 1024).toFixed(1)} KB
                    </p>
                  </>
                ) : (
                  <>
                    <Upload size={48} className="text-slate-400" strokeWidth={1.5} />
                    <div className="text-center">
                      <p className="font-semibold text-slate-900">
                        Drop image here or click to select
                      </p>
                      <p className="text-sm text-slate-500 mt-1">
                        JPG, PNG, BMP, WebP, TIFF
                      </p>
                    </div>
                  </>
                )}
              </label>
            </div>

            {/* Error Message */}
            {error && (
              <div className="bg-red-50 border border-red-200 rounded-lg p-4 flex gap-3">
                <AlertCircle size={20} className="text-red-600 flex-shrink-0 mt-0.5" />
                <div>
                  <p className="font-semibold text-red-900">Error</p>
                  <p className="text-sm text-red-700">{error}</p>
                </div>
              </div>
            )}

            {/* Action Buttons */}
            <div className="flex gap-3">
              <button
                onClick={handleExtract}
                disabled={!file || loading}
                className={`flex-1 py-3 px-4 rounded-lg font-semibold flex items-center justify-center gap-2 transition-all ${
                  file && !loading
                    ? 'bg-blue-600 text-white hover:bg-blue-700 active:scale-95'
                    : 'bg-slate-300 text-slate-500 cursor-not-allowed'
                }`}
              >
                {loading ? (
                  <>
                    <Loader size={20} className="animate-spin" />
                    Processing...
                  </>
                ) : (
                  <>
                    <Upload size={20} />
                    Extract Specs
                  </>
                )}
              </button>

              {file && (
                <button
                  onClick={() => setFile(null)}
                  disabled={loading}
                  className="px-4 py-3 rounded-lg font-semibold text-slate-900 bg-slate-200 hover:bg-slate-300 transition-all disabled:opacity-50"
                >
                  <X size={20} />
                </button>
              )}
            </div>
          </div>
        )}

        {/* Results Section */}
        {result && (
          <div className="space-y-6">
            {/* Success Message */}
            <div className="bg-green-50 border border-green-200 rounded-lg p-4 flex gap-3">
              <CheckCircle size={24} className="text-green-600 flex-shrink-0" />
              <div>
                <p className="font-semibold text-green-900">Extraction Complete</p>
                <p className="text-sm text-green-700">
                  {result.rows_extracted} measurements extracted successfully
                </p>
              </div>
            </div>

            {/* Result Details */}
            <div className="bg-white rounded-lg border border-slate-200 p-6 space-y-4">
              <div>
                <p className="text-sm font-semibold text-slate-600 uppercase tracking-wide mb-1">
                  Document Title
                </p>
                <p className="text-lg font-semibold text-slate-900">
                  {result.title || 'Garment Specification'}
                </p>
              </div>

              <div>
                <p className="text-sm font-semibold text-slate-600 uppercase tracking-wide mb-2">
                  Size Groups
                </p>
                <div className="flex flex-wrap gap-2">
                  {result.size_groups.map((group, i) => (
                    <span
                      key={i}
                      className="px-3 py-1 bg-blue-100 text-blue-700 text-sm font-medium rounded-full"
                    >
                      {group}
                    </span>
                  ))}
                </div>
              </div>

              <div>
                <p className="text-sm font-semibold text-slate-600 uppercase tracking-wide mb-2">
                  Measurements
                </p>
                <div className="bg-slate-50 rounded p-4 max-h-96 overflow-y-auto">
                  <table className="w-full text-sm">
                    <thead>
                      <tr className="border-b border-slate-200">
                        <th className="text-left py-2 px-2 font-semibold text-slate-700">
                          Description
                        </th>
                        <th className="text-right py-2 px-2 font-semibold text-slate-700">
                          Tolerance
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {result.data.rows.slice(0, 5).map((row, i) => (
                        <tr key={i} className="border-b border-slate-100">
                          <td className="py-2 px-2 text-slate-700 text-xs">
                            {typeof row === 'string' ? row : row[0] || ''}
                          </td>
                          <td className="py-2 px-2 text-slate-600 text-right text-xs">
                            {typeof row === 'string' ? '' : row[1] || ''}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  {result.data.rows.length > 5 && (
                    <p className="text-slate-600 text-xs mt-3 italic">
                      +{result.data.rows.length - 5} more measurements in Excel
                    </p>
                  )}
                </div>
              </div>
            </div>

            {/* Download & Reset Buttons */}
            <div className="flex gap-3">
              <button
                onClick={handleDownload}
                className="flex-1 py-3 px-4 rounded-lg font-semibold bg-green-600 text-white hover:bg-green-700 transition-all active:scale-95 flex items-center justify-center gap-2"
              >
                <Download size={20} />
                Download Excel
              </button>

              <button
                onClick={handleReset}
                className="flex-1 py-3 px-4 rounded-lg font-semibold bg-slate-200 text-slate-900 hover:bg-slate-300 transition-all"
              >
                Extract Another
              </button>
            </div>

            {/* File Info */}
            <div className="text-xs text-slate-500 text-center">
              File: <span className="font-mono">{result.excel_filename}</span>
            </div>
          </div>
        )}
      </div>

      {/* Footer */}
      <div className="bg-slate-900 text-slate-400 text-center text-xs py-4 mt-12">
        <p>Garment Spec OCR v3 | Two-pass extraction with Mistral Pixtral</p>
      </div>
    </div>
  );
}
