'use client';

import { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';

type Region = 'Eastern' | 'Western';
type DivisionOptions = {
  [key in Region]: string[];
};

type CountyMappings = {
  [key: string]: string[];
};

const backendUrl = process.env.NEXT_PUBLIC_BACKEND_URL;

export default function DomainSelectionPage() {
  const [region, setRegion] = useState<Region | ''>('');
  const [division, setDivision] = useState('');
  const [county, setCounty] = useState('');
  const [date, setDate] = useState('');
  const [urls, setUrls] = useState<string[]>([]);
  const [selectedUrl, setSelectedUrl] = useState('');
  const [isLoading, setIsLoading] = useState(false);
  const [isRetest, setIsRetest] = useState(false);

  useEffect(() => {
    const loadUrls = async () => {
      try {
        const response = await fetch(`${backendUrl}/api/urls`);
        
        if (!response.ok) {
          throw new Error('Failed to fetch URLs');
        }
  
        const urls = await response.json();
        setUrls(urls);
      } catch (error) {
        console.error('Failed to load URLs:', error);
      }
    };
  
    loadUrls();
  }, []);

  const divisionOptions: DivisionOptions = {
    'Eastern': [
      'Eastern Division',
      'Northern Division',
      'Southeastern Division'
    ],
    'Western': ['Western Division']
  };

  const countyMappings: CountyMappings = {
    "Eastern Division": ["Crawford", "Dent", "Franklin", "Gasconade", "Jefferson", "Lincoln", "Maries", "Phelps", "St. Charles", "St. Francois", "St. Louis City", "St. Louis County", "Warren", "Washington"],
    "Northern Division": ["Adair", "Audrain", "Chariton", "Clark", "Knox", "Lewis", "Linn", "Macon", "Marion", "Monroe", "Montgomery", "Pike", "Ralls", "Randolph", "Schuyler", "Scotland", "Shelby"],
    "Southeastern Division": ["Bollinger", "Butler", "Cape Girardeau", "Carter", "Dunklin", "Iron", "Madison", "Mississippi", "New Madrid", "Pemiscot", "Perry", "Reynolds", "Ripley", "Scott", "Shannon", "Ste. Genevieve", "Stoddard", "Wayne"],
    'Western Division': ["St. Joseph"]
  };

  const handleSubmit = async () => {
    setIsLoading(true);
    try {
      const response = await fetch(`${backendUrl}/api/pdf/generate`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          url: selectedUrl,
          region,
          division,
          county,
          bobVisitDate: new Date().toISOString(),
          emailSentDate: new Date().toISOString(),
          earliestExpertScanDate: new Date().toISOString(),
          date,
          isRetest
        })
      });

      if (response.ok) {
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'generated_report.pdf';
        a.click();
        window.URL.revokeObjectURL(url);
      } else {
        const errorData = await response.json();
        console.error('PDF generation failed:', errorData);
        alert(`Error: ${errorData.message || 'Failed to generate PDF'}`);
      }
    } catch (error) {
      console.error('PDF generation failed', error);
      alert('An error occurred while generating the PDF');
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-black text-white p-6">
      {isLoading && (
        <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
          <div className="animate-spin rounded-full h-32 w-32 border-t-4 border-blue-500"></div>
        </div>
      )}
      <div className="max-w-md mx-auto space-y-4">
        <div>
          <label className="block text-gray-300 mb-2">Region</label>
          <select
            value={region}
            onChange={(e) => {
              setRegion(e.target.value as Region);
              setDivision('');
              setCounty('');
            }}
            className="w-full bg-gray-800 border border-gray-700 text-white p-2 rounded"
          >
            <option value="">Select Region</option>
            <option value="Eastern">Eastern</option>
            <option value="Western">Western</option>
          </select>
        </div>

        <div>
          <label className="block text-gray-300 mb-2">URL</label>
          <select
            value={selectedUrl}
            onChange={(e) => setSelectedUrl(e.target.value)}
            className="w-full bg-gray-800 border border-gray-700 text-white p-2 rounded"
          >
            <option value="">Select URL</option>
            {urls.map((url, index) => (
              <option key={index} value={url}>
                {url}
              </option>
            ))}
          </select>
        </div>

        <div>
          <label className="block text-gray-300 mb-2">Date</label>
          <div 
            className="relative w-full"
            onClick={() => {
              const dateInput = document.querySelector('input[type="date"]') as HTMLInputElement;
              dateInput?.showPicker();
            }}
          >
            <input
              type="date"
              value={date}
              onChange={(e) => setDate(e.target.value)}
              className="w-full bg-gray-800 border border-gray-700 text-gray-300 p-2 rounded cursor-pointer"
            />
            <div className="absolute inset-0" />
          </div>
        </div>

        {region && (
          <div>
            <label className="block text-gray-300 mb-2">Division</label>
            <select
              value={division}
              onChange={(e) => {
                setDivision(e.target.value);
                setCounty('');
              }}
              className="w-full bg-gray-800 border border-gray-700 text-white p-2 rounded"
            >
              <option value="">Select Division</option>
              {divisionOptions[region].map(div => (
                <option key={div} value={div}>{div}</option>
              ))}
            </select>
          </div>
        )}

        {division && (
          <div>
            <label className="block text-gray-300 mb-2">County</label>
            <select
              value={county}
              onChange={(e) => setCounty(e.target.value)}
              className="w-full bg-gray-800 border border-gray-700 text-white p-2 rounded"
            >
              <option value="">Select County</option>
              {countyMappings[division].map(countyName => (
                <option key={countyName} value={countyName}>{countyName}</option>
              ))}
            </select>
          </div>
        )}

        <div className="flex items-center space-x-2">
          <input
            type="checkbox"
            id="retest"
            checked={isRetest}
            onChange={(e) => setIsRetest(e.target.checked)}
            className="bg-gray-800 border border-gray-700 rounded"
          />
          <label htmlFor="retest" className="text-gray-300">Is this a retest?</label>
        </div>

        <button
          onClick={handleSubmit}
          disabled={!region || !selectedUrl || !date || !division || !county || isLoading}
          className="w-full bg-gray-800 hover:bg-gray-700 text-white p-2 rounded disabled:opacity-50"
        >
          {isLoading ? 'Generating...' : 'Generate Report'}
        </button>
      </div>
    </div>
  );
}