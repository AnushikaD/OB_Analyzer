/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo } from 'react';
import { Download, Upload, Loader2, FileText, X } from 'lucide-react';
import * as XLSX from 'xlsx';
import { motion, AnimatePresence } from 'motion/react';
import { GoogleGenAI, Type } from "@google/genai";

interface LevyItem {
  lot: string;
  unit: string;
  owner: string;
  dueDate: string;
  description: string;
  fund: 'Admin' | 'Sinking';
  amount: number;
  type: 'Arrear' | 'Advance';
}

const formatCurrency = (amount: number) => {
  return new Intl.NumberFormat('en-AU', { style: 'currency', currency: 'AUD' }).format(amount);
};

const formatDate = (dateStr: string) => {
  if (!dateStr.includes('-')) return dateStr;
  const [year, month, day] = dateStr.split('-');
  return `${day}/${month}/${year}`;
};

export default function App() {
  const [levyData, setLevyData] = useState<LevyItem[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [files, setFiles] = useState<File[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [viewType, setViewType] = useState<'Arrear' | 'Advance'>('Arrear');

  const filteredData = useMemo(() => levyData.filter(i => i.type === viewType), [levyData, viewType]);

  const stats = useMemo(() => {
    const calc = (type: 'Arrear' | 'Advance') => {
      const data = levyData.filter(i => i.type === type);
      return {
        admin: data.filter(i => i.fund === 'Admin').reduce((s, i) => s + i.amount, 0),
        sinking: data.filter(i => i.fund === 'Sinking').reduce((s, i) => s + i.amount, 0),
        total: data.reduce((s, i) => s + i.amount, 0)
      };
    };
    return {
      arrears: calc('Arrear'),
      advances: calc('Advance')
    };
  }, [levyData]);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      setFiles(prev => [...prev, ...Array.from(e.target.files!)].slice(0, 5));
    }
  };

  const removeFile = (index: number) => {
    setFiles(prev => prev.filter((_, i) => i !== index));
  };

  const processFiles = async () => {
    if (files.length === 0) return;
    setIsProcessing(true);
    setError(null);

    try {
      const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });
      
      const fileParts = await Promise.all(
        files.map(async (file) => {
          const reader = new FileReader();
          const base64Promise = new Promise<string>((resolve) => {
            reader.onload = () => {
              const base64 = (reader.result as string).split(',')[1];
              resolve(base64);
            };
          });
          reader.readAsDataURL(file);
          const data = await base64Promise;
          return {
            inlineData: {
              data,
              mimeType: file.type || 'application/pdf'
            }
          };
        })
      );

      const response = await ai.models.generateContent({
        model: "gemini-3-flash-preview",
        contents: [
          ...fileParts,
          { text: "Analyze these strata reports. You must extract EXACT and COMPLETE descriptions for every line item exactly as written in the reports. \n\n1. RECONCILIATION: First, look at the Balance Sheet to find the control totals for 'Levies Receivable' (Arrears) and 'Levies Paid in Advance' (Credits/Advances) for both Administrative and Sinking funds.\n2. DETAILS: Then, extract individual line items from the Lot Position report or Transaction Summary that sum up to these control totals. \n- Arrears are usually debits.\n- Advances are usually credits marked with 'CR' or appearing in the 'Paid in Advance' section.\n3. LOTS: Ensure Lot numbers are captured accurately (especially T3/994).\n4. OUTPUT: Provide a flat JSON array of objects with the exact schema. For each item, identify if it is an 'Arrear' or 'Advance'. Return only the JSON array." }
        ],
        config: {
          responseMimeType: "application/json",
          responseSchema: {
            type: Type.ARRAY,
            items: {
              type: Type.OBJECT,
              properties: {
                lot: { type: Type.STRING, description: "Lot number or identifier" },
                unit: { type: Type.STRING, description: "Unit number" },
                owner: { type: Type.STRING, description: "Owner name" },
                dueDate: { type: Type.STRING, description: "YYYY-MM-DD" },
                description: { type: Type.STRING, description: "EXACT description as it appears in the report" },
                fund: { type: Type.STRING, enum: ["Admin", "Sinking"], description: "The fund the levy belongs to" },
                amount: { type: Type.NUMBER, description: "Absolute positive amount" },
                type: { type: Type.STRING, enum: ["Arrear", "Advance"], description: "Whether this is an Arrear (owing) or Advance (credit)" }
              },
              required: ["lot", "unit", "owner", "dueDate", "description", "fund", "amount", "type"]
            }
          }
        }
      });

      const result = JSON.parse(response.text || '[]') as LevyItem[];
      setLevyData(result);
    } catch (err: any) {
      setError(err.message || "An error occurred during processing.");
    } finally {
      setIsProcessing(false);
    }
  };

  const exportToExcel = () => {
    const dataForExport = [...filteredData].sort((a, b) => {
      const numA = parseInt(a.lot);
      const numB = parseInt(b.lot);
      const lotDiff = (!isNaN(numA) && !isNaN(numB)) ? numA - numB : a.lot.localeCompare(b.lot);
      if (lotDiff !== 0) return lotDiff;
      if (a.fund !== b.fund) return a.fund === 'Admin' ? -1 : 1;
      return a.dueDate.localeCompare(b.dueDate);
    }).map(item => ({
      'Lot #': item.lot,
      'Unit #': item.unit,
      'Owner': item.owner,
      'Due Date': formatDate(item.dueDate),
      'Description': item.description,
      'Fund': item.fund === 'Admin' ? 'Administrative Fund' : 'Sinking Fund',
      'Amount': item.amount,
      'Type': item.type
    }));

    const worksheet = XLSX.utils.json_to_sheet(dataForExport);
    const workbook = XLSX.utils.book_new();
    const sheetName = viewType === 'Arrear' ? "Arrears" : "Advances";
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
    XLSX.writeFile(workbook, `Levies_${sheetName}_Breakdown.xlsx`);
  };

  const lots = useMemo(() => 
    Array.from(new Set(filteredData.map(i => i.lot))).sort((a: string, b: string) => {
      const numA = parseInt(a);
      const numB = parseInt(b);
      if (!isNaN(numA) && !isNaN(numB)) return numA - numB;
      return a.localeCompare(b);
    }), [filteredData]);

  return (
    <div className="min-h-screen bg-gray-50 p-4 md:p-8 font-sans">
      <div className="max-w-7xl mx-auto">
        <header className="mb-8 flex flex-wrap justify-between items-start gap-4">
          <div>
            <h1 className="text-3xl font-bold text-gray-900 mb-2">Levies Breakdown</h1>
            <p className="text-gray-600">Extract data dynamically from strata reports</p>
          </div>
          
          <div className="flex flex-wrap items-center gap-4">
            {levyData.length > 0 && (
              <div className="flex bg-gray-200/50 p-1 rounded-xl">
                <button
                  onClick={() => setViewType('Arrear')}
                  className={`px-4 py-2 rounded-lg text-sm font-bold transition-all ${
                    viewType === 'Arrear' 
                      ? "bg-white text-gray-900 shadow-sm" 
                      : "text-gray-500 hover:text-gray-700"
                  }`}
                >
                  Arrears
                </button>
                <button
                  onClick={() => setViewType('Advance')}
                  className={`px-4 py-2 rounded-lg text-sm font-bold transition-all ${
                    viewType === 'Advance' 
                      ? "bg-white text-gray-900 shadow-sm" 
                      : "text-gray-500 hover:text-gray-700"
                  }`}
                >
                  Advances
                </button>
              </div>
            )}

            {filteredData.length > 0 && (
              <button 
                onClick={exportToExcel}
                className="flex items-center gap-2 bg-neutral-900 text-white font-semibold py-2.5 px-4 rounded-lg shadow-sm hover:bg-neutral-800 transition-all active:scale-95 cursor-pointer"
              >
                <Download size={18} />
                Export {viewType === 'Arrear' ? 'Arrears' : 'Advances'}
              </button>
            )}
          </div>
        </header>

        <AnimatePresence mode="wait">
          {levyData.length === 0 && !isProcessing ? (
            <motion.div 
              key="upload"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
              className="bg-white p-12 rounded-3xl shadow-sm border border-gray-100 text-center"
            >
              <div className="w-16 h-16 bg-gray-50 rounded-2xl flex items-center justify-center mx-auto mb-6">
                <Upload size={32} className="text-gray-400" />
              </div>
              <h2 className="text-xl font-bold text-gray-900 mb-2">Upload Strata Reports</h2>
              <p className="text-gray-500 mb-8">Upload Balance Sheet, Lot Positions, or Transaction Summaries (PDF and Images supported)</p>
              
              <div className="max-w-md mx-auto space-y-4">
                <label className="block w-full py-4 px-6 border-2 border-dashed border-gray-200 rounded-2xl cursor-pointer hover:border-neutral-900 transition-colors">
                  <input type="file" multiple className="hidden" onChange={handleFileChange} />
                  <span className="text-sm font-semibold text-gray-600">Select files to process</span>
                </label>

                {files.length > 0 && (
                  <div className="text-left space-y-2">
                    {files.map((f, i) => (
                      <div key={i} className="flex items-center justify-between bg-gray-50 p-3 rounded-xl border border-gray-100">
                        <div className="flex items-center gap-2 overflow-hidden">
                          <FileText size={16} className="text-gray-400 flex-shrink-0" />
                          <span className="text-xs font-medium text-gray-600 truncate">{f.name}</span>
                        </div>
                        <button onClick={() => removeFile(i)} className="text-gray-400 hover:text-red-500"><X size={16} /></button>
                      </div>
                    ))}
                    <button 
                      onClick={processFiles}
                      className="w-full mt-4 bg-neutral-900 text-white font-bold py-4 rounded-xl shadow-lg hover:bg-neutral-800 transition-all active:scale-[0.98]"
                    >
                      Process Reports
                    </button>
                  </div>
                )}
                {error && <p className="text-red-500 text-sm mt-4">{error}</p>}
              </div>
            </motion.div>
          ) : isProcessing ? (
            <motion.div 
              key="loading"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              className="py-24 text-center"
            >
              <Loader2 size={48} className="text-neutral-900 animate-spin mx-auto mb-4" />
              <h2 className="text-xl font-bold text-gray-900">Analyzing Financial Data...</h2>
              <p className="text-gray-500">Extracting lot positions and fund details</p>
            </motion.div>
          ) : (
            <motion.div 
              key="results"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              className="space-y-10"
            >
              <div className="grid grid-cols-1 lg:grid-cols-4 gap-6">
                <div className="lg:col-span-3 grid grid-cols-1 md:grid-cols-3 gap-6">
                  <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100">
                    <p className="text-sm font-medium text-gray-500 uppercase tracking-wider mb-1">Admin {viewType === 'Arrear' ? 'Arrears' : 'Advances'}</p>
                    <p className={`text-2xl font-bold ${viewType === 'Arrear' ? 'text-blue-600' : 'text-indigo-600'}`}>
                      {formatCurrency(viewType === 'Arrear' ? stats.arrears.admin : stats.advances.admin)}
                    </p>
                  </div>
                  <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100">
                    <p className="text-sm font-medium text-gray-500 uppercase tracking-wider mb-1">Sinking {viewType === 'Arrear' ? 'Arrears' : 'Advances'}</p>
                    <p className={`text-2xl font-bold ${viewType === 'Arrear' ? 'text-emerald-600' : 'text-teal-600'}`}>
                      {formatCurrency(viewType === 'Arrear' ? stats.arrears.sinking : stats.advances.sinking)}
                    </p>
                  </div>
                  <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100">
                    <p className="text-sm font-medium text-gray-500 uppercase tracking-wider mb-1">Total {viewType === 'Arrear' ? 'Outstanding' : 'Credits'}</p>
                    <p className="text-2xl font-bold text-gray-900">
                      {formatCurrency(viewType === 'Arrear' ? stats.arrears.total : stats.advances.total)}
                    </p>
                  </div>
                </div>

                <div className="bg-neutral-900 text-white p-6 rounded-xl shadow-xl flex flex-col justify-center">
                  <p className="text-xs font-bold text-neutral-400 uppercase tracking-widest mb-3">Overall Context</p>
                  <div className="space-y-3">
                    <div className="flex justify-between items-center text-sm">
                      <span className="text-neutral-400">Total Arrears:</span>
                      <span className="font-bold text-blue-400">{formatCurrency(stats.arrears.total)}</span>
                    </div>
                    <div className="flex justify-between items-center text-sm">
                      <span className="text-neutral-400">Total Advances:</span>
                      <span className="font-bold text-indigo-400">{formatCurrency(stats.advances.total)}</span>
                    </div>
                    <div className="pt-2 border-t border-neutral-800 flex justify-between items-center text-xs">
                      <span className="text-neutral-500">Net Position:</span>
                      <span className={`font-black ${stats.arrears.total - stats.advances.total > 0 ? 'text-red-400' : 'text-green-400'}`}>
                        {formatCurrency(stats.arrears.total - stats.advances.total)}
                      </span>
                    </div>
                  </div>
                </div>
              </div>

              <div className="space-y-12">
                {lots.map(lotId => {
                  const items = filteredData.filter(i => i.lot === lotId).sort((a, b) => {
                    if (a.fund !== b.fund) return a.fund === 'Admin' ? -1 : 1;
                    return a.dueDate.localeCompare(b.dueDate);
                  });
                  const adminItems = items.filter(i => i.fund === 'Admin');
                  const sinkingItems = items.filter(i => i.fund === 'Sinking');
                  const lotAdmin = adminItems.reduce((sum, i) => sum + i.amount, 0);
                  const lotSinking = sinkingItems.reduce((sum, i) => sum + i.amount, 0);

                  return (
                    <section key={lotId} className="bg-white rounded-2xl shadow-md overflow-hidden border border-gray-200">
                      <div className="bg-gray-50 px-6 py-4 border-b border-gray-200 flex flex-wrap justify-between items-center gap-4">
                        <div>
                          <h2 className="text-xl font-bold text-gray-900">Lot {lotId} <span className="font-normal text-gray-500 ml-2">(Unit {items[0].unit})</span></h2>
                          <p className="text-sm text-gray-600">{items[0].owner}</p>
                        </div>
                        <div className="flex gap-4">
                          {lotAdmin > 0 && (
                            <div className="text-right">
                              <p className="text-xs font-bold text-blue-400 uppercase tracking-tight">Admin Total</p>
                              <p className="font-semibold text-blue-700">{formatCurrency(lotAdmin)}</p>
                            </div>
                          )}
                          {lotSinking > 0 && (
                            <div className="text-right border-l border-gray-300 pl-4">
                              <p className="text-xs font-bold text-emerald-400 uppercase tracking-tight">Sinking Total</p>
                              <p className="font-semibold text-emerald-700">{formatCurrency(lotSinking)}</p>
                            </div>
                          )}
                        </div>
                      </div>
                      
                      <div className="p-6 space-y-8">
                        {adminItems.length > 0 && (
                          <div className="space-y-3">
                            <h3 className="text-sm font-bold text-blue-600 uppercase tracking-wide flex items-center gap-2">
                              <div className="w-2 h-2 rounded-full bg-blue-600" />
                              Administrative Fund
                            </h3>
                            <div className="overflow-x-auto border border-gray-100 rounded-xl">
                              <table className="w-full text-left border-collapse">
                                <thead>
                                  <tr className="bg-blue-50/30">
                                    <th className="px-6 py-3 text-xs font-bold text-gray-400 uppercase tracking-wider w-32">Due Date</th>
                                    <th className="px-6 py-3 text-xs font-bold text-gray-400 uppercase tracking-wider">Description</th>
                                    <th className="px-6 py-3 text-xs font-bold text-gray-400 uppercase tracking-wider w-32 text-right">Amount</th>
                                  </tr>
                                </thead>
                                <tbody className="divide-y divide-gray-100">
                                  {adminItems.map((item, idx) => (
                                    <tr key={idx} className="hover:bg-gray-50/80 transition-colors">
                                      <td className="px-6 py-4 text-sm font-medium text-gray-700 whitespace-nowrap">{formatDate(item.dueDate)}</td>
                                      <td className="px-6 py-4 text-sm text-gray-600">{item.description}</td>
                                      <td className="px-6 py-4 text-sm font-bold text-gray-900 text-right">{formatCurrency(item.amount)}</td>
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
                          </div>
                        )}

                        {sinkingItems.length > 0 && (
                          <div className="space-y-3">
                            <h3 className="text-sm font-bold text-emerald-600 uppercase tracking-wide flex items-center gap-2">
                              <div className="w-2 h-2 rounded-full bg-emerald-600" />
                              Sinking Fund
                            </h3>
                            <div className="overflow-x-auto border border-gray-100 rounded-xl">
                              <table className="w-full text-left border-collapse">
                                <thead>
                                  <tr className="bg-emerald-50/30">
                                    <th className="px-6 py-3 text-xs font-bold text-gray-400 uppercase tracking-wider w-32">Due Date</th>
                                    <th className="px-6 py-3 text-xs font-bold text-gray-400 uppercase tracking-wider">Description</th>
                                    <th className="px-6 py-3 text-xs font-bold text-gray-400 uppercase tracking-wider w-32 text-right">Amount</th>
                                  </tr>
                                </thead>
                                <tbody className="divide-y divide-gray-100">
                                  {sinkingItems.map((item, idx) => (
                                    <tr key={idx} className="hover:bg-gray-50/80 transition-colors">
                                      <td className="px-6 py-4 text-sm font-medium text-gray-700 whitespace-nowrap">{formatDate(item.dueDate)}</td>
                                      <td className="px-6 py-4 text-sm text-gray-600">{item.description}</td>
                                      <td className="px-6 py-4 text-sm font-bold text-gray-900 text-right">{formatCurrency(item.amount)}</td>
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
                          </div>
                        )}
                      </div>
                    </section>
                  );
                })}
                <div className="flex justify-center pb-12">
                   <button 
                    onClick={() => { setLevyData([]); setFiles([]); }}
                    className="text-sm font-bold text-gray-400 hover:text-gray-600 underline"
                  >
                    Clear and start new analysis
                  </button>
                </div>
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </div>
    </div>
  );
}
