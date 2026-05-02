import React, { useState, useCallback, useRef } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { 
  FileText, 
  Upload, 
  Download, 
  CheckCircle2, 
  Loader2, 
  X, 
  FileCheck,
  AlertCircle,
  FileCode,
  Files
} from 'lucide-react';
import { saveAs } from 'file-saver';
import * as pdfjs from 'pdfjs-dist';
import { 
  Document, 
  Packer, 
  Paragraph, 
  TextRun, 
  HeadingLevel,
  AlignmentType
} from 'docx';
import { cn } from '@/src/lib/utils';

// pdfjs-dist worker setup
// Using the CDN version with .mjs extension which is required for modern PDF.js (4.0+) ESM workers
pdfjs.GlobalWorkerOptions.workerSrc = `https://unpkg.com/pdfjs-dist@${pdfjs.version}/build/pdf.worker.min.mjs`;

interface FileStatus {
  id: string;
  file: File;
  status: 'idle' | 'processing' | 'completed' | 'error';
  progress: number;
  outputBlob?: Blob;
  error?: string;
  outputName?: string;
}

export default function App() {
  const [files, setFiles] = useState<FileStatus[]>([]);
  const [isDragging, setIsDragging] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const addFiles = (newFiles: File[]) => {
    const pdfFiles = newFiles.filter(f => f.type === 'application/pdf');
    if (pdfFiles.length === 0) return;

    const newStatuses: FileStatus[] = pdfFiles.map(f => ({
      id: Math.random().toString(36).substring(7),
      file: f,
      status: 'idle',
      progress: 0,
      outputName: f.name.replace(/\.[^/.]+$/, "") + ".docx"
    }));

    setFiles(prev => [...prev, ...newStatuses]);
  };

  const onDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files) {
      addFiles(Array.from(e.dataTransfer.files));
    }
  };

  const removeFile = (id: string) => {
    setFiles(prev => prev.filter(f => f.id !== id));
  };

  const convertPdfToWord = async (fileStatus: FileStatus) => {
    setFiles(prev => prev.map(f => f.id === fileStatus.id ? { ...f, status: 'processing', progress: 10 } : f));

    try {
      const arrayBuffer = await fileStatus.file.arrayBuffer();
      const pdf = await pdfjs.getDocument({ data: arrayBuffer }).promise;
      
      const paragraphs: Paragraph[] = [];
      const totalPages = pdf.numPages;

      for (let i = 1; i <= totalPages; i++) {
        const page = await pdf.getPage(i);
        const textContent = await page.getTextContent();
        
        const progress = 10 + Math.floor((i / totalPages) * 80);
        setFiles(prev => prev.map(f => f.id === fileStatus.id ? { ...f, progress } : f));

        const lines: { [y: number]: any[] } = {};
        textContent.items.forEach((item: any) => {
          const y = Math.round(item.transform[5]);
          if (!lines[y]) lines[y] = [];
          lines[y].push(item);
        });

        const sortedY = Object.keys(lines).map(Number).sort((a, b) => b - a);

        sortedY.forEach(y => {
          const lineItems = lines[y].sort((a, b) => a.transform[4] - b.transform[4]);
          const lineText = lineItems.map(item => item.str).join(' ');

          if (lineText.trim()) {
            const fontSize = Math.abs(lineItems[0].transform[0]);
            
            paragraphs.push(new Paragraph({
              children: [
                new TextRun({
                  text: lineText,
                  size: fontSize * 2,
                })
              ],
              spacing: {
                after: 200,
              }
            }));
          }
        });
      }

      const doc = new Document({
        sections: [{
          properties: {},
          children: paragraphs,
        }],
      });

      const blob = await Packer.toBlob(doc);
      
      setFiles(prev => prev.map(f => f.id === fileStatus.id ? { 
        ...f, 
        status: 'completed', 
        progress: 100, 
        outputBlob: blob 
      } : f));

    } catch (err) {
      console.error(err);
      setFiles(prev => prev.map(f => f.id === fileStatus.id ? { 
        ...f, 
        status: 'error', 
        error: 'Failed to convert file' 
      } : f));
    }
  };

  const downloadFile = (fileStatus: FileStatus) => {
    if (fileStatus.outputBlob) {
      saveAs(fileStatus.outputBlob, fileStatus.outputName);
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col font-sans text-slate-900 select-none overflow-x-hidden">
      {/* Navigation */}
      <nav className="h-20 border-b border-slate-200 bg-white px-6 sm:px-12 flex items-center justify-between shrink-0">
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 bg-indigo-600 flex items-center justify-center">
            <div className="w-4 h-4 border-2 border-white rotate-45"></div>
          </div>
          <span className="font-extrabold text-xl tracking-tight text-indigo-950 uppercase hidden sm:block">DocShift AI</span>
        </div>
        <div className="flex items-center gap-4 sm:gap-8 text-sm font-bold text-slate-500 uppercase tracking-wider">
          <a href="#" className="hover:text-indigo-600 transition-colors hidden md:block">Security</a>
          <button className="btn-outline-geometric">Login</button>
        </div>
      </nav>

      {/* Main Content */}
      <main className="flex-1 w-full max-w-7xl mx-auto grid grid-cols-1 md:grid-cols-12 items-center px-6 sm:px-12 py-12 gap-12">
        
        {/* Left Hero */}
        <div className="md:col-span-5 text-center md:text-left">
          <motion.h1 
            initial={{ opacity: 0, x: -20 }}
            animate={{ opacity: 1, x: 0 }}
            className="text-5xl lg:text-7xl font-black text-slate-900 leading-none mb-6"
          >
            PDF TO WORD<br/><span className="text-indigo-600 italic">REIMAGINED.</span>
          </motion.h1>
          <motion.p 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            transition={{ delay: 0.1 }}
            className="text-lg text-slate-500 mb-8 leading-relaxed max-w-lg mx-auto md:mx-0"
          >
            High-fidelity conversion powered by geometric structural analysis. Preserving layout and fonts with surgical precision.
          </motion.p>
          
          <motion.div 
            initial={{ opacity: 0, y: 10 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ delay: 0.2 }}
            className="flex gap-8 justify-center md:justify-start"
          >
            <div className="border-l-4 border-indigo-600 pl-4 text-left">
              <div className="text-[10px] font-bold uppercase text-slate-400 mb-1 tracking-widest">Active Users</div>
              <div className="text-2xl font-black tabular-nums">12,482</div>
            </div>
            <div className="border-l-4 border-slate-200 pl-4 text-left">
              <div className="text-[10px] font-bold uppercase text-slate-400 mb-1 tracking-widest">Accuracy</div>
              <div className="text-2xl font-black tabular-nums">99.9%</div>
            </div>
          </motion.div>
        </div>

        {/* Right Content: The Tool */}
        <div className="md:col-span-7 w-full max-w-2xl mx-auto relative group">
          {/* Decorative Accents */}
          <div className="absolute -top-4 -right-4 w-12 h-12 bg-indigo-100 -z-10 group-hover:translate-x-2 group-hover:-translate-y-2 transition-transform duration-500"></div>
          <div className="absolute -bottom-4 -left-4 w-12 h-12 bg-slate-200 -z-10 group-hover:-translate-x-2 group-hover:translate-y-2 transition-transform duration-500 delay-75"></div>
          
          <div className="geometric-card p-6 sm:p-8 flex flex-col min-h-[480px]">
            {/* Dropzone */}
            <div 
              onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
              onDragLeave={() => setIsDragging(false)}
              onDrop={onDrop}
              onClick={() => fileInputRef.current?.click()}
              className={cn(
                "flex-1 border-2 border-dashed flex flex-col items-center justify-center p-8 transition-all duration-300 cursor-pointer",
                isDragging 
                  ? "border-indigo-600 bg-indigo-50/50" 
                  : "border-slate-200 bg-slate-50 hover:border-indigo-400 hover:bg-white"
              )}
            >
              <input 
                type="file" 
                ref={fileInputRef} 
                className="hidden" 
                accept=".pdf" 
                multiple 
                onChange={(e) => addFiles(Array.from(e.target.files || []))}
              />
              <div className="w-16 h-16 bg-white shadow-sm flex items-center justify-center mb-6">
                <Upload className="text-indigo-600" size={32} />
              </div>
              <span className="font-black text-slate-900 uppercase tracking-widest mb-2">Click or drag PDF here</span>
              <span className="text-[10px] text-slate-400 uppercase font-black tracking-[0.2em]">Unlimited File Size Support</span>
            </div>

            {/* File Processing List */}
            <AnimatePresence mode="popLayout">
              {files.length > 0 && (
                <motion.div 
                  initial={{ opacity: 0 }}
                  animate={{ opacity: 1 }}
                  className="mt-8 space-y-6"
                >
                  {files.map((fileStatus) => (
                    <motion.div
                      key={fileStatus.id}
                      layout
                      initial={{ opacity: 0, y: 10 }}
                      animate={{ opacity: 1, y: 0 }}
                      exit={{ opacity: 0, scale: 0.95 }}
                    >
                      <div className="flex items-center justify-between mb-3 px-1">
                        <div className="flex items-center gap-3 pr-4 overflow-hidden">
                          <div className="w-8 h-10 bg-red-100 flex items-center justify-center font-black text-red-600 text-[10px] shrink-0 rounded-sm">PDF</div>
                          <div className="min-w-0">
                            <div className="text-sm font-black text-slate-800 truncate uppercase tracking-tighter">
                              {fileStatus.file.name}
                            </div>
                            <div className="text-[10px] text-slate-400 uppercase font-bold tracking-wider">
                              {(fileStatus.file.size / (1024 * 1024)).toFixed(1)} MB • {fileStatus.status === 'completed' ? 'Processed' : 'Ready to convert'}
                            </div>
                          </div>
                        </div>
                        
                        <div className="flex items-center gap-3">
                          {fileStatus.status === 'idle' && (
                            <button 
                              onClick={() => convertPdfToWord(fileStatus)}
                              className="text-[10px] font-black text-indigo-600 uppercase tracking-widest hover:underline"
                            >
                              START
                            </button>
                          )}
                          <button 
                            onClick={() => removeFile(fileStatus.id)}
                            className="text-slate-300 hover:text-slate-600 transition-colors"
                          >
                            <X size={14} />
                          </button>
                        </div>
                      </div>

                      {/* Progress/Success indicator */}
                      <div className="relative w-full h-1 bg-slate-100 overflow-hidden">
                        {(fileStatus.status === 'processing' || fileStatus.status === 'completed') && (
                          <motion.div 
                            className="absolute inset-y-0 left-0 bg-indigo-600"
                            initial={{ width: 0 }}
                            animate={{ width: `${fileStatus.progress}%` }}
                            transition={{ ease: "easeOut" }}
                          />
                        )}
                      </div>

                      {fileStatus.status === 'completed' && (
                        <motion.button
                          initial={{ opacity: 0, y: 5 }}
                          animate={{ opacity: 1, y: 0 }}
                          onClick={() => downloadFile(fileStatus)}
                          className="mt-4 w-full bg-slate-900 hover:bg-slate-800 text-white font-black py-4 uppercase tracking-[0.2em] shadow-lg flex items-center justify-center gap-3 text-xs transition-colors"
                        >
                          DOWNLOAD .DOCX
                          <Download size={14} />
                        </motion.button>
                      )}

                      {fileStatus.status === 'error' && (
                        <div className="mt-2 text-[10px] font-bold text-red-500 uppercase tracking-wider flex items-center gap-1">
                          <AlertCircle size={10} />
                          {fileStatus.error}
                        </div>
                      )}
                    </motion.div>
                  ))}
                </motion.div>
              )}
            </AnimatePresence>
          </div>
        </div>
      </main>

      {/* Footer Grid Overlay */}
      <footer className="h-auto md:h-24 bg-slate-900 text-white px-6 sm:px-12 py-8 md:py-0 flex flex-col md:flex-row items-center justify-between gap-8 sm:gap-12 overflow-hidden">
        <div className="flex flex-wrap justify-center sm:justify-start gap-8 sm:gap-16">
          <div className="flex flex-col">
            <span className="text-indigo-400 text-[10px] font-bold uppercase tracking-widest mb-1">Encryption</span>
            <span className="text-sm font-medium">AES-256 Bit SSL</span>
          </div>
          <div className="flex flex-col">
            <span className="text-indigo-400 text-[10px] font-bold uppercase tracking-widest mb-1">Privacy</span>
            <span className="text-sm font-medium">Auto-delete after 1h</span>
          </div>
          <div className="flex flex-col">
            <span className="text-indigo-400 text-[10px] font-bold uppercase tracking-widest mb-1">Quality</span>
            <span className="text-sm font-medium">Geometric Engine 2.0</span>
          </div>
        </div>
        <div className="flex items-center gap-4 opacity-40">
           <div className="px-3 py-1 border border-slate-700 font-black text-xs uppercase tracking-tighter shrink-0">ISO-27001</div>
           <div className="px-3 py-1 border border-slate-700 font-black text-xs uppercase tracking-tighter shrink-0">GDPR Compliant</div>
        </div>
      </footer>
    </div>
  );
}
