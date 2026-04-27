import { useState, useCallback, useRef, ChangeEvent, useEffect } from 'react';
import { Users, UserPlus, Grid2X2, Shuffle, Copy, Trash2, ListChecks, Check, FileSpreadsheet, Info, Plus, X, ArrowRightLeft, Merge, GripVertical, Sparkles } from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import * as XLSX from 'xlsx';
import confetti from 'canvas-confetti';

type GroupMode = 'numGroups' | 'sizeGroups';

interface Member {
  id: string;
  name: string;
  empId: string;
  unit: string;
}

export default function App() {
  const [members, setMembers] = useState<Member[]>([]);
  const [newMember, setNewMember] = useState({ name: '', empId: '', unit: '' });
  const [mode, setMode] = useState<GroupMode>('numGroups');
  const [value, setValue] = useState(2);
  const [groups, setGroups] = useState<Member[][]>([]);
  const [copied, setCopied] = useState(false);
  const [showExample, setShowExample] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const [isGenerating, setIsGenerating] = useState(false);
  const [shuffleDisplay, setShuffleDisplay] = useState('');
  const [wheelRotation, setWheelRotation] = useState(0);
  const [draggedMember, setDraggedMember] = useState<{ memberId: string, fromGroup: number } | null>(null);

  const addMember = () => {
    if (!newMember.name.trim()) return;
    const member: Member = {
      id: Math.random().toString(36).substr(2, 9),
      ...newMember
    };
    setMembers([...members, member]);
    setNewMember({ name: '', empId: '', unit: '' });
  };

  const removeMember = (id: string) => {
    setMembers(members.filter(m => m.id !== id));
  };

  const handleFileUpload = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = evt.target?.result;
        // Use XLSX.read with type: 'array' for better cross-encoding support
        const workbook = XLSX.read(data, { type: 'array' });
        const wsname = workbook.SheetNames[0];
        const ws = workbook.Sheets[wsname];
        const sheetData = XLSX.utils.sheet_to_json(ws, { header: 1 }) as any[][];

        const newMembers: Member[] = sheetData.slice(1)
          .filter(row => row[0]) 
          .map(row => ({
            id: Math.random().toString(36).substr(2, 9),
            name: String(row[0] || '').trim(),
            empId: String(row[1] || '').trim(),
            unit: String(row[2] || '').trim(),
          }));
        
        setMembers(prev => [...prev, ...newMembers]);
      } catch (err) {
        console.error("File import error:", err);
      } finally {
        if (fileInputRef.current) fileInputRef.current.value = '';
      }
    };
    // readAsArrayBuffer is more robust for different file encodings in sheetjs
    reader.readAsArrayBuffer(file);
  };

  const importExample = () => {
    const surnames = ['王', '李', '張', '劉', '陳', '楊', '黃', '趙', '吳', '周'];
    const names = ['小明', '志豪', '雅婷', '怡君', '淑芬', '俊宏', '威廷', '冠宇', '家豪', '欣怡'];
    const units = ['研發部', '行銷部', '人力資源部', '財務部', '資訊部', '業務部', '生產部'];
    
    const exampleMembers: Member[] = Array.from({ length: 15 }).map(() => {
      const surname = surnames[Math.floor(Math.random() * surnames.length)];
      const name = names[Math.floor(Math.random() * names.length)];
      const unit = units[Math.floor(Math.random() * units.length)];
      const randomId = Math.floor(Math.random() * 9000) + 1000;
      
      return {
        id: Math.random().toString(36).substr(2, 9),
        name: `${surname}${name}`,
        empId: `EMP${randomId}`,
        unit: unit
      };
    });

    setMembers(prev => [...prev, ...exampleMembers]);
  };

  const handleGroup = useCallback(() => {
    if (members.length === 0) return;
    
    setIsGenerating(true);
    setGroups([]);
    setWheelRotation(0);

    // Wheel Animation logic
    let startTimestamp: number | null = null;
    const duration = 4000; // 4 seconds
    
    const animate = (timestamp: number) => {
      if (!startTimestamp) startTimestamp = timestamp;
      const progress = timestamp - startTimestamp;
      
      const rotation = (progress / duration) * 3600; // 10 circles
      setWheelRotation(rotation);
      
      // Select random name every few frames
      if (progress % 100 < 20) {
        setShuffleDisplay(members[Math.floor(Math.random() * members.length)].name);
      }

      if (progress < duration) {
        requestAnimationFrame(animate);
      } else {
        finalizeGroups();
      }
    };

    requestAnimationFrame(animate);

    const finalizeGroups = () => {
      // Fireworks!
      confetti({
        particleCount: 150,
        spread: 70,
        origin: { y: 0.6 },
        colors: ['#4f46e5', '#10b981', '#f59e0b', '#ef4444']
      });

      const shuffled = [...members].sort(() => Math.random() - 0.5);
      const result: Member[][] = [];
      
      if (mode === 'numGroups') {
        const numGroups = Math.max(1, Math.min(value, members.length));
        for (let i = 0; i < numGroups; i++) {
          result.push([]);
        }
        shuffled.forEach((m, index) => {
          result[index % numGroups].push(m);
        });
      } else {
        const size = Math.max(1, value);
        for (let i = 0; i < shuffled.length; i += size) {
          result.push(shuffled.slice(i, i + size));
        }
      }

      setGroups(result);
      setIsGenerating(false);
    };
  }, [members, mode, value]);

  const clearAll = () => {
    setMembers([]);
    setGroups([]);
  };

  const copyResults = async () => {
    if (groups.length === 0) return;
    const text = groups
      .map((g, i) => `第 ${i + 1} 組:\n${g.map(m => `- ${m.name} (${m.empId}/${m.unit})`).join('\n')}`)
      .join('\n\n');
    await navigator.clipboard.writeText(text);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const exportExcel = () => {
    if (groups.length === 0) return;
    
    const exportData: any[] = [];
    groups.forEach((group, index) => {
      group.forEach(member => {
        exportData.push({
          '組別': `第 ${index + 1} 組`,
          '姓名': member.name,
          '員工代號': member.empId,
          '單位': member.unit
        });
      });
    });

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "分組名單");
    
    // For .xlsx, encoding is handled by the format. 
    // If user needs CSV, we would ensure UTF-8 BOM, but .xlsx is preferred for Excel.
    XLSX.writeFile(workbook, `分組結果_${new Date().getFullYear()}${(new Date().getMonth()+1).toString().padStart(2,'0')}${new Date().getDate().toString().padStart(2,'0')}.xlsx`);
  };

  const moveMember = (fromIdx: number, toIdx: number, memberId: string) => {
    if (fromIdx === toIdx) return;
    const newGroups = [...groups];
    const memberIndex = newGroups[fromIdx].findIndex(m => m.id === memberId);
    if (memberIndex === -1) return;
    const [member] = newGroups[fromIdx].splice(memberIndex, 1);
    newGroups[toIdx].push(member);
    setGroups(newGroups.filter(g => g.length > 0));
  };

  const mergeGroups = (idx1: number, idx2: number) => {
    if (idx1 === idx2) return;
    const newGroups = [...groups];
    newGroups[idx1] = [...newGroups[idx1], ...newGroups[idx2]];
    newGroups.splice(idx2, 1);
    setGroups(newGroups);
  };

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col font-sans selection:bg-indigo-100 selection:text-indigo-900">
      {/* Top Navigation Bar */}
      <header className="h-16 bg-white/80 backdrop-blur-md border-b border-slate-200 flex items-center justify-between px-8 shrink-0 shadow-sm z-30 sticky top-0">
        <div className="flex items-center gap-3">
          <div className="w-9 h-9 bg-gradient-to-br from-indigo-600 to-violet-700 rounded-xl flex items-center justify-center text-white font-bold text-lg shadow-lg shadow-indigo-200">
            <Sparkles className="w-5 h-5" />
          </div>
          <span className="text-xl font-black tracking-tight text-slate-800 bg-clip-text text-transparent bg-gradient-to-r from-slate-900 to-slate-600">分組工具大師</span>
        </div>
        <div className="flex items-center gap-4">
          <nav className="hidden md:flex gap-8 text-sm font-bold text-slate-400">
            <a href="#" className="text-indigo-600 border-b-2 border-indigo-600 py-5">工作台</a>
            <a href="#" className="hover:text-slate-800 py-5 transition-colors tracking-wide">歷史紀錄</a>
          </nav>
          <div className="h-4 w-[1px] bg-slate-200 mx-2 hidden md:block" />
          <div className="flex items-center gap-2">
            {groups.length > 0 && (
              <>
                <button 
                  onClick={exportExcel}
                  className="bg-emerald-500 hover:bg-emerald-600 active:scale-95 text-white px-5 py-2.5 rounded-xl text-sm font-bold shadow-lg shadow-emerald-200/50 flex items-center gap-2 transition-all cursor-pointer whitespace-nowrap"
                >
                  <FileSpreadsheet className="w-4 h-4" />
                  匯出 Excel
                </button>
                <button 
                  onClick={copyResults}
                  className="bg-white border border-slate-200 hover:bg-slate-50 active:scale-95 text-slate-700 px-5 py-2.5 rounded-xl text-sm font-bold shadow-sm flex items-center gap-2 transition-all cursor-pointer whitespace-nowrap"
                >
                  {copied ? <Check className="w-4 h-4 text-emerald-500" /> : <Copy className="w-4 h-4 text-slate-400" />}
                  {copied ? '已複製' : '複製結果'}
                </button>
              </>
            )}
          </div>
        </div>
      </header>

      {/* Main Content Area */}
      <main className="flex-1 flex flex-col lg:flex-row gap-8 p-6 lg:p-10 overflow-hidden max-w-[1600px] mx-auto w-full">
        {/* Left Sidebar: Member Pool */}
        <aside className="w-full lg:w-[400px] bg-white rounded-[2.5rem] border border-slate-200 flex flex-col shadow-xl shadow-slate-200/40 relative group/aside z-10">
          <div className="p-8 border-b border-slate-100 bg-slate-50/30 rounded-t-[2.5rem]">
            <div className="flex justify-between items-center mb-6 relative">
              <div className="flex items-center gap-3 group cursor-help"
                   onMouseEnter={() => setShowExample(true)}
                   onMouseLeave={() => setShowExample(false)}>
                <div className="w-6 h-6 bg-indigo-50 rounded-lg flex items-center justify-center">
                   <Users className="w-3.5 h-3.5 text-indigo-500" />
                </div>
                <h2 className="font-black text-slate-800 tracking-tight">人員名冊</h2>
                <Info className="w-4 h-4 text-slate-300 group-hover:text-indigo-500 transition-colors" />
                
                <AnimatePresence>
                  {showExample && (
                    <motion.div
                      initial={{ opacity: 0, y: 15, scale: 0.95 }}
                      animate={{ opacity: 1, y: 0, scale: 1 }}
                      exit={{ opacity: 0, y: 15, scale: 0.95 }}
                      className="absolute top-10 left-0 w-72 bg-slate-900 text-white p-5 rounded-[2rem] shadow-2xl z-50 text-xs leading-relaxed border border-slate-800"
                    >
                      <div className="px-3 py-1 bg-indigo-500/20 text-indigo-400 rounded-full w-fit mb-3 font-bold scale-90 origin-left tracking-widest">FORMAT EXAMPLE</div>
                      <p className="text-slate-400 mb-3 px-1">Excel 標題建議包含：姓名, 代號, 單位</p>
                      <div className="bg-slate-800/80 p-3 rounded-xl font-mono text-[11px] border border-white/5 space-y-1">
                         <div className="text-white/40">姓名 | 工號 | 單位</div>
                         <div>王小明 | A001 | 技術科</div>
                         <div>李大維 | B023 | 展業組</div>
                      </div>
                    </motion.div>
                  )}
                </AnimatePresence>
              </div>
              <div className="flex items-center gap-2">
                <span className="bg-white border border-slate-200 text-slate-600 text-[10px] px-3 py-1.5 rounded-full font-black tracking-widest shadow-sm">{members.length} MEMBERS</span>
              </div>
            </div>
            
            {/* Direct Input Fields */}
            <div className="space-y-4 mb-8 bg-white p-6 rounded-3xl border border-slate-100 shadow-sm ring-4 ring-slate-50/50">
              <div className="grid grid-cols-2 gap-3">
                <div className="space-y-1.5">
                  <label className="text-[10px] font-bold text-slate-400 ml-1">姓名</label>
                  <input
                    type="text"
                    value={newMember.name}
                    onChange={e => setNewMember({ ...newMember, name: e.target.value })}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-3 text-sm focus:ring-4 focus:ring-indigo-100 focus:border-indigo-500 outline-none transition-all placeholder:text-slate-300 font-medium"
                  />
                </div>
                <div className="space-y-1.5">
                  <label className="text-[10px] font-bold text-slate-400 ml-1">代號</label>
                  <input
                    type="text"
                    value={newMember.empId}
                    onChange={e => setNewMember({ ...newMember, empId: e.target.value })}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-3 text-sm focus:ring-4 focus:ring-indigo-100 focus:border-indigo-500 outline-none transition-all placeholder:text-slate-300 font-medium font-mono"
                  />
                </div>
              </div>
              <div className="space-y-1.5">
                <label className="text-[10px] font-bold text-slate-400 ml-1">單位部門</label>
                <div className="flex gap-3">
                  <input
                    type="text"
                    value={newMember.unit}
                    onChange={e => setNewMember({ ...newMember, unit: e.target.value })}
                    className="flex-1 bg-slate-50 border border-slate-200 rounded-xl px-4 py-3 text-sm focus:ring-4 focus:ring-indigo-100 focus:border-indigo-500 outline-none transition-all placeholder:text-slate-300 font-medium"
                  />
                  <button
                    onClick={addMember}
                    className="bg-slate-900 text-white w-12 h-12 flex items-center justify-center rounded-xl hover:bg-slate-800 transition-all shadow-lg shadow-slate-900/20 active:scale-95 shrink-0"
                  >
                    <Plus className="w-6 h-6" />
                  </button>
                </div>
              </div>
            </div>

            <div className="grid grid-cols-2 gap-3">
              <input type="file" ref={fileInputRef} onChange={handleFileUpload} accept=".xlsx,.xls,.csv" className="hidden" />
              <button onClick={() => fileInputRef.current?.click()} className="flex items-center justify-center gap-2 py-4 bg-emerald-50 text-emerald-700 font-black text-xs rounded-2xl border border-emerald-100 hover:bg-emerald-100 transition-all active:scale-[0.97] shadow-sm tracking-widest">
                <FileSpreadsheet className="w-4 h-4" />
                EXCEL 匯入
              </button>
              <button onClick={importExample} className="flex items-center justify-center gap-2 py-4 bg-indigo-50 text-indigo-700 font-black text-xs rounded-2xl border border-indigo-100 hover:bg-indigo-100 transition-all active:scale-[0.97] shadow-sm tracking-widest">
                <Plus className="w-4 h-4" />
                範例匯入
              </button>
            </div>
          </div>

          <div className="flex-1 overflow-y-auto p-6 custom-scrollbar space-y-3 max-h-[400px] lg:max-h-none">
            <AnimatePresence>
              {members.length === 0 ? (
                <div className="h-full flex flex-col items-center justify-center py-20 text-slate-300 opacity-50">
                  <div className="p-4 bg-slate-50 rounded-3xl mb-4 border border-slate-100">
                    <ListChecks className="w-10 h-10" />
                  </div>
                  <span className="text-sm font-bold tracking-widest">名冊尚無資料</span>
                </div>
              ) : (
                members.map(m => (
                  <motion.div
                    key={m.id}
                    initial={{ opacity: 0, scale: 0.95 }}
                    animate={{ opacity: 1, scale: 1 }}
                    exit={{ opacity: 0, scale: 0.95 }}
                    className="group flex items-center justify-between p-4 bg-white hover:bg-slate-50 border border-slate-100 hover:border-slate-200 rounded-3xl transition-all shadow-sm"
                  >
                    <div className="flex items-center gap-4 overflow-hidden">
                      <div className="shrink-0 w-10 h-10 rounded-2xl bg-gradient-to-tr from-slate-100 to-indigo-50 flex items-center justify-center text-indigo-600 font-black text-sm border border-white">
                        {m.name.charAt(0)}
                      </div>
                      <div className="overflow-hidden">
                        <p className="text-sm font-black text-slate-800 truncate">{m.name}</p>
                        <p className="text-[10px] font-bold text-slate-400 truncate tracking-wide flex items-center gap-1.5 mt-0.5">
                          <span className="bg-slate-100 px-1.5 py-0.5 rounded uppercase font-mono">{m.empId}</span>
                          <span className="w-1 h-1 rounded-full bg-slate-200" />
                          <span>{m.unit}</span>
                        </p>
                      </div>
                    </div>
                    <button 
                      onClick={() => removeMember(m.id)}
                      className="p-2.5 text-slate-300 border border-transparent hover:border-red-100 hover:bg-red-50 hover:text-red-500 rounded-xl transition-all"
                    >
                      <X className="w-4 h-4" />
                    </button>
                  </motion.div>
                ))
              )}
            </AnimatePresence>
          </div>
        </aside>

        {/* Center: Grouping Canvas */}
        <section className="flex-1 flex flex-col gap-8 overflow-hidden">
          <header className="flex items-center justify-between px-4">
            <h2 className="text-xs font-black text-slate-400 uppercase tracking-[0.3em] flex items-center gap-3">
              <Grid2X2 className="w-4 h-4 text-indigo-400" />
              分組實驗工作區
            </h2>
            <div className="flex items-center gap-2">
              <div className={`w-2 h-2 rounded-full ${members.length > 0 ? 'bg-emerald-500 animate-pulse' : 'bg-slate-300'}`} />
              <span className="text-[10px] text-slate-400 font-black uppercase tracking-widest">{members.length > 0 ? 'READY' : 'WAITING'}</span>
            </div>
          </header>
          
          <div className="flex-1 overflow-y-auto pr-4 custom-scrollbar">
            <AnimatePresence>
              {isGenerating && (
                <motion.div 
                  initial={{ opacity: 0 }}
                  animate={{ opacity: 1 }}
                  exit={{ opacity: 0 }}
                  className="fixed inset-0 z-[100] flex items-center justify-center p-6"
                >
                  {/* Glassmorphic Backdrop */}
                  <div className="absolute inset-0 bg-slate-900/80 backdrop-blur-xl" />
                  
                  {/* Floating Particles */}
                  <div className="absolute inset-0 overflow-hidden pointer-events-none">
                    {[...Array(24)].map((_, i) => (
                      <motion.div
                        key={i}
                        animate={{ 
                          y: [-20, -1200],
                          x: [0, (Math.random() - 0.5) * 400],
                          opacity: [0, 0.4, 0],
                          scale: [0.5, 1.5, 0.5]
                        }}
                        transition={{ 
                          duration: Math.random() * 2 + 2, 
                          repeat: Infinity, 
                          delay: Math.random() * 2,
                          ease: "easeOut"
                        }}
                        className="absolute bottom-[-100px] text-indigo-400 font-black text-2xl"
                        style={{ left: `${Math.random() * 100}%` }}
                      >
                        {Math.random() > 0.5 ? <Sparkles className="w-6 h-6" /> : "01"}
                      </motion.div>
                    ))}
                  </div>

                  {/* Main Cinematic Wheel Card */}
                  <motion.div 
                    initial={{ scale: 0.8, y: 50, rotateX: 30 }}
                    animate={{ scale: 1, y: 0, rotateX: 0 }}
                    exit={{ scale: 1.1, opacity: 0 }}
                    className="relative z-10 w-full max-w-2xl bg-white rounded-[4rem] p-12 shadow-[0_50px_100px_-20px_rgba(0,0,0,0.5)] border border-white/20 overflow-hidden flex flex-col items-center"
                    style={{ perspective: "1000px" }}
                  >
                    {/* Interior Decorative Glow */}
                    <div className="absolute top-[-50%] left-[-50%] w-[200%] h-[200%] bg-gradient-to-br from-indigo-500/10 via-transparent to-emerald-500/10 animate-pulse pointer-events-none" />
                    
                    <div className="relative mb-14 drop-shadow-[0_20px_50px_rgba(79,70,229,0.3)]">
                      {/* The Wheel Visuals */}
                      <motion.div 
                        style={{ rotate: wheelRotation }}
                        className="w-64 h-64 rounded-full border-[16px] border-slate-900 relative flex items-center justify-center bg-white overflow-hidden shadow-inner"
                      >
                        {/* Wheel Spokes */}
                        {[...Array(24)].map((_, i) => (
                          <div 
                            key={i} 
                            className="absolute w-full h-[1px] bg-slate-100" 
                            style={{ transform: `rotate(${i * 15}deg)` }}
                          />
                        ))}
                        {/* Center Hub */}
                        <div className="w-16 h-16 rounded-full bg-slate-900 border-4 border-indigo-500 flex items-center justify-center z-10 shadow-lg">
                           <Shuffle className="w-6 h-6 text-indigo-400" />
                        </div>
                      </motion.div>
                      
                      {/* Pointer Indicator */}
                      <div className="absolute top-[-25px] left-1/2 -translate-x-1/2 w-10 h-12 bg-indigo-600 z-20 shadow-[0_10px_20px_rgba(0,0,0,0.3)]" 
                           style={{ clipPath: "polygon(50% 100%, 0 0, 100% 0)" }}
                      />
                    </div>

                    <div className="text-center relative z-20 w-full px-4">
                      <div className="h-32 mb-8 flex items-center justify-center overflow-visible">
                        <AnimatePresence mode="wait">
                          <motion.div 
                            key={shuffleDisplay}
                            initial={{ scale: 2, y: 60, opacity: 0, filter: "blur(20px)", rotateX: 45 }}
                            animate={{ scale: 1, y: 0, opacity: 1, filter: "blur(0px)", rotateX: 0 }}
                            exit={{ scale: 0.5, y: -40, opacity: 0, filter: "blur(10px)" }}
                            transition={{ 
                              type: "spring", 
                              damping: 15, 
                              stiffness: 120,
                              mass: 0.8
                            }}
                            className="text-7xl md:text-8xl font-black tracking-tighter bg-clip-text text-transparent bg-gradient-to-br from-indigo-600 via-violet-600 to-indigo-900 drop-shadow-sm pb-2"
                          >
                            {shuffleDisplay}
                          </motion.div>
                        </AnimatePresence>
                      </div>
                      
                      <div className="flex flex-col items-center gap-4">
                        <div className="px-10 py-3.5 bg-slate-900 rounded-full text-white text-[10px] font-black tracking-[0.5em] uppercase shadow-[0_20px_40px_-5px_rgba(0,0,0,0.3)] border border-slate-700">
                          Processing Protocols...
                        </div>
                        <p className="text-indigo-500 font-black animate-pulse uppercase tracking-[0.3em] text-[10px]">大轉盤智能隨機計算中</p>
                      </div>
                    </div>
                  </motion.div>
                </motion.div>
              )}
            </AnimatePresence>

            {groups.length === 0 ? (
              <div className="h-full border-2 border-dashed border-slate-200 rounded-[3rem] flex flex-col items-center justify-center p-16 text-center bg-white/40 group/empty">
                <div className="w-24 h-24 bg-white rounded-[2rem] flex items-center justify-center mb-10 shadow-xl border border-slate-100 group-hover/empty:scale-110 transition-transform cursor-pointer">
                  <Grid2X2 className="w-10 h-10 text-slate-200 group-hover/empty:text-indigo-400 transition-colors" />
                </div>
                <h3 className="font-black text-slate-800 mb-4 text-2xl tracking-tight">準備好釋放數據能量嗎？</h3>
                <p className="text-slate-400 text-sm max-w-[320px] leading-relaxed font-medium">
                  導入名單後，點擊右側的隨機按鈕。我們的算法將為您創造出完美的團隊組合。
                </p>
                <button 
                  onClick={handleGroup}
                  disabled={members.length === 0}
                  className="mt-12 text-indigo-600 font-black uppercase text-xs tracking-widest hover:tracking-[0.4em] transition-all disabled:hidden"
                >
                  START NOW →
                </button>
              </div>
            ) : (
              <motion.div 
                layout
                className="grid grid-cols-1 xl:grid-cols-2 gap-8 pb-12"
              >
                <AnimatePresence mode="popLayout">
                  {groups.map((group, groupIndex) => (
                    <motion.div
                      key={`group-${groupIndex}`}
                      layout
                      initial={{ opacity: 0, y: 30 }}
                      animate={{ opacity: 1, y: 0 }}
                      className="bg-white rounded-[2.5rem] border border-slate-200 p-8 flex flex-col shadow-lg shadow-slate-200/20 hover:border-indigo-200 transition-all group/card relative"
                      onDragOver={(e) => {
                        e.preventDefault();
                        e.currentTarget.classList.add('ring-4', 'ring-indigo-100');
                      }}
                      onDragLeave={(e) => {
                        e.currentTarget.classList.remove('ring-4', 'ring-indigo-100');
                      }}
                      onDrop={(e) => {
                        e.preventDefault();
                        e.currentTarget.classList.remove('ring-4', 'ring-indigo-100');
                        if (draggedMember) {
                          moveMember(draggedMember.fromGroup, groupIndex, draggedMember.memberId);
                          setDraggedMember(null);
                        }
                      }}
                    >
                      <div className="flex justify-between items-center mb-8">
                        <div className="flex items-center gap-4">
                          <div className="w-11 h-11 rounded-2xl bg-slate-900 text-white flex items-center justify-center font-black text-lg shadow-lg shadow-slate-900/20">
                            {groupIndex + 1}
                          </div>
                          <div>
                            <h3 className="font-black text-slate-800">第 {groupIndex + 1} 團隊</h3>
                            <p className="text-[10px] font-bold text-slate-300 uppercase tracking-widest mt-0.5">TEAM ALPHA-BRAVO</p>
                          </div>
                        </div>
                        <div className="flex items-center gap-3">
                          {groupIndex > 0 && (
                            <button 
                              onClick={() => mergeGroups(groupIndex - 1, groupIndex)}
                              className="p-2 text-slate-400 hover:text-indigo-600 hover:bg-indigo-50 rounded-lg transition-all"
                              title="合併至前一組"
                            >
                              <Merge className="w-4 h-4" />
                            </button>
                          )}
                          <span className="text-[10px] text-indigo-600 font-black uppercase tracking-[0.2em] bg-indigo-50 px-4 py-2 rounded-full border border-indigo-100">
                            {group.length} PERSONS
                          </span>
                        </div>
                      </div>
                      
                      <div className="space-y-3">
                        {group.map((m) => (
                          <motion.div 
                            key={m.id}
                            layout
                            draggable
                            onDragStart={() => setDraggedMember({ memberId: m.id, fromGroup: groupIndex })}
                            className="p-4 bg-slate-50 border border-transparent hover:border-slate-200 hover:bg-white rounded-2xl text-sm flex items-center justify-between group/item cursor-grab active:cursor-grabbing transition-all hover:shadow-md"
                          >
                            <div className="flex items-center gap-4">
                              <GripVertical className="w-4 h-4 text-slate-200 group-hover/item:text-slate-400 transition-colors" />
                              <div>
                                <span className="font-black text-slate-800">{m.name}</span>
                                <div className="text-[10px] font-bold text-slate-400 uppercase tracking-tighter">{m.unit}</div>
                              </div>
                            </div>
                            <div className="flex flex-col items-end">
                              <span className="text-[10px] text-slate-300 font-black font-mono group-hover/item:text-indigo-500 transition-colors">
                                {m.empId}
                              </span>
                            </div>
                          </motion.div>
                        ))}
                      </div>
                    </motion.div>
                  ))}
                </AnimatePresence>
                
                {groups.length > 1 && (
                  <div className="xl:col-span-2 flex items-center justify-center gap-4 py-8 pointer-events-none opacity-40">
                    <ArrowRightLeft className="w-5 h-5 text-slate-300" />
                    <span className="text-xs font-bold text-slate-400 uppercase tracking-widest leading-none">拖曳成員可自由調整組別 • 點擊圖示可合併相鄰組別</span>
                  </div>
                )}
              </motion.div>
            )}
          </div>
        </section>

        {/* Right Sidebar: Configuration */}
        <aside className="w-full lg:w-96 flex flex-col gap-6 shrink-0 z-10">
          {/* Quick Clear Center */}
          <div className="grid grid-cols-2 gap-4">
            <button
              onClick={clearAll}
              className="flex items-center justify-center gap-3 py-5 px-4 bg-slate-900 border border-slate-800 rounded-[1.5rem] text-slate-100 font-black text-xs hover:bg-slate-800 transition-all active:scale-[0.97] shadow-xl tracking-widest whitespace-nowrap overflow-hidden"
            >
              <Trash2 className="w-4 h-4 text-red-400 shrink-0" />
              開新名單
            </button>
            <button
              onClick={() => { setMembers([]); }}
              className="flex items-center justify-center gap-3 py-5 px-4 bg-white border border-slate-200 rounded-[1.5rem] text-slate-700 font-black text-xs hover:bg-slate-50 transition-all active:scale-[0.97] shadow-lg shadow-slate-200/30 tracking-widest whitespace-nowrap overflow-hidden"
            >
              <X className="w-4 h-4 text-slate-400 shrink-0" />
              清空池子
            </button>
          </div>

          <div className="bg-slate-900 rounded-[3rem] p-10 text-white shadow-2xl relative overflow-hidden group/config border border-slate-800">
            <div className="absolute top-[-20%] right-[-10%] w-60 h-60 bg-indigo-500/20 blur-[80px] rounded-full group-hover/config:scale-110 transition-transform duration-1000" />
            <h2 className="text-2xl font-black mb-10 flex items-center gap-4 relative z-10">
              <div className="p-3 bg-indigo-500/20 rounded-2xl text-indigo-400 border border-indigo-500/10">
                <Shuffle className="w-6 h-6" />
              </div>
              全智隨機設定
            </h2>
            <div className="space-y-10 relative z-10">
              <div>
                <label className="text-[10px] text-indigo-400/60 font-black uppercase tracking-[0.3em] mb-4 block">分組算法模式</label>
                <div className="grid grid-cols-2 gap-3 p-1.5 bg-slate-800/80 backdrop-blur-md rounded-2xl border border-white/5">
                  <button
                    onClick={() => setMode('numGroups')}
                    className={`py-4 px-4 rounded-xl text-xs font-black transition-all ${
                      mode === 'numGroups' 
                      ? 'bg-indigo-600 text-white shadow-lg shadow-indigo-500/30' 
                      : 'text-slate-400 hover:text-white'
                    }`}
                  >
                    按組數分配
                  </button>
                  <button
                    onClick={() => setMode('sizeGroups')}
                    className={`py-4 px-4 rounded-xl text-xs font-black transition-all ${
                      mode === 'sizeGroups' 
                      ? 'bg-indigo-600 text-white shadow-lg shadow-indigo-500/30' 
                      : 'text-slate-400 hover:text-white'
                    }`}
                  >
                    按人數分配
                  </button>
                </div>
              </div>

              <div className="space-y-6">
                <div className="flex items-center justify-between">
                  <label className="text-[10px] text-indigo-400/60 font-black uppercase tracking-[0.3em] block">
                    {mode === 'numGroups' ? '目標組數數量' : '每組期望人數'}
                  </label>
                  <div className="flex items-center gap-3">
                    <input
                      type="number"
                      min="1"
                      max={Math.max(100, members.length)}
                      value={value}
                      onChange={(e) => setValue(parseInt(e.target.value) || 0)}
                      className="w-20 text-center text-lg font-mono font-black text-indigo-400 bg-white/5 px-2 py-2 rounded-2xl border border-white/10 outline-none focus:border-indigo-500/50 transition-all hover:bg-white/10"
                    />
                  </div>
                </div>
                <div className="relative pt-2">
                   <input
                    type="range"
                    min="1"
                    max={Math.max(2, members.length)}
                    value={value}
                    onChange={(e) => setValue(parseInt(e.target.value))}
                    className="w-full h-1.5 bg-slate-800 rounded-full appearance-none cursor-pointer accent-indigo-500"
                  />
                  <div className="flex justify-between text-[10px] text-slate-600 font-black font-mono mt-3 uppercase tracking-tighter">
                    <span>min: 1</span>
                    <span>max: {members.length}</span>
                  </div>
                </div>
              </div>

              <button
                onClick={handleGroup}
                disabled={members.length === 0 || isGenerating}
                className="w-full bg-indigo-500 hover:bg-indigo-400 disabled:opacity-10 py-6 rounded-[2rem] font-black mt-4 shadow-[0_25px_60px_-15px_rgba(79,70,229,0.5)] transition-all active:scale-[0.98] flex items-center justify-center gap-4 text-sm tracking-widest group/btn"
              >
                <motion.div 
                  animate={isGenerating ? { rotate: 360 } : {}} 
                  transition={{ duration: 1, repeat: Infinity, ease: 'linear' }}
                >
                  <Shuffle className="w-6 h-6 group-hover/btn:scale-110 transition-transform" />
                </motion.div>
                大轉盤隨機起動
              </button>
            </div>
          </div>

          <div className="bg-white rounded-[3rem] border border-slate-200 p-10 flex-1 shadow-2xl shadow-slate-200/50 relative overflow-hidden group/stats">
            <div className="absolute bottom-[-10%] left-[-10%] w-40 h-40 bg-indigo-50 blur-[50px] rounded-full group-hover/stats:scale-125 transition-transform duration-1000" />
            <h3 className="font-black text-slate-800 mb-10 flex items-center gap-4 relative z-10">
              <div className="p-3 bg-indigo-50 rounded-2xl border border-indigo-100">
                <ListChecks className="w-5 h-5 text-indigo-600" />
              </div>
              數據分析中心
            </h3>
            <div className="space-y-8 relative z-10">
              <div className="flex justify-between items-end border-b border-slate-100 pb-5">
                <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest">總計納入人數</span>
                <span className="text-3xl font-mono font-black text-slate-800 leading-none">{members.length}</span>
              </div>
              <div className="flex justify-between items-end border-b border-slate-100 pb-5">
                <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest">有效作業組別</span>
                <span className="text-3xl font-mono font-black text-indigo-600 leading-none">{groups.length}</span>
              </div>
              <div className="flex justify-between items-end border-b border-slate-100 pb-5">
                <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest">分組平均係數</span>
                <span className="text-3xl font-mono font-black text-slate-800 leading-none font-bold">
                  {groups.length > 0 ? (members.length / groups.length).toFixed(1) : '0.0'}
                </span>
              </div>
            </div>
            <div className="mt-12 relative z-10">
               <div className="w-full bg-slate-100 h-3 rounded-full overflow-hidden border border-slate-200 shadow-inner">
                 <motion.div 
                   initial={{ width: 0 }}
                   animate={{ width: `${groups.length > 0 ? 100 : 0}%` }}
                   className="bg-gradient-to-r from-indigo-500 to-indigo-700 h-full shadow-[0_0_10px_rgba(79,70,229,0.3)]"
                 />
               </div>
               <p className="text-[11px] font-black text-slate-400 mt-5 text-center uppercase tracking-[0.4em] leading-none opacity-60">
                 {groups.length > 0 ? 'SYSTEM STATUS: OPTIMIZED' : 'WATING FOR INPUT CORE'}
               </p>
            </div>
          </div>
        </aside>
      </main>
      
      <footer className="py-10 text-center">
         <p className="text-[10px] font-black text-slate-300 uppercase tracking-[0.5em]">Global GroupMaster Protocol • v2.5 Cinematic Update</p>
      </footer>

      {/* Global CSS for pointer custom scrollbar */}
      <style>{`
        .custom-scrollbar::-webkit-scrollbar {
          width: 6px;
        }
        .custom-scrollbar::-webkit-scrollbar-track {
          background: transparent;
        }
        .custom-scrollbar::-webkit-scrollbar-thumb {
          background: #e2e8f0;
          border-radius: 10px;
        }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover {
          background: #cbd5e1;
        }
      `}</style>
    </div>
  );
}

