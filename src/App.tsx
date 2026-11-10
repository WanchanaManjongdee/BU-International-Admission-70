import React, { useState, useEffect, useMemo, useRef } from 'react';
import {
  FileSpreadsheet,
  Upload,
  Shield,
  ShieldCheck,
  Search,
  Filter,
  Download,
  AlertCircle,
  Clock,
  UserCheck,
  FileCheck2,
  CheckCircle2,
  Calendar,
  Users,
  Layers,
  ChevronLeft,
  ChevronRight,
  Info,
  Building2,
  X,
  FileText,
  UserPlus
} from 'lucide-react';
import { ApplicationRecord, sample50Applications } from './sampleData';
import {
  parseCSVText,
  parseExcelFile,
  computeDashboardStats,
  maskStudentName,
  maskEmail,
  exportToCSV,
  getAdmissionOwnerDisplayName,
  DashboardStats
} from './dataUtils';

type DataSourceMode = 'hubspot_csv' | 'sample_50' | 'uploaded';
type AgingFilter = 'all' | 'less_1w' | '2_to_4w' | 'more_4w';
type OwnerViewMode = 'admission' | 'application';

export default function App() {
  // State management
  const [dataSource, setDataSource] = useState<DataSourceMode>('hubspot_csv');
  const [allRecords, setAllRecords] = useState<ApplicationRecord[]>(sample50Applications);
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [uploadedFileName, setUploadedFileName] = useState<string>('');
  
  // Privacy and Display
  const [isPrivacyMasked, setIsPrivacyMasked] = useState<boolean>(true);
  const [showGuideModal, setShowGuideModal] = useState<boolean>(false);
  
  // Owner Distribution View Mode (Admission Owner vs Application Owner)
  const [ownerViewMode, setOwnerViewMode] = useState<OwnerViewMode>('admission');

  // Filters
  const [selectedYear, setSelectedYear] = useState<string>('2570'); // '2570', 'all', '2569'
  const [selectedAdmissionOwner, setSelectedAdmissionOwner] = useState<string>('all');
  const [selectedApplicationOwner, setSelectedApplicationOwner] = useState<string>('all');
  const [selectedStage, setSelectedStage] = useState<string>('all');
  const [selectedAging, setSelectedAging] = useState<AgingFilter>('all');
  const [searchQuery, setSearchQuery] = useState<string>('');
  
  // Pagination & Sorting
  const [currentPage, setCurrentPage] = useState<number>(1);
  const [pageSize, setPageSize] = useState<number>(10);
  const [sortField, setSortField] = useState<'createDate' | 'agingDays' | 'admissionOwner' | 'owner' | 'stage'>('createDate');
  const [sortAsc, setSortAsc] = useState<boolean>(false);

  // File Upload Ref
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Load HubSpot CRM dataset by default on mount
  useEffect(() => {
    async function loadHubSpotData() {
      setIsLoading(true);
      try {
        const response = await fetch('/assets/aistudio/hubspot-crm-exports-all-applications-2026-11-10_25c55717.csv');
        if (response.ok) {
          const csvText = await response.text();
          const parsed = parseCSVText(csvText);
          if (parsed.length > 0) {
            setAllRecords(parsed);
            setDataSource('hubspot_csv');
            setIsLoading(false);
            return;
          }
        }
      } catch (err) {
        console.warn('Could not load HubSpot CSV from server, falling back to 50 sample records:', err);
      }
      // Fallback
      setAllRecords(sample50Applications);
      setDataSource('sample_50');
      setIsLoading(false);
    }

    loadHubSpotData();
  }, []);

  // Handle Switch to Sample Data
  const handleSwitchToSample = () => {
    setIsLoading(true);
    setAllRecords(sample50Applications);
    setDataSource('sample_50');
    setCurrentPage(1);
    setIsLoading(false);
  };

  // Handle Switch to Real HubSpot File
  const handleSwitchToHubSpot = async () => {
    setIsLoading(true);
    try {
      const response = await fetch('/assets/aistudio/hubspot-crm-exports-all-applications-2026-11-10_25c55717.csv');
      if (response.ok) {
        const csvText = await response.text();
        const parsed = parseCSVText(csvText);
        setAllRecords(parsed);
        setDataSource('hubspot_csv');
      }
    } catch (e) {
      console.error(e);
    } finally {
      setIsLoading(false);
      setCurrentPage(1);
    }
  };

  // Handle Custom File Upload (Excel or CSV)
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsLoading(true);
    setUploadedFileName(file.name);
    const reader = new FileReader();

    if (file.name.endsWith('.csv')) {
      reader.onload = (evt) => {
        const text = evt.target?.result as string;
        const parsed = parseCSVText(text);
        setAllRecords(parsed);
        setDataSource('uploaded');
        setIsLoading(false);
        setCurrentPage(1);
      };
      reader.readAsText(file);
    } else {
      // Excel file .xlsx or .xls
      reader.onload = (evt) => {
        const buffer = evt.target?.result as ArrayBuffer;
        try {
          const parsed = parseExcelFile(buffer);
          setAllRecords(parsed);
          setDataSource('uploaded');
        } catch (err) {
          alert('ไม่สามารถอ่านไฟล์ Excel ได้ กรุณาตรวจสอบว่าเป็นไฟล์ .xlsx หรือ .csv ที่ถูกต้อง');
          console.error(err);
        } finally {
          setIsLoading(false);
          setCurrentPage(1);
        }
      };
      reader.readAsArrayBuffer(file);
    }
  };

  // Filter Records by Year first (Bangkok University Year 2570 filter)
  const yearFilteredRecords = useMemo(() => {
    if (selectedYear === 'all') return allRecords;
    if (selectedYear === '2570') {
      return allRecords.filter(r => {
        return (
          r.applicationYear === '70' ||
          r.applicationYear === '2570' ||
          r.intakeYearApp === '2027' ||
          r.intakeTerm?.includes('2027')
        );
      });
    }
    if (selectedYear === '2569') {
      return allRecords.filter(r => {
        return (
          r.applicationYear === '69' ||
          r.applicationYear === '2569' ||
          r.intakeYearApp === '2026' ||
          r.intakeTerm?.includes('2026')
        );
      });
    }
    return allRecords;
  }, [allRecords, selectedYear]);

  // Overall Stats based on the active Year scope
  const stats: DashboardStats = useMemo(() => {
    return computeDashboardStats(yearFilteredRecords);
  }, [yearFilteredRecords]);

  // Unique list of Admission Owners and Application Owners
  const availableAdmissionOwners = useMemo(() => {
    const set = new Set<string>();
    yearFilteredRecords.forEach(r => {
      const trimmed = (r.admissionOwner || '').trim();
      if (trimmed) set.add(trimmed);
    });
    return Array.from(set).sort();
  }, [yearFilteredRecords]);

  const availableApplicationOwners = useMemo(() => {
    const set = new Set<string>();
    yearFilteredRecords.forEach(r => {
      if (r.owner) set.add(r.owner);
    });
    return Array.from(set).sort();
  }, [yearFilteredRecords]);

  const availableStages = useMemo(() => {
    const set = new Set<string>();
    yearFilteredRecords.forEach(r => {
      if (r.stage) set.add(r.stage);
    });
    return Array.from(set).sort();
  }, [yearFilteredRecords]);

  // Fully Filtered Records for Table & Drill-down
  const filteredRecords = useMemo(() => {
    return yearFilteredRecords.filter(r => {
      // Admission Owner Filter
      if (selectedAdmissionOwner !== 'all') {
        const rawAdm = (r.admissionOwner || '').trim();
        if (selectedAdmissionOwner === 'Unassigned') {
          if (rawAdm && rawAdm.toLowerCase() !== 'unassigned') return false;
        } else {
          if (rawAdm !== selectedAdmissionOwner) return false;
        }
      }
      // Application Owner Filter
      if (selectedApplicationOwner !== 'all' && r.owner !== selectedApplicationOwner) {
        return false;
      }
      // Stage Filter
      if (selectedStage !== 'all' && r.stage !== selectedStage) {
        return false;
      }
      // Aging Filter (applies to Review Document stage)
      if (selectedAging !== 'all') {
        if (r.stage !== 'Review Document') return false;
        if (r.agingBucket !== selectedAging) return false;
      }
      // Search Query
      if (searchQuery.trim()) {
        const query = searchQuery.toLowerCase();
        const matchName = r.name.toLowerCase().includes(query);
        const matchEmail = r.email.toLowerCase().includes(query);
        const matchId = r.id.toLowerCase().includes(query);
        const matchAdmOwner = (r.admissionOwner || '').toLowerCase().includes(query);
        const matchAppOwner = r.owner.toLowerCase().includes(query);
        const matchStage = r.stage.toLowerCase().includes(query);
        if (!matchName && !matchEmail && !matchId && !matchAdmOwner && !matchAppOwner && !matchStage) {
          return false;
        }
      }
      return true;
    });
  }, [yearFilteredRecords, selectedAdmissionOwner, selectedApplicationOwner, selectedStage, selectedAging, searchQuery]);

  // Sorted and Paginated Records
  const sortedRecords = useMemo(() => {
    return [...filteredRecords].sort((a, b) => {
      let valA: string | number = '';
      let valB: string | number = '';

      if (sortField === 'createDate') {
        valA = a.createDate || '';
        valB = b.createDate || '';
      } else if (sortField === 'agingDays') {
        valA = a.agingDays ?? -1;
        valB = b.agingDays ?? -1;
      } else if (sortField === 'admissionOwner') {
        valA = a.admissionOwner || 'zzz';
        valB = b.admissionOwner || 'zzz';
      } else if (sortField === 'owner') {
        valA = a.owner || '';
        valB = b.owner || '';
      } else if (sortField === 'stage') {
        valA = a.stage || '';
        valB = b.stage || '';
      }

      if (valA < valB) return sortAsc ? -1 : 1;
      if (valA > valB) return sortAsc ? 1 : -1;
      return 0;
    });
  }, [filteredRecords, sortField, sortAsc]);

  const totalPages = Math.ceil(sortedRecords.length / pageSize) || 1;
  const paginatedRecords = useMemo(() => {
    const start = (currentPage - 1) * pageSize;
    return sortedRecords.slice(start, start + pageSize);
  }, [sortedRecords, currentPage, pageSize]);

  // Reset to page 1 on filter changes
  useEffect(() => {
    setCurrentPage(1);
  }, [selectedYear, selectedAdmissionOwner, selectedApplicationOwner, selectedStage, selectedAging, searchQuery]);

  // Color helper for stages
  const getStageBadgeStyle = (stage: string) => {
    switch (stage) {
      case 'Review Document':
        return 'text-amber-800 bg-amber-50 border border-amber-200';
      case 'Wait for Language Result':
        return 'text-blue-800 bg-blue-50 border border-blue-200';
      case 'Invoiced':
        return 'text-purple-800 bg-purple-50 border border-purple-200';
      case 'Confirmed Payment':
        return 'text-emerald-800 bg-emerald-50 border border-emerald-200';
      case 'Scholarship':
        return 'text-indigo-800 bg-indigo-50 border border-indigo-200';
      case 'Closed Lost':
        return 'text-slate-600 bg-slate-100 border border-slate-200';
      default:
        return 'text-slate-700 bg-slate-50 border border-slate-200';
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 pb-16">
      {/* Hidden file input */}
      <input
        type="file"
        ref={fileInputRef}
        onChange={handleFileUpload}
        accept=".csv, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel"
        className="hidden"
      />

      {/* Top Header with Bangkok University Branding */}
      <header className="sticky top-0 z-30 bg-white border-b border-slate-200 shadow-xs">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between gap-4">
          {/* Brand */}
          <div className="flex items-center gap-3 shrink-0">
            <div className="w-10 h-10 rounded-lg bg-blue-900 flex items-center justify-center text-white shadow-xs">
              <Building2 className="w-5 h-5" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <span className="font-bold text-blue-950 text-base sm:text-lg tracking-tight">
                  BU International Admissions
                </span>
                <span className="text-xs font-semibold px-2 py-0.5 rounded-sm bg-blue-100 text-blue-900">
                  ปี 2570
                </span>
              </div>
              <p className="text-xs text-slate-500 hidden sm:block">
                ระบบวิเคราะห์และติดตามสถานะใบสมัครนักศึกษาต่างชาติ มหาวิทยาลัยกรุงเทพ
              </p>
            </div>
          </div>

          {/* Quick Actions */}
          <div className="flex items-center gap-2 sm:gap-3">
            {/* Privacy Shield Toggle */}
            <button
              onClick={() => setIsPrivacyMasked(!isPrivacyMasked)}
              title={isPrivacyMasked ? 'คลิกเพื่อแสดงข้อมูลเต็ม (ปิดโหมด PDPA)' : 'คลิกเพื่อมาสก์ซ่อนข้อมูลนักศึกษา (เปิดโหมด PDPA)'}
              className={`flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium rounded-lg border transition-all ${
                isPrivacyMasked
                  ? 'bg-emerald-50 text-emerald-800 border-emerald-200 hover:bg-emerald-100'
                  : 'bg-amber-50 text-amber-800 border-amber-200 hover:bg-amber-100'
              }`}
            >
              {isPrivacyMasked ? <ShieldCheck className="w-4 h-4 text-emerald-600" /> : <Shield className="w-4 h-4 text-amber-600" />}
              <span className="hidden md:inline">
                {isPrivacyMasked ? 'PDPA Masked' : 'แสดงข้อมูลจริง'}
              </span>
            </button>

            {/* How to use real file guide */}
            <button
              onClick={() => setShowGuideModal(true)}
              className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded-lg hover:bg-slate-50 transition-colors"
            >
              <Info className="w-4 h-4 text-blue-600" />
              <span className="hidden sm:inline">วิธีเปลี่ยนเป็นไฟล์จริง</span>
            </button>

            {/* Upload Button */}
            <button
              onClick={() => fileInputRef.current?.click()}
              className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-white bg-blue-900 hover:bg-blue-800 rounded-lg shadow-xs transition-colors shrink-0"
            >
              <Upload className="w-4 h-4" />
              <span>อัปโหลด Excel / CSV</span>
            </button>
          </div>
        </div>
      </header>

      {/* Main Container */}
      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 pt-6 space-y-6">
        
        {/* Source Switcher & Global Filters Bar */}
        <section className="bg-white border border-slate-200 rounded-xl p-4 shadow-xs">
          <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
            
            {/* Source Segmented Control */}
            <div className="flex flex-wrap items-center gap-2">
              <span className="text-xs font-medium text-slate-500 uppercase tracking-wider mr-1">
                แหล่งข้อมูล:
              </span>
              <div className="inline-flex p-1 bg-slate-100 rounded-lg border border-slate-200/60">
                <button
                  onClick={handleSwitchToHubSpot}
                  className={`px-3 py-1.5 text-xs font-medium rounded-md transition-all ${
                    dataSource === 'hubspot_csv'
                      ? 'bg-white text-blue-950 shadow-xs font-semibold'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  HubSpot CRM ล่าสุด ({allRecords.length > 50 ? `${allRecords.length} แถว` : 'ไฟล์จริง'})
                </button>
                <button
                  onClick={handleSwitchToSample}
                  className={`px-3 py-1.5 text-xs font-medium rounded-md transition-all ${
                    dataSource === 'sample_50'
                      ? 'bg-white text-blue-950 shadow-xs font-semibold'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  ตัวอย่าง 50 แถว
                </button>
                {dataSource === 'uploaded' && (
                  <span className="px-3 py-1.5 text-xs font-semibold rounded-md bg-white text-blue-900 shadow-xs">
                    📁 {uploadedFileName || 'ไฟล์นำเข้า'} ({allRecords.length} แถว)
                  </span>
                )}
              </div>
            </div>

            {/* Scope Year & Reference Info */}
            <div className="flex flex-wrap items-center gap-3">
              {/* Year Selector */}
              <div className="flex items-center gap-2">
                <span className="text-xs font-medium text-slate-500">ปีการศึกษา:</span>
                <select
                  value={selectedYear}
                  onChange={(e) => setSelectedYear(e.target.value)}
                  className="text-xs font-medium bg-slate-50 border border-slate-300 rounded-lg px-2.5 py-1.5 text-slate-800 focus:outline-blue-600"
                >
                  <option value="2570">ปี 2570 (Intake 2027) [เป้าหมายหลัก]</option>
                  <option value="all">ทุกปีการศึกษา (ทั้งหมด)</option>
                  <option value="2569">ปี 2569 (Intake 2026)</option>
                </select>
              </div>

              {/* Status indicator */}
              <div className="text-xs text-slate-500 flex items-center gap-1.5 bg-slate-50 px-3 py-1.5 rounded-lg border border-slate-200">
                <Clock className="w-3.5 h-3.5 text-blue-600" />
                <span>เกณฑ์อ้างอิง: 10 พ.ย. 2569</span>
              </div>
            </div>

          </div>
        </section>

        {/* Top 4 KPI Summary Cards */}
        <section className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
          {/* Card 1: Total Applications */}
          <div className="bg-white border border-slate-200 rounded-xl p-5 shadow-xs flex flex-col justify-between">
            <div className="flex items-center justify-between">
              <span className="text-xs font-medium text-slate-500">ใบสมัครทั้งหมด ({selectedYear === '2570' ? 'ปี 2570' : 'ในเงื่อนไข'})</span>
              <div className="w-8 h-8 rounded-lg bg-blue-50 text-blue-800 flex items-center justify-center">
                <FileSpreadsheet className="w-4 h-4" />
              </div>
            </div>
            <div className="mt-3">
              <div className="text-3xl font-bold tracking-tight text-slate-900 tabular-nums">
                {stats.totalApplications.toLocaleString()}
              </div>
              <p className="text-xs text-slate-500 mt-1">
                Admission Owners: <span className="font-semibold text-slate-700">{stats.byAdmissionOwner.filter(o => !o.isUnassigned).length} ท่าน</span>
              </p>
            </div>
          </div>

          {/* Card 2: Review Document Stage */}
          <div className="bg-white border border-amber-200 rounded-xl p-5 shadow-xs flex flex-col justify-between relative overflow-hidden">
            <div className="flex items-center justify-between">
              <span className="text-xs font-medium text-amber-900 font-semibold">
                รีวิวเอกสาร (Review Document)
              </span>
              <div className="w-8 h-8 rounded-lg bg-amber-100 text-amber-800 flex items-center justify-center">
                <FileCheck2 className="w-4 h-4" />
              </div>
            </div>
            <div className="mt-3">
              <div className="text-3xl font-bold tracking-tight text-amber-900 tabular-nums">
                {stats.reviewAging.totalReviewDocs.toLocaleString()}
              </div>
              <p className="text-xs text-amber-800 mt-1">
                คิดเป็น {((stats.reviewAging.totalReviewDocs / (stats.totalApplications || 1)) * 100).toFixed(1)}% ของทั้งหมด
              </p>
            </div>
          </div>

          {/* Card 3: Admission Owner Assignment KPI */}
          <div
            onClick={() => {
              if (selectedAdmissionOwner === 'Unassigned') {
                setSelectedAdmissionOwner('all');
              } else {
                setSelectedAdmissionOwner('Unassigned');
              }
            }}
            className={`bg-white border rounded-xl p-5 shadow-xs flex flex-col justify-between cursor-pointer transition-all ${
              selectedAdmissionOwner === 'Unassigned'
                ? 'border-amber-400 ring-2 ring-amber-300'
                : 'border-slate-200 hover:border-blue-300'
            }`}
            title="คลิกเพื่อกรองใบสมัครที่ยังไม่ได้ระบุ Admission Owner"
          >
            <div className="flex items-center justify-between">
              <span className="text-xs font-medium text-slate-500">การมอบหมาย Admission Owner</span>
              <div className="w-8 h-8 rounded-lg bg-indigo-50 text-indigo-800 flex items-center justify-center">
                <UserCheck className="w-4 h-4" />
              </div>
            </div>
            <div className="mt-3">
              <div className="text-3xl font-bold tracking-tight text-indigo-900 tabular-nums">
                {(stats.totalApplications - stats.unassignedAdmissionCount).toLocaleString()}
                <span className="text-xs font-normal text-slate-500 ml-1.5">
                  / {stats.totalApplications} ใบ
                </span>
              </div>
              <p className="text-xs text-slate-500 mt-1 flex items-center justify-between">
                <span>ยังไม่ระบุ (Unassigned):</span>
                <span className={`font-semibold tabular-nums px-1.5 py-0.2 rounded text-[11px] ${
                  stats.unassignedAdmissionCount > 0 ? 'bg-amber-100 text-amber-800' : 'bg-slate-100 text-slate-700'
                }`}>
                  {stats.unassignedAdmissionCount} ใบ ({((stats.unassignedAdmissionCount / (stats.totalApplications || 1)) * 100).toFixed(1)}%)
                </span>
              </p>
            </div>
          </div>

          {/* Card 4: Invoiced & Confirmed Payment */}
          <div className="bg-white border border-slate-200 rounded-xl p-5 shadow-xs flex flex-col justify-between">
            <div className="flex items-center justify-between">
              <span className="text-xs font-medium text-slate-500">ออกบิล & ยืนยันชำระเงิน</span>
              <div className="w-8 h-8 rounded-lg bg-emerald-50 text-emerald-800 flex items-center justify-center">
                <CheckCircle2 className="w-4 h-4" />
              </div>
            </div>
            <div className="mt-3">
              <div className="text-3xl font-bold tracking-tight text-emerald-800 tabular-nums">
                {(stats.confirmedCount + stats.invoicedCount).toLocaleString()}
              </div>
              <p className="text-xs text-slate-500 mt-1">
                จ่ายแล้ว <span className="font-semibold text-emerald-700 tabular-nums">{stats.confirmedCount}</span> · รอชำระ <span className="font-semibold text-purple-700 tabular-nums">{stats.invoicedCount}</span>
              </p>
            </div>
          </div>
        </section>

        {/* SPECIAL HIGHLIGHT SECTION: Review Document Aging Process */}
        <section className="bg-white border border-slate-200 rounded-xl p-6 shadow-xs space-y-4">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-slate-100 pb-3">
            <div>
              <h2 className="text-base font-bold text-slate-900 flex items-center gap-2">
                <AlertCircle className="w-5 h-5 text-amber-600" />
                วิเคราะห์ระยะเวลาใบสมัครในสเตจ "รีวิวเอกสาร" (Review Document Aging)
              </h2>
              <p className="text-xs text-slate-500 mt-0.5">
                ติดตามระยะเวลาที่ใบสมัครค้างอยู่ในขั้นตอนตรวจสอบเอกสาร เพื่อป้องกันเอกสารตกค้างเกินกำหนด
              </p>
            </div>
            <div className="text-xs font-medium text-slate-600">
              รวมทั้งหมดในสเตจนี้: <span className="font-bold text-slate-900 tabular-nums">{stats.reviewAging.totalReviewDocs} ใบ</span>
            </div>
          </div>

          {/* 3 Process Aging Cards */}
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
            
            {/* 1. Less than 1 week */}
            <button
              onClick={() => {
                setSelectedStage('Review Document');
                setSelectedAging(selectedAging === 'less_1w' ? 'all' : 'less_1w');
              }}
              className={`text-left p-4 rounded-xl border transition-all ${
                selectedAging === 'less_1w'
                  ? 'bg-emerald-50 border-emerald-400 ring-2 ring-emerald-400'
                  : 'bg-slate-50/80 border-slate-200 hover:bg-emerald-50/40 hover:border-emerald-200'
              }`}
            >
              <div className="flex items-center justify-between text-xs text-emerald-800 font-semibold mb-1">
                <span>ไม่เกิน 1 สัปดาห์ (≤ 7 วัน)</span>
                <span className="text-[11px] px-2 py-0.5 rounded-full bg-emerald-100">ปกติ</span>
              </div>
              <div className="text-2xl font-bold text-emerald-900 tabular-nums">
                {stats.reviewAging.lessThan1Week.count}{' '}
                <span className="text-xs font-normal text-emerald-700">ใบ</span>
              </div>
              <div className="mt-2 w-full bg-slate-200 rounded-full h-1.5 overflow-hidden">
                <div
                  className="bg-emerald-500 h-1.5 rounded-full"
                  style={{ width: `${stats.reviewAging.lessThan1Week.percentage}%` }}
                />
              </div>
              <p className="text-[11px] text-slate-500 mt-1.5 flex justify-between">
                <span>สัดส่วนในสเตจรีวิว:</span>
                <span className="font-semibold text-slate-700 tabular-nums">{stats.reviewAging.lessThan1Week.percentage.toFixed(1)}%</span>
              </p>
            </button>

            {/* 2. 2-4 weeks */}
            <button
              onClick={() => {
                setSelectedStage('Review Document');
                setSelectedAging(selectedAging === '2_to_4w' ? 'all' : '2_to_4w');
              }}
              className={`text-left p-4 rounded-xl border transition-all ${
                selectedAging === '2_to_4w'
                  ? 'bg-amber-50 border-amber-400 ring-2 ring-amber-400'
                  : 'bg-slate-50/80 border-slate-200 hover:bg-amber-50/40 hover:border-amber-200'
              }`}
            >
              <div className="flex items-center justify-between text-xs text-amber-800 font-semibold mb-1">
                <span>2 - 4 สัปดาห์ (8 - 28 วัน)</span>
                <span className="text-[11px] px-2 py-0.5 rounded-full bg-amber-100 text-amber-800">ควรเร่งรัด</span>
              </div>
              <div className="text-2xl font-bold text-amber-900 tabular-nums">
                {stats.reviewAging.twoToFourWeeks.count}{' '}
                <span className="text-xs font-normal text-amber-700">ใบ</span>
              </div>
              <div className="mt-2 w-full bg-slate-200 rounded-full h-1.5 overflow-hidden">
                <div
                  className="bg-amber-500 h-1.5 rounded-full"
                  style={{ width: `${stats.reviewAging.twoToFourWeeks.percentage}%` }}
                />
              </div>
              <p className="text-[11px] text-slate-500 mt-1.5 flex justify-between">
                <span>สัดส่วนในสเตจรีวิว:</span>
                <span className="font-semibold text-slate-700 tabular-nums">{stats.reviewAging.twoToFourWeeks.percentage.toFixed(1)}%</span>
              </p>
            </button>

            {/* 3. More than 4 weeks */}
            <button
              onClick={() => {
                setSelectedStage('Review Document');
                setSelectedAging(selectedAging === 'more_4w' ? 'all' : 'more_4w');
              }}
              className={`text-left p-4 rounded-xl border transition-all ${
                selectedAging === 'more_4w'
                  ? 'bg-rose-50 border-rose-400 ring-2 ring-rose-400'
                  : 'bg-slate-50/80 border-slate-200 hover:bg-rose-50/40 hover:border-rose-200'
              }`}
            >
              <div className="flex items-center justify-between text-xs text-rose-800 font-semibold mb-1">
                <span>มากกว่า 4 สัปดาห์ (&gt; 28 วัน)</span>
                <span className="text-[11px] px-2 py-0.5 rounded-full bg-rose-100 text-rose-800">ตกค้างเร่งด่วน</span>
              </div>
              <div className="text-2xl font-bold text-rose-900 tabular-nums">
                {stats.reviewAging.moreThan4Weeks.count}{' '}
                <span className="text-xs font-normal text-rose-700">ใบ</span>
              </div>
              <div className="mt-2 w-full bg-slate-200 rounded-full h-1.5 overflow-hidden">
                <div
                  className="bg-rose-500 h-1.5 rounded-full"
                  style={{ width: `${stats.reviewAging.moreThan4Weeks.percentage}%` }}
                />
              </div>
              <p className="text-[11px] text-slate-500 mt-1.5 flex justify-between">
                <span>สัดส่วนในสเตจรีวิว:</span>
                <span className="font-semibold text-slate-700 tabular-nums">{stats.reviewAging.moreThan4Weeks.percentage.toFixed(1)}%</span>
              </p>
            </button>

          </div>

          {selectedAging !== 'all' && (
            <div className="flex items-center justify-between bg-blue-50/70 border border-blue-200/60 rounded-lg px-4 py-2 text-xs text-blue-900">
              <span>กำลังกรองตารางข้อมูลเฉพาะกลุ่ม: <strong>{selectedAging === 'less_1w' ? 'ไม่เกิน 1 สัปดาห์' : selectedAging === '2_to_4w' ? '2-4 สัปดาห์' : 'มากกว่า 4 สัปดาห์'}</strong></span>
              <button
                onClick={() => setSelectedAging('all')}
                className="text-blue-700 underline font-semibold hover:text-blue-900"
              >
                ล้างตัวกรองนี้
              </button>
            </div>
          )}
        </section>

        {/* 3 CORE CHARTS SECTION */}
        <section className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          
          {/* Chart 1: Admission Owner Distribution (Bar Chart) */}
          <div className="bg-white border border-slate-200 rounded-xl p-5 shadow-xs lg:col-span-1 flex flex-col justify-between">
            <div>
              <div className="flex items-center justify-between mb-2">
                <h3 className="text-sm font-bold text-slate-900 flex items-center gap-2">
                  <UserCheck className="w-4 h-4 text-blue-800" />
                  (1) การกระจายของ {ownerViewMode === 'admission' ? 'Admission Owner' : 'Application Owner'}
                </h3>
              </div>
              
              {/* Toggle View Mode: Admission Owner vs Application Owner */}
              <div className="flex items-center gap-1 p-0.5 bg-slate-100 rounded-lg border border-slate-200/70 mb-3">
                <button
                  onClick={() => setOwnerViewMode('admission')}
                  className={`flex-1 py-1 px-2 text-[11px] font-semibold rounded-md transition-all ${
                    ownerViewMode === 'admission'
                      ? 'bg-white text-blue-950 shadow-xs'
                      : 'text-slate-500 hover:text-slate-800'
                  }`}
                >
                  Admission Owner
                </button>
                <button
                  onClick={() => setOwnerViewMode('application')}
                  className={`flex-1 py-1 px-2 text-[11px] font-semibold rounded-md transition-all ${
                    ownerViewMode === 'application'
                      ? 'bg-white text-blue-950 shadow-xs'
                      : 'text-slate-500 hover:text-slate-800'
                  }`}
                >
                  Application Owner
                </button>
              </div>

              <p className="text-xs text-slate-500 mb-3">
                {ownerViewMode === 'admission'
                  ? 'จำนวนใบสมัครที่เจ้าหน้าที่ฝ่ายรับสมัคร (Admission Team) แต่ละท่านดูแล'
                  : 'จำนวนใบสมัครตาม Application Owner (ผู้ประสานงาน/เอเจนต์/สถาบัน)'}
              </p>
            </div>

            {/* Horizontal Bar Chart for Owners */}
            <div className="space-y-2.5 max-h-[350px] overflow-y-auto pr-1">
              {ownerViewMode === 'admission' ? (
                // ADMISSION OWNER BARS
                stats.byAdmissionOwner.map((item, idx) => {
                  const isSelected = selectedAdmissionOwner === item.owner;
                  return (
                    <div
                      key={item.owner + '-' + idx}
                      onClick={() => {
                        setSelectedAdmissionOwner(isSelected ? 'all' : item.owner);
                      }}
                      className={`cursor-pointer p-2.5 rounded-lg transition-colors border ${
                        isSelected
                          ? 'bg-blue-50 border-blue-400 ring-1 ring-blue-300'
                          : item.isUnassigned
                          ? 'bg-amber-50/50 border-amber-200/70 hover:bg-amber-50'
                          : 'border-transparent hover:bg-slate-50'
                      }`}
                    >
                      <div className="flex items-center justify-between text-xs mb-1">
                        <span className="font-medium text-slate-800 truncate max-w-[190px]" title={item.displayName}>
                          {item.isUnassigned ? (
                            <span className="text-amber-800 font-semibold">⚠️ {item.displayName}</span>
                          ) : (
                            <span>{idx + 1}. {item.displayName}</span>
                          )}
                        </span>
                        <span className="tabular-nums font-semibold text-slate-900">
                          {item.count} <span className="text-slate-400 font-normal">({item.percentage.toFixed(1)}%)</span>
                        </span>
                      </div>
                      <div className="w-full bg-slate-100 rounded-full h-2 overflow-hidden">
                        <div
                          className={`h-2 rounded-full transition-all duration-300 ${
                            item.isUnassigned ? 'bg-amber-500' : 'bg-blue-900'
                          }`}
                          style={{ width: `${Math.max(item.percentage, 3)}%` }}
                        />
                      </div>
                    </div>
                  );
                })
              ) : (
                // APPLICATION OWNER BARS
                stats.byOwner.slice(0, 8).map((item, idx) => {
                  const isSelected = selectedApplicationOwner === item.owner;
                  return (
                    <div
                      key={item.owner}
                      onClick={() => setSelectedApplicationOwner(isSelected ? 'all' : item.owner)}
                      className={`cursor-pointer p-2 rounded-lg transition-colors border ${
                        isSelected ? 'bg-blue-50 border-blue-300' : 'border-transparent hover:bg-slate-50'
                      }`}
                    >
                      <div className="flex items-center justify-between text-xs mb-1">
                        <span className="font-medium text-slate-800 truncate max-w-[170px]" title={item.owner}>
                          {idx + 1}. {item.owner}
                        </span>
                        <span className="tabular-nums font-semibold text-slate-900">
                          {item.count} <span className="text-slate-400 font-normal">({item.percentage.toFixed(1)}%)</span>
                        </span>
                      </div>
                      <div className="w-full bg-slate-100 rounded-full h-2 overflow-hidden">
                        <div
                          className="bg-blue-800 h-2 rounded-full transition-all duration-300"
                          style={{ width: `${Math.max(item.percentage, 3)}%` }}
                        />
                      </div>
                    </div>
                  );
                })
              )}
            </div>

            {/* Clear owner filter button */}
            {selectedAdmissionOwner !== 'all' && ownerViewMode === 'admission' && (
              <div className="mt-3 pt-3 border-t border-slate-100 flex items-center justify-between text-xs">
                <span className="text-blue-900 font-medium">
                  กรองโดย: {getAdmissionOwnerDisplayName(selectedAdmissionOwner)}
                </span>
                <button
                  onClick={() => setSelectedAdmissionOwner('all')}
                  className="text-slate-500 hover:text-slate-800 underline"
                >
                  ล้างตัวกรอง
                </button>
              </div>
            )}
            {selectedApplicationOwner !== 'all' && ownerViewMode === 'application' && (
              <div className="mt-3 pt-3 border-t border-slate-100 flex items-center justify-between text-xs">
                <span className="text-blue-900 font-medium">กรองโดย: {selectedApplicationOwner}</span>
                <button
                  onClick={() => setSelectedApplicationOwner('all')}
                  className="text-slate-500 hover:text-slate-800 underline"
                >
                  ล้างตัวกรอง
                </button>
              </div>
            )}
          </div>

          {/* Chart 2: Stage Comparison Line / Area Trend */}
          <div className="bg-white border border-slate-200 rounded-xl p-5 shadow-xs lg:col-span-1 flex flex-col justify-between">
            <div>
              <div className="flex items-center justify-between mb-1">
                <h3 className="text-sm font-bold text-slate-900 flex items-center gap-2">
                  <Layers className="w-4 h-4 text-blue-800" />
                  (2) เปรียบเทียบจำนวนใบสมัครแต่ละสเตจ
                </h3>
              </div>
              <p className="text-xs text-slate-500 mb-4">
                สัดส่วนและจำนวนใบสมัครในแต่ละสเตจปัจจุบัน
              </p>
            </div>

            {/* Custom SVG Stage Comparison Curve & Bars */}
            <div className="h-[280px] flex flex-col justify-end">
              <div className="space-y-2.5">
                {stats.byStage.slice(0, 6).map((stageItem) => {
                  const isSelected = selectedStage === stageItem.stage;
                  const maxCount = stats.byStage[0]?.count || 1;
                  const widthPercent = (stageItem.count / maxCount) * 100;

                  return (
                    <div
                      key={stageItem.stage}
                      onClick={() => setSelectedStage(isSelected ? 'all' : stageItem.stage)}
                      className={`cursor-pointer p-2 rounded-lg transition-colors border ${
                        isSelected ? 'bg-blue-50 border-blue-300' : 'border-transparent hover:bg-slate-50'
                      }`}
                    >
                      <div className="flex items-center justify-between text-xs mb-1">
                        <span className="font-medium text-slate-800 truncate max-w-[170px]">
                          {stageItem.stage}
                        </span>
                        <span className="tabular-nums font-semibold text-slate-900">
                          {stageItem.count} <span className="text-slate-400 font-normal">({stageItem.percentage.toFixed(1)}%)</span>
                        </span>
                      </div>
                      <div className="w-full bg-slate-100 rounded-full h-2.5 overflow-hidden">
                        <div
                          className={`h-2.5 rounded-full transition-all duration-300 ${
                            stageItem.stage === 'Review Document'
                              ? 'bg-amber-500'
                              : stageItem.stage === 'Wait for Language Result'
                              ? 'bg-sky-600'
                              : stageItem.stage === 'Confirmed Payment'
                              ? 'bg-emerald-600'
                              : stageItem.stage === 'Invoiced'
                              ? 'bg-purple-600'
                              : stageItem.stage === 'Closed Lost'
                              ? 'bg-slate-400'
                              : 'bg-blue-700'
                          }`}
                          style={{ width: `${Math.max(widthPercent, 4)}%` }}
                        />
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>

            {selectedStage !== 'all' && (
              <div className="mt-3 pt-3 border-t border-slate-100 flex items-center justify-between text-xs">
                <span className="text-blue-900 font-medium">กรองโดย: {selectedStage}</span>
                <button
                  onClick={() => setSelectedStage('all')}
                  className="text-slate-500 hover:text-slate-800 underline"
                >
                  ล้างตัวกรอง
                </button>
              </div>
            )}
          </div>

          {/* Chart 3: Monthly Influx (Bar Chart) */}
          <div className="bg-white border border-slate-200 rounded-xl p-5 shadow-xs lg:col-span-1 flex flex-col justify-between">
            <div>
              <div className="flex items-center justify-between mb-1">
                <h3 className="text-sm font-bold text-slate-900 flex items-center gap-2">
                  <Calendar className="w-4 h-4 text-blue-800" />
                  (3) ใบสมัครที่เข้ามาในแต่ละเดือน
                </h3>
              </div>
              <p className="text-xs text-slate-500 mb-2">
                แนวโน้มปริมาณใบสมัครรายเดือนตาม Create Date
              </p>
            </div>

            {/* SVG Monthly Vertical Bar Chart */}
            <div className="h-[280px] flex items-end justify-between gap-2 pt-6 px-2 border-b border-slate-200">
              {(() => {
                const displayedMonths = stats.byMonth.slice(-6);
                const maxMonthCount = Math.max(...displayedMonths.map(x => x.count), 1);

                return displayedMonths.map((m) => {
                  const heightPercent = Math.max((m.count / maxMonthCount) * 100, 8);

                  return (
                    <div key={m.month} className="flex-1 flex flex-col items-center h-full justify-end group">
                      <span className="text-xs font-bold text-slate-800 tabular-nums mb-1 group-hover:text-blue-800">
                        {m.count}
                      </span>
                      <div className="w-full bg-slate-100 rounded-t-md h-[180px] flex items-end">
                        <div
                          className="w-full bg-blue-900 group-hover:bg-blue-700 rounded-t-md transition-all duration-300"
                          style={{ height: `${heightPercent}%` }}
                          title={`${m.label}: ${m.count} ใบสมัคร`}
                        />
                      </div>
                      <span className="text-[11px] text-slate-700 font-medium mt-2 text-center whitespace-nowrap">
                        {m.label.split(' ')[0]}
                      </span>
                      <span className="text-[10px] text-slate-400">
                        {m.label.split(' ')[1]}
                      </span>
                    </div>
                  );
                });
              })()}
            </div>

            <div className="mt-3 text-xs text-slate-500 text-center">
              {(() => {
                const displayedMonths = stats.byMonth.slice(-6);
                const displayedTotal = displayedMonths.reduce((sum, item) => sum + item.count, 0);
                const avg = displayedMonths.length > 0 ? Math.round(displayedTotal / displayedMonths.length) : 0;
                return (
                  <span>
                    เฉลี่ย 6 เดือนล่าสุด: <strong className="text-blue-900 font-bold tabular-nums">{avg.toLocaleString()}</strong> ใบ/เดือน
                    <span className="text-slate-400 ml-1.5">(รวม {displayedTotal.toLocaleString()} ใบ)</span>
                  </span>
                );
              })()}
            </div>
          </div>

        </section>

        {/* DATA TABLE SECTION: Filterable, Sortable, Anonymized */}
        <section className="bg-white border border-slate-200 rounded-xl shadow-xs overflow-hidden">
          {/* Table Header Controls */}
          <div className="p-4 sm:p-5 border-b border-slate-200 flex flex-col md:flex-row md:items-center justify-between gap-4">
            <div>
              <div className="flex items-center gap-2">
                <h3 className="font-bold text-slate-900 text-base">
                  ตารางรายงานข้อมูลใบสมัคร (Application Master Data)
                </h3>
                <span className="text-xs bg-slate-100 text-slate-700 px-2 py-0.5 rounded-full font-semibold tabular-nums">
                  {filteredRecords.length} รายการ
                </span>
              </div>
              <p className="text-xs text-slate-500 mt-0.5">
                ค้นหาและกรองข้อมูลตาม Admission Owner, สเตจ, ผู้รับผิดชอบ หรือระยะเวลาการรีวิว
              </p>
            </div>

            {/* Search and Action Bar */}
            <div className="flex flex-wrap items-center gap-2.5">
              {/* Search Input */}
              <div className="relative">
                <Search className="w-3.5 h-3.5 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                <input
                  type="text"
                  placeholder="ค้นหาชื่อ, รหัส, เจ้าหน้าที่..."
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  className="pl-8 pr-3 py-1.5 text-xs bg-slate-50 border border-slate-300 rounded-lg focus:outline-blue-600 w-44 sm:w-56"
                />
                {searchQuery && (
                  <button
                    onClick={() => setSearchQuery('')}
                    className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"
                  >
                    <X className="w-3.5 h-3.5" />
                  </button>
                )}
              </div>

              {/* Admission Owner Filter Dropdown */}
              <select
                value={selectedAdmissionOwner}
                onChange={(e) => setSelectedAdmissionOwner(e.target.value)}
                className="text-xs bg-slate-50 border border-slate-300 rounded-lg px-2.5 py-1.5 text-slate-800 focus:outline-blue-600 max-w-[160px] truncate font-medium"
              >
                <option value="all">ทุก Admission Owner</option>
                <option value="Unassigned">⚠️ ยังไม่ระบุ (Unassigned)</option>
                {availableAdmissionOwners.map(owner => (
                  <option key={owner} value={owner}>
                    {getAdmissionOwnerDisplayName(owner)}
                  </option>
                ))}
              </select>

              {/* Application Owner Filter Dropdown */}
              <select
                value={selectedApplicationOwner}
                onChange={(e) => setSelectedApplicationOwner(e.target.value)}
                className="text-xs bg-slate-50 border border-slate-300 rounded-lg px-2.5 py-1.5 text-slate-800 focus:outline-blue-600 max-w-[150px] truncate"
              >
                <option value="all">ทุก App Owner</option>
                {availableApplicationOwners.map(owner => (
                  <option key={owner} value={owner}>{owner}</option>
                ))}
              </select>

              {/* Stage Filter Dropdown */}
              <select
                value={selectedStage}
                onChange={(e) => setSelectedStage(e.target.value)}
                className="text-xs bg-slate-50 border border-slate-300 rounded-lg px-2.5 py-1.5 text-slate-800 focus:outline-blue-600 max-w-[140px] truncate"
              >
                <option value="all">ทุก Stage</option>
                {availableStages.map(stage => (
                  <option key={stage} value={stage}>{stage}</option>
                ))}
              </select>

              {/* Export CSV Button */}
              <button
                onClick={() => exportToCSV(filteredRecords, `bu-international-applications-${selectedYear}.csv`)}
                className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded-lg hover:bg-slate-50 transition-colors"
                title="ส่งออกรายการที่กรองไว้เป็นไฟล์ CSV"
              >
                <Download className="w-3.5 h-3.5 text-slate-600" />
                <span className="hidden sm:inline">ส่งออก CSV</span>
              </button>
            </div>
          </div>

          {/* Active Filter Chips */}
          {(selectedAdmissionOwner !== 'all' || selectedApplicationOwner !== 'all' || selectedStage !== 'all' || selectedAging !== 'all' || searchQuery) && (
            <div className="bg-slate-50 px-4 py-2 border-b border-slate-200 flex flex-wrap items-center gap-2 text-xs">
              <span className="text-slate-500 font-medium">ตัวกรองที่ใช้งาน:</span>
              {selectedAdmissionOwner !== 'all' && (
                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-sm bg-indigo-100 text-indigo-900">
                  Admission Owner: {getAdmissionOwnerDisplayName(selectedAdmissionOwner)}
                  <button onClick={() => setSelectedAdmissionOwner('all')} className="hover:text-indigo-950"><X className="w-3 h-3" /></button>
                </span>
              )}
              {selectedApplicationOwner !== 'all' && (
                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-sm bg-blue-100 text-blue-900">
                  App Owner: {selectedApplicationOwner}
                  <button onClick={() => setSelectedApplicationOwner('all')} className="hover:text-blue-950"><X className="w-3 h-3" /></button>
                </span>
              )}
              {selectedStage !== 'all' && (
                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-sm bg-blue-100 text-blue-900">
                  Stage: {selectedStage}
                  <button onClick={() => setSelectedStage('all')} className="hover:text-blue-950"><X className="w-3 h-3" /></button>
                </span>
              )}
              {selectedAging !== 'all' && (
                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-sm bg-amber-100 text-amber-900">
                  ระยะเวลารีวิว: {selectedAging === 'less_1w' ? '≤ 1 สัปดาห์' : selectedAging === '2_to_4w' ? '2-4 สัปดาห์' : '> 4 สัปดาห์'}
                  <button onClick={() => setSelectedAging('all')} className="hover:text-amber-950"><X className="w-3 h-3" /></button>
                </span>
              )}
              {searchQuery && (
                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-sm bg-slate-200 text-slate-800">
                  คำค้น: "{searchQuery}"
                  <button onClick={() => setSearchQuery('')} className="hover:text-slate-950"><X className="w-3 h-3" /></button>
                </span>
              )}
              <button
                onClick={() => {
                  setSelectedAdmissionOwner('all');
                  setSelectedApplicationOwner('all');
                  setSelectedStage('all');
                  setSelectedAging('all');
                  setSearchQuery('');
                }}
                className="text-blue-700 hover:text-blue-900 underline text-xs ml-auto"
              >
                ล้างตัวกรองทั้งหมด
              </button>
            </div>
          )}

          {/* Table Container */}
          <div className="overflow-x-auto">
            <table className="w-full text-left text-xs border-collapse">
              <thead>
                <tr className="bg-slate-50 text-slate-600 border-b border-slate-200">
                  <th className="py-3 px-4 font-semibold w-12 text-center">#</th>
                  <th className="py-3 px-4 font-semibold">Record ID</th>
                  <th className="py-3 px-4 font-semibold">
                    <span className="flex items-center gap-1">
                      ชื่อผู้สมัคร / อีเมล
                      {isPrivacyMasked && (
                        <span className="text-[10px] text-emerald-700 bg-emerald-50 px-1.5 py-0.2 rounded border border-emerald-200">
                          PDPA Masked
                        </span>
                      )}
                    </span>
                  </th>
                  <th
                    className="py-3 px-4 font-semibold cursor-pointer select-none hover:text-blue-900"
                    onClick={() => {
                      if (sortField === 'stage') setSortAsc(!sortAsc);
                      else { setSortField('stage'); setSortAsc(true); }
                    }}
                  >
                    <span className="flex items-center gap-1">
                      Application Stage
                      {sortField === 'stage' && (sortAsc ? ' ↑' : ' ↓')}
                    </span>
                  </th>
                  <th
                    className="py-3 px-4 font-semibold cursor-pointer select-none hover:text-blue-900 bg-blue-50/40"
                    onClick={() => {
                      if (sortField === 'admissionOwner') setSortAsc(!sortAsc);
                      else { setSortField('admissionOwner'); setSortAsc(true); }
                    }}
                  >
                    <span className="flex items-center gap-1 text-blue-950 font-bold">
                      Admission Owner
                      {sortField === 'admissionOwner' && (sortAsc ? ' ↑' : ' ↓')}
                    </span>
                  </th>
                  <th
                    className="py-3 px-4 font-semibold cursor-pointer select-none hover:text-blue-900"
                    onClick={() => {
                      if (sortField === 'owner') setSortAsc(!sortAsc);
                      else { setSortField('owner'); setSortAsc(true); }
                    }}
                  >
                    <span className="flex items-center gap-1">
                      Application Owner
                      {sortField === 'owner' && (sortAsc ? ' ↑' : ' ↓')}
                    </span>
                  </th>
                  <th
                    className="py-3 px-4 font-semibold cursor-pointer select-none hover:text-blue-900"
                    onClick={() => {
                      if (sortField === 'createDate') setSortAsc(!sortAsc);
                      else { setSortField('createDate'); setSortAsc(false); }
                    }}
                  >
                    <span className="flex items-center gap-1">
                      Create Date
                      {sortField === 'createDate' && (sortAsc ? ' ↑' : ' ↓')}
                    </span>
                  </th>
                  <th
                    className="py-3 px-4 font-semibold cursor-pointer select-none hover:text-blue-900 text-right"
                    onClick={() => {
                      if (sortField === 'agingDays') setSortAsc(!sortAsc);
                      else { setSortField('agingDays'); setSortAsc(false); }
                    }}
                  >
                    <span className="flex items-center justify-end gap-1">
                      ระยะเวลารีวิว (Aging)
                      {sortField === 'agingDays' && (sortAsc ? ' ↑' : ' ↓')}
                    </span>
                  </th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {paginatedRecords.length > 0 ? (
                  paginatedRecords.map((r, index) => {
                    const rowNumber = (currentPage - 1) * pageSize + index + 1;
                    const isReviewStage = r.stage === 'Review Document';
                    const hasAdmissionOwner = Boolean(r.admissionOwner && r.admissionOwner.trim() && r.admissionOwner.toLowerCase() !== 'unassigned');

                    return (
                      <tr key={r.id + '-' + index} className="hover:bg-blue-50/30 transition-colors">
                        <td className="py-3 px-4 text-center text-slate-400 tabular-nums">
                          {rowNumber}
                        </td>
                        <td className="py-3 px-4 font-mono text-slate-500 tabular-nums">
                          {r.id}
                        </td>
                        <td className="py-3 px-4">
                          <div className="font-medium text-slate-900">
                            {maskStudentName(r.name, isPrivacyMasked)}
                          </div>
                          {r.email && (
                            <div className="text-[11px] text-slate-400 font-mono">
                              {maskEmail(r.email, isPrivacyMasked)}
                            </div>
                          )}
                        </td>
                        <td className="py-3 px-4">
                          <span className={`inline-block px-2.5 py-1 text-[11px] font-medium rounded-md whitespace-nowrap ${getStageBadgeStyle(r.stage)}`}>
                            {r.stage}
                          </span>
                        </td>
                        {/* Admission Owner Column */}
                        <td className="py-3 px-4 bg-blue-50/20">
                          {hasAdmissionOwner ? (
                            <span className="font-semibold text-blue-950 font-mono text-[11px]">
                              {r.admissionOwner}
                            </span>
                          ) : (
                            <span className="inline-block px-2 py-0.5 rounded text-[10px] font-medium bg-amber-100/70 text-amber-800 border border-amber-200">
                              ยังไม่ระบุ
                            </span>
                          )}
                        </td>
                        {/* Application Owner Column */}
                        <td className="py-3 px-4 text-slate-700">
                          {r.owner || '-'}
                        </td>
                        <td className="py-3 px-4 text-slate-600 font-mono tabular-nums whitespace-nowrap">
                          {r.createDate || '-'}
                        </td>
                        <td className="py-3 px-4 text-right tabular-nums">
                          {isReviewStage && r.agingDays !== undefined ? (
                            <span
                              className={`inline-block px-2 py-0.5 rounded text-[11px] font-semibold ${
                                r.agingBucket === 'less_1w'
                                  ? 'text-emerald-700 bg-emerald-50 border border-emerald-200'
                                  : r.agingBucket === '2_to_4w'
                                  ? 'text-amber-700 bg-amber-50 border border-amber-200'
                                  : 'text-rose-700 bg-rose-50 border border-rose-200'
                              }`}
                            >
                              {r.agingDays} วัน ({r.agingBucket === 'less_1w' ? '≤ 1 สัปดาห์' : r.agingBucket === '2_to_4w' ? '2-4 สัปดาห์' : '> 4 สัปดาห์'})
                            </span>
                          ) : (
                            <span className="text-slate-400">-</span>
                          )}
                        </td>
                      </tr>
                    );
                  })
                ) : (
                  <tr>
                    <td colSpan={8} className="py-12 text-center text-slate-400">
                      <div className="flex flex-col items-center justify-center space-y-2">
                        <Filter className="w-8 h-8 text-slate-300" />
                        <span className="text-sm font-medium">ไม่พบใบสมัครที่ตรงกับเงื่อนไขตัวกรอง</span>
                        <button
                          onClick={() => {
                            setSelectedAdmissionOwner('all');
                            setSelectedApplicationOwner('all');
                            setSelectedStage('all');
                            setSelectedAging('all');
                            setSearchQuery('');
                          }}
                          className="text-xs text-blue-800 underline hover:text-blue-950 font-semibold"
                        >
                          ล้างตัวกรองทั้งหมด
                        </button>
                      </div>
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>

          {/* Table Pagination */}
          <div className="p-4 border-t border-slate-200 flex flex-col sm:flex-row items-center justify-between gap-3 text-xs text-slate-600 bg-slate-50/50">
            <div className="flex items-center gap-2">
              <span>แสดงแถวต่อหน้า:</span>
              <select
                value={pageSize}
                onChange={(e) => {
                  setPageSize(Number(e.target.value));
                  setCurrentPage(1);
                }}
                className="bg-white border border-slate-300 rounded px-2 py-1 text-slate-800 focus:outline-blue-600"
              >
                <option value={10}>10</option>
                <option value={25}>25</option>
                <option value={50}>50</option>
                <option value={100}>100</option>
              </select>
              <span className="text-slate-400">
                (รายการที่ {(currentPage - 1) * pageSize + 1} - {Math.min(currentPage * pageSize, filteredRecords.length)} จาก {filteredRecords.length})
              </span>
            </div>

            <div className="flex items-center gap-1.5">
              <button
                onClick={() => setCurrentPage(p => Math.max(p - 1, 1))}
                disabled={currentPage === 1}
                className="p-1.5 rounded border border-slate-300 bg-white hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed"
                title="หน้าก่อนหน้า"
              >
                <ChevronLeft className="w-4 h-4" />
              </button>
              <span className="px-2 font-medium tabular-nums">
                หน้า {currentPage} / {totalPages}
              </span>
              <button
                onClick={() => setCurrentPage(p => Math.min(p + 1, totalPages))}
                disabled={currentPage >= totalPages}
                className="p-1.5 rounded border border-slate-300 bg-white hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed"
                title="หน้าถัดไป"
              >
                <ChevronRight className="w-4 h-4" />
              </button>
            </div>
          </div>
        </section>

      </main>

      {/* Guide Modal: วิธีเปลี่ยนเป็นไฟล์จริง & การนำเข้าข้อมูล */}
      {showGuideModal && (
        <div className="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl max-w-2xl w-full p-6 shadow-2xl border border-slate-200 space-y-5 animate-in fade-in zoom-in-95 duration-200">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <div className="flex items-center gap-2">
                <FileText className="w-5 h-5 text-blue-900" />
                <h3 className="text-base font-bold text-slate-900">
                  คู่มือวิธีเปลี่ยนและนำเข้าไฟล์ข้อมูลจริง (Excel / CSV)
                </h3>
              </div>
              <button
                onClick={() => setShowGuideModal(false)}
                className="text-slate-400 hover:text-slate-600 p-1 rounded-lg"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="space-y-4 text-xs text-slate-600 leading-relaxed">
              <div className="p-3 bg-blue-50/70 border border-blue-200/60 rounded-xl">
                <p className="font-semibold text-blue-900 mb-1">
                  1. คอลัมน์ที่จำเป็นต้องมีในไฟล์ Excel หรือ CSV:
                </p>
                <ul className="list-disc pl-5 space-y-1 text-slate-700">
                  <li><strong>Admission owner:</strong> อีเมลหรือชื่อเจ้าหน้าที่ฝ่ายรับสมัคร เช่น <code>sarawut.p@bu.ac.th</code>, <code>krittin.c@bu.ac.th</code></li>
                  <li><strong>Application Stage:</strong> สถานะใบสมัคร เช่น <code>Review Document</code>, <code>Wait for Language Result</code>, <code>Invoiced</code>, <code>Confirmed Payment</code></li>
                  <li><strong>Create Date:</strong> วันที่ส่งใบสมัคร เช่น <code>2026-11-10 10:23</code> หรือ <code>YYYY-MM-DD</code></li>
                  <li><strong>Application owner:</strong> ชื่อเจ้าหน้าที่หรือทีมผู้ดูแลการสมัคร</li>
                  <li>(ถ้ามี) <strong>Application Name / Record ID / Application Year:</strong> เพื่อใช้ระบุใบสมัครและคัดกรองปี 2570</li>
                </ul>
              </div>

              <div>
                <p className="font-semibold text-slate-900 mb-1">
                  2. วิธีส่งออกไฟล์จาก HubSpot CRM:
                </p>
                <ol className="list-decimal pl-5 space-y-1 text-slate-700">
                  <li>ไปที่ระบบ HubSpot CRM เมนู <strong>CRM Objects / Applications</strong></li>
                  <li>คลิกปุ่ม <strong>Export</strong> ด้านขวาบน</li>
                  <li>เลือกรูปแบบไฟล์เป็น <strong>CSV</strong> หรือ <strong>Excel (.xlsx)</strong></li>
                  <li>กด Download ไฟล์ลงในคอมพิวเตอร์ของคุณ</li>
                </ol>
              </div>

              <div>
                <p className="font-semibold text-slate-900 mb-1">
                  3. วิธีนำเข้าไฟล์เข้าสู่ Dashboard นี้:
                </p>
                <ol className="list-decimal pl-5 space-y-1 text-slate-700">
                  <li>คลิกปุ่มสีน้ำเงินด้านขวาบน <strong>"อัปโหลด Excel / CSV"</strong></li>
                  <li>เลือกไฟล์ที่คุณบันทึกไว้ ระบบจะประมวลผลคำนวณกราฟและการ์ดสรุปให้ทันที</li>
                  <li>สามารถสลับดูระหว่าง <strong>ไฟล์จริง</strong>, <strong>ตัวอย่าง 50 แถว</strong>, และ <strong>ไฟล์ที่อัปโหลด</strong> ได้ตลอดเวลาจากแถบด้านบน</li>
                </ol>
              </div>

              <div className="p-3 bg-emerald-50 border border-emerald-200 rounded-xl flex items-start gap-2.5">
                <ShieldCheck className="w-5 h-5 text-emerald-700 shrink-0 mt-0.5" />
                <div>
                  <span className="font-bold text-emerald-900">มาตรฐานความปลอดภัย PDPA และ Client-Side Processing</span>
                  <p className="text-emerald-800 text-[11px] mt-0.5">
                    ไฟล์และข้อมูลทั้งหมดถูกประมวลผลภายในเบราว์เซอร์ของท่าน 100% (In-Browser Execution) ไม่มีการส่งข้อมูลนักศึกษาออกไปยังเซิร์ฟเวอร์ภายนอกใดๆ ทั้งสิ้น พร้อมโหมดปกป้องข้อมูลส่วนตัว (PDPA Masking) เพื่อความปลอดภัยสูงสุด
                  </p>
                </div>
              </div>
            </div>

            <div className="flex justify-end pt-2 border-t border-slate-100">
              <button
                onClick={() => setShowGuideModal(false)}
                className="px-4 py-2 text-xs font-semibold text-white bg-blue-900 hover:bg-blue-800 rounded-lg transition-colors"
              >
                เข้าใจแล้ว / ปิดหน้าต่าง
              </button>
            </div>
          </div>
        </div>
      )}

    </div>
  );
}
