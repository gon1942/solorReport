'use client';

import React, { useState, useEffect, useCallback } from 'react';
import { X, Upload, FileSpreadsheet, Brain, Settings2, Eye, Columns, Filter, CheckSquare, Square, Sparkles, User, Zap, Sun, Wind } from 'lucide-react';

interface PVSolarExcelPreprocessModalProps {
  isOpen: boolean;
  onClose: () => void;
  onProcess: (options: PVSolarExcelProcessOptions) => void;
}

interface PVSolarExcelProcessOptions {
  file: File;
  mode: 'ai' | 'manual';
  selectedSheet: string;
  selectedColumns: string[];
  dataDescription?: string;
  rowFilter?: {
    column: string;
    operator: 'equals' | 'contains' | 'greaterThan' | 'lessThan';
    value: string;
  };
  headerNormalization: boolean;
  preserveFormatting: boolean;
  pvSolarSpecificOptions?: {
    energyUnit: 'kW' | 'MW' | 'Wh' | 'MWh';
    normalizeDates: boolean;
    detectWeatherPatterns: boolean;
    normalizeLocations: boolean;
    includeMetadata: boolean;
    customMetadata?: {
      projectType: 'residential' | 'commercial' | 'utility' | 'hybrid';
      installationDate?: string;
      location?: string;
      capacity?: string;
      operator?: string;
    };
  };
  aiRecommendations?: {
    structure: string;
    columns: string[];
    filters: any[];
    pvSolarInsights?: {
      peakProductionTimes?: string[];
      weatherFactors?: string[];
      maintenancePatterns?: string[];
      performanceMetrics?: string[];
    };
  };
  dataSplitting?: {
    enabled: boolean;
    maxRowsPerFile: number;
    preserveHeaders: boolean;
  };
}

interface SheetData {
  name: string;
  headers: string[];
  preview: any[][];
  rowCount: number;
  columnCount: number;
}

export default function PVSolarExcelPreprocessModal({ isOpen, onClose, onProcess }: PVSolarExcelPreprocessModalProps) {
  const [processingStatus, setProcessingStatus] = useState<{
    isProcessing: boolean;
    step: string;
    progress: number;
  }>({
    isProcessing: false,
    step: '',
    progress: 0
  });
  const [file, setFile] = useState<File | null>(null);
  const [mode, setMode] = useState<'ai' | 'manual'>('manual');
  const [sheets, setSheets] = useState<SheetData[]>([]);
  const [selectedSheet, setSelectedSheet] = useState<string>('');
  const [selectedColumns, setSelectedColumns] = useState<string[]>([]);
  const [headerNormalization, setHeaderNormalization] = useState(true);
  const [preserveFormatting, setPreserveFormatting] = useState(true);
  const [isProcessing, setIsProcessing] = useState(false);
  const [aiRecommendations, setAiRecommendations] = useState<any>(null);
  const [previewData, setPreviewData] = useState<any[][]>([]);
  const [dataDescription, setDataDescription] = useState('');
  const [rowFilter, setRowFilter] = useState<any>({
    enabled: false,
    column: '',
    operator: 'contains',
    value: ''
  });
  const [dataSplitting, setDataSplitting] = useState({
    enabled: false,
    maxRowsPerFile: 100,
    preserveHeaders: true
  });

  // PV Solar specific options
  const [pvSolarOptions, setPvSolarOptions] = useState({
    energyUnit: 'kW' as 'kW' | 'MW' | 'Wh' | 'MWh',
    normalizeDates: true,
    detectWeatherPatterns: true,
    normalizeLocations: true,
    includeMetadata: true,
    customMetadata: {
      projectType: 'commercial' as 'residential' | 'commercial' | 'utility' | 'hybrid',
      installationDate: '',
      location: '',
      capacity: '',
      operator: ''
    }
  });

  const resetModalState = useCallback(() => {
    setFile(null);
    setMode('manual');
    setSheets([]);
    setSelectedSheet('');
    setSelectedColumns([]);
    setHeaderNormalization(true);
    setPreserveFormatting(true);
    setIsProcessing(false);
    setAiRecommendations(null);
    setPreviewData([]);
    setDataDescription('');
    setRowFilter({
      enabled: false,
      column: '',
      operator: 'contains',
      value: ''
    });
    setDataSplitting({
      enabled: false,
      maxRowsPerFile: 100,
      preserveHeaders: true
    });
    setPvSolarOptions({
      energyUnit: 'kW',
      normalizeDates: true,
      detectWeatherPatterns: true,
      normalizeLocations: true,
      includeMetadata: true,
      customMetadata: {
        projectType: 'commercial',
        installationDate: '',
        location: '',
        capacity: '',
        operator: ''
      }
    });
  }, []);

  useEffect(() => {
    if (!isOpen) {
      resetModalState();
    }
  }, [isOpen, resetModalState]);

  // PV Solar specific AI analysis
  const requestPVSolarAiRecommendations = async (file: File, sheets: SheetData[]) => {
    try {
      const currentSheet = sheets.find((s) => s.name === selectedSheet);
      if (!currentSheet) {
        console.error('❌ 선택된 시트를 찾을 수 없음:', selectedSheet);
        return;
      }

      console.log('🔆 PV Solar 데이터 분석 시작:', selectedSheet);
      
      const formData = new FormData();
      formData.append('file', file);
      formData.append(
        'sheets',
        JSON.stringify([
          {
            name: currentSheet.name,
            headers: currentSheet.headers,
            sampleData: currentSheet.preview.slice(0, 5),
            rowCount: currentSheet.rowCount,
            columnCount: currentSheet.columnCount,
            isPVSolarData: true // PV Solar 데이터임을 표시
          }
        ])
      );

      const response = await fetch('/api/preprocess/pv-solar/excel/recommend', {
        method: 'POST',
        body: formData
      });

      if (response.ok) {
        const recommendations = await response.json();
        setAiRecommendations(recommendations.data);
        
        if (recommendations.data.columns) {
          setSelectedColumns(recommendations.data.columns);
        }

        if (recommendations.data.aiGeneratedDescription) {
          setDataDescription(recommendations.data.aiGeneratedDescription);
        }

        // PV Solar specific insights 업데이트
        if (recommendations.data.pvSolarInsights) {
          console.log('🔆 PV Solar AI 인사이트:', recommendations.data.pvSolarInsights);
        }
      }
    } catch (error) {
      console.error('PV Solar AI 추천 요청 실패:', error);
    }
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const uploadedFile = e.target.files?.[0];
    if (!uploadedFile) return;

    setFile(uploadedFile);
    setIsProcessing(true);

    try {
      // PV Solar 특화 분석 API 호출
      const formData = new FormData();
      formData.append('file', uploadedFile);

      const response = await fetch('/api/preprocess/pv-solar/excel/analyze', {
        method: 'POST',
        body: formData
      });

      if (response.ok) {
        const result = await response.json();
        const sheetsData = result.data.sheets;

        setSheets(sheetsData);
        if (sheetsData.length > 0) {
          setSelectedSheet(sheetsData[0].name);
          setSelectedColumns(sheetsData[0].headers);
          setPreviewData(sheetsData[0].preview);
        }
      } else {
        throw new Error('파일 분석 실패');
      }
    } catch (error) {
      console.error('파일 읽기 오류:', error);
      alert('Excel 파일을 읽는 중 오류가 발생했습니다.');
    } finally {
      setIsProcessing(false);
    }
  };

  const handleModeChange = async (newMode: 'ai' | 'manual') => {
    setMode(newMode);

    if (newMode === 'ai' && file && sheets.length > 0) {
      setAiRecommendations(null);
      setDataDescription('');
      setIsProcessing(true);
      await requestPVSolarAiRecommendations(file, sheets);
      setIsProcessing(false);
    }

    if (newMode === 'manual') {
      setAiRecommendations(null);
      setDataDescription('');
      const currentSheet = sheets.find((s) => s.name === selectedSheet);
      if (currentSheet) {
        setSelectedColumns(currentSheet.headers);
      }
    }
  };

  const handleSheetChange = (sheetName: string) => {
    setSelectedSheet(sheetName);
    const sheet = sheets.find((s) => s.name === sheetName);
    if (sheet) {
      setSelectedColumns(sheet.headers);
      setPreviewData(sheet.preview);
    }
    setAiRecommendations(null);
    setDataDescription('');
  };

  const toggleColumn = (column: string) => {
    setSelectedColumns((prev) => (prev.includes(column) ? prev.filter((c) => c !== column) : [...prev, column]));
  };

  const handleProcess = async () => {
    if (!file) return;

    const options: PVSolarExcelProcessOptions = {
      file,
      mode,
      selectedSheet,
      selectedColumns,
      dataDescription,
      headerNormalization,
      preserveFormatting,
      rowFilter: rowFilter.enabled ? rowFilter : undefined,
      pvSolarSpecificOptions: pvSolarOptions,
      aiRecommendations,
      dataSplitting: dataSplitting.enabled ? dataSplitting : undefined
    };

    setProcessingStatus({
      isProcessing: true,
      step: '파일 분석 중...',
      progress: 10
    });

    try {
      setTimeout(() => setProcessingStatus((prev) => ({ ...prev, step: 'PV 데이터 전처리 중...', progress: 30 })), 1000);
      setTimeout(() => setProcessingStatus((prev) => ({ ...prev, step: '에너지 단위 정규화 중...', progress: 50 })), 2000);
      setTimeout(() => setProcessingStatus((prev) => ({ ...prev, step: '날씨 데이터 분석 중...', progress: 70 })), 3000);
      setTimeout(() => setProcessingStatus((prev) => ({ ...prev, step: '메타데이터 추가 중...', progress: 90 })), 4000);

      await onProcess(options);

      setTimeout(() => {
        setProcessingStatus({ isProcessing: false, step: '', progress: 0 });
        resetModalState();
        onClose();
      }, 5000);
    } catch (error) {
      console.error('전처리 실패:', error);
      setProcessingStatus({ isProcessing: false, step: '', progress: 0 });
      alert('전처리 중 오류가 발생했습니다.');
    }
  };

  if (!isOpen) return null;

  // PV Solar specific AI recommendations display
  const renderPVSolarAIRecommendations = () => {
    if (!aiRecommendations || mode !== 'ai') return null;

    return (
      <div className="mb-8">
        <div className="card p-4" style={{ backgroundColor: 'var(--secondary-500)', color: 'white' }}>
          <div className="flex items-center gap-2 mb-3">
            <Zap className="w-5 h-5" />
            <h3 className="font-semibold text-white">PV Solar AI 분석 결과</h3>
          </div>
          
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div>
              <p className="text-white/90 mb-2">{aiRecommendations.description}</p>
              
              {aiRecommendations.pvSolarInsights && (
                <div className="mt-3 space-y-2">
                  <div className="bg-white/10 p-2 rounded">
                    <span className="text-sm font-medium">⏰ 생산 피크 시간:</span>
                    <div className="text-xs mt-1">
                      {aiRecommendations.pvSolarInsights.peakProductionTimes?.join(', ')}
                    </div>
                  </div>
                  
                  <div className="bg-white/10 p-2 rounded">
                    <span className="text-sm font-medium">🌤️ 날씨 요인:</span>
                    <div className="text-xs mt-1">
                      {aiRecommendations.pvSolarInsights.weatherFactors?.join(', ')}
                    </div>
                  </div>
                  
                  <div className="bg-white/10 p-2 rounded">
                    <span className="text-sm font-medium">🔧 유지보수 패턴:</span>
                    <div className="text-xs mt-1">
                      {aiRecommendations.pvSolarInsights.maintenancePatterns?.join(', ')}
                    </div>
                  </div>
                </div>
              )}
            </div>
            
            <div>
              <h4 className="font-medium mb-2">📊 성능 메트릭</h4>
              {aiRecommendations.pvSolarInsights?.performanceMetrics?.map((metric: string, index: number) => (
                <div key={index} className="bg-white/10 p-2 rounded mb-2 text-sm">
                  {metric}
                </div>
              ))}
              
              {aiRecommendations.confidence && (
                <div className="mt-3 p-2 bg-white/10 rounded">
                  <span className="text-sm">AI 분석 신뢰도: </span>
                  <span className="font-bold">{Math.round((aiRecommendations.confidence || 0) * 100)}%</span>
                </div>
              )}
            </div>
          </div>
        </div>
      </div>
    );
  };

  if (processingStatus.isProcessing) {
    return (
      <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
        <div className="rounded-lg shadow-xl w-full max-w-md p-8" style={{ backgroundColor: 'var(--sidebar-bg)', borderColor: 'var(--card-border)' }}>
          <div className="text-center">
            <div className="mb-6 flex justify-center">
              <div className="relative">
                <div className="w-16 h-16 border-4 rounded-full animate-pulse" style={{ borderColor: 'var(--secondary-200)' }}></div>
                <div className="absolute top-0 left-0 w-16 h-16 border-4 rounded-full border-t-transparent animate-spin" style={{ borderColor: 'var(--secondary-600)' }}></div>
              </div>
            </div>
            <h3 className="font-semibold mb-2" style={{ color: 'var(--text-primary)' }}>
              PV Solar 파일 전처리 중
            </h3>
            <p className="text-sm mb-6" style={{ color: 'var(--text-secondary)' }}>
              {processingStatus.step || '처리 중입니다...'}
            </p>
            <div className="w-full rounded-full h-2 mb-4" style={{ backgroundColor: 'var(--neutral-200)' }}>
              <div
                className="h-2 rounded-full transition-all duration-300"
                style={{
                  width: `${processingStatus.progress}%`,
                  backgroundColor: 'var(--secondary-600)'
                }}
              ></div>
            </div>
            <p className="text-sm" style={{ color: 'var(--text-muted)' }}>
              {processingStatus.progress}% 완료
            </p>
            <div className="mt-6 p-3 rounded-lg" style={{ backgroundColor: 'var(--info-bg)', borderColor: 'var(--info-border)' }}>
              <p className="text-xs" style={{ color: 'var(--text-secondary)' }}>
                🔆 PV Solar 데이터는 특화된 전처리 알고리즘을 적용합니다.
              </p>
            </div>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
      <div className="rounded-lg shadow-xl w-full max-w-8xl max-h-[95vh] overflow-hidden" style={{ backgroundColor: 'var(--sidebar-bg)', borderColor: 'var(--card-border)' }}>
        {/* 헤더 */}
        <div className="flex items-center justify-between p-6 border-b" style={{ backgroundColor: 'var(--sidebar-bg)', borderColor: 'var(--card-border)' }}>
          <div className="flex items-center gap-3">
            <Sun className="w-6 h-6 text-yellow-600 dark:text-yellow-400" />
            <Zap className="w-6 h-6 text-orange-600 dark:text-orange-400" />
            <h2 className="font-bold" style={{ color: 'var(--text-primary)' }}>
              PV Solar 데이터 전처리
            </h2>
          </div>
          <button onClick={onClose} className="p-2 rounded-lg transition-colors" style={{ backgroundColor: 'transparent', color: 'var(--text-secondary)' }}>
            <X className="w-6 h-6" />
          </button>
        </div>

        <div className="p-6 overflow-y-auto max-h-[calc(95vh-180px)]" style={{ backgroundColor: 'var(--sidebar-bg)' }}>
          {/* 1단계: 파일 선택 및 시트 선택 */}
          <div className="mb-8">
            <div className="flex items-center gap-2 mb-4">
              <div className="w-8 h-8 bg-yellow-600 text-white rounded-full flex items-center justify-center text-sm font-bold">1</div>
              <h3 className="font-semibold" style={{ color: 'var(--text-primary)' }}>
                파일 및 시트 선택
              </h3>
            </div>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <div className="card p-4" style={{ backgroundColor: 'var(--card-bg)' }}>
                <div className="flex items-center gap-2 mb-3">
                  <Upload className="w-5 h-5 text-yellow-600" />
                  <label className="text-base font-medium" style={{ color: 'var(--text-primary)' }}>
                    PV Solar Excel 파일 선택
                  </label>
                </div>
                <div className="relative">
                  <input 
                    type="file" 
                    accept=".xlsx,.xls,.csv" 
                    onChange={handleFileUpload} 
                    className="block w-full text-base file:mr-4 file:py-2 file:px-4 file:rounded-full file:border-0 file:text-base file:font-semibold file:bg-yellow-50 file:text-yellow-700 hover:file:bg-yellow-100 dark:file:bg-yellow-900/20 dark:file:text-yellow-300" 
                    style={{ color: 'var(--text-secondary)' }} 
                  />
                  {file && (
                    <div className="mt-2 flex items-center gap-2 text-sm" style={{ color: 'var(--text-secondary)' }}>
                      <FileSpreadsheet className="w-4 h-4 text-green-500" />
                      <span>{file.name}</span>
                    </div>
                  )}
                </div>
              </div>

              <div className="card p-4" style={{ backgroundColor: 'var(--card-bg)' }}>
                <div className="flex items-center gap-2 mb-3">
                  <Columns className="w-5 h-5 text-yellow-600" />
                  <label className="text-base font-medium" style={{ color: 'var(--text-primary)' }}>
                    시트 선택
                  </label>
                </div>
                {sheets.length > 0 ? (
                  <select value={selectedSheet} onChange={(e) => handleSheetChange(e.target.value)} className="w-full px-3 py-2 border rounded-lg input" style={{ borderColor: 'var(--card-border)', backgroundColor: 'var(--card-bg)', color: 'var(--text-primary)' }}>
                    {sheets.map((sheet) => (
                      <option key={sheet.name} value={sheet.name}>
                        {sheet.name} ({sheet.rowCount}행 × {sheet.columnCount}열)
                      </option>
                    ))}
                  </select>
                ) : (
                  <div className="text-center text-sm py-4" style={{ color: 'var(--text-muted)' }}>
                    Excel 파일을 먼저 선택해주세요
                  </div>
                )}
              </div>
            </div>
          </div>

          {/* 2단계: 처리 모드 선택 */}
          {sheets.length > 0 && (
            <div className="mb-8">
              <div className="flex items-center gap-2 mb-4">
                <div className="w-8 h-8 bg-yellow-600 text-white rounded-full flex items-center justify-center text-sm font-bold">2</div>
                <h3 className="font-semibold" style={{ color: 'var(--text-primary)' }}>
                  처리 모드 선택
                </h3>
              </div>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <label className="card p-4 cursor-pointer transition-all duration-200" style={{ backgroundColor: mode === 'manual' ? 'var(--primary-100)' : 'var(--card-bg)', borderColor: mode === 'manual' ? 'var(--primary-500)' : 'var(--card-border)', borderWidth: mode === 'manual' ? '2px' : '1px' }}>
                  <div className="flex items-center gap-3">
                    <input type="radio" value="manual" checked={mode === 'manual'} onChange={(e) => handleModeChange(e.target.value as 'manual' | 'ai')} className="sr-only" />
                    <div className="w-6 h-6 rounded-full border-2 flex items-center justify-center" style={{ borderColor: mode === 'manual' ? 'var(--primary-500)' : 'var(--neutral-300)', backgroundColor: mode === 'manual' ? 'var(--primary-500)' : 'transparent' }}>
                      {mode === 'manual' && <div className="w-2 h-2 bg-white rounded-full" />}
                    </div>
                    <User className="w-5 h-5" style={{ color: mode === 'manual' ? 'var(--primary-600)' : 'var(--primary-600)' }} />
                    <div>
                      <div className="text-base font-medium" style={{ color: mode === 'manual' ? 'var(--primary-900)' : 'var(--text-primary)' }}>
                        수동 설정
                      </div>
                      <div className="text-sm" style={{ color: mode === 'manual' ? 'var(--primary-700)' : 'var(--text-secondary)' }}>
                        직접 열과 옵션을 선택합니다
                      </div>
                    </div>
                  </div>
                </label>
                <label className="card p-4 cursor-pointer transition-all duration-200" style={{ backgroundColor: mode === 'ai' ? 'var(--secondary-100)' : 'var(--card-bg)', borderColor: mode === 'ai' ? 'var(--secondary-500)' : 'var(--card-border)', borderWidth: mode === 'ai' ? '2px' : '1px' }}>
                  <div className="flex items-center gap-3">
                    <input type="radio" value="ai" checked={mode === 'ai'} onChange={(e) => handleModeChange(e.target.value as 'manual' | 'ai')} className="sr-only" />
                    <div className="w-6 h-6 rounded-full border-2 flex items-center justify-center" style={{ borderColor: mode === 'ai' ? 'var(--secondary-500)' : 'var(--neutral-300)', backgroundColor: mode === 'ai' ? 'var(--secondary-500)' : 'transparent' }}>
                      {mode === 'ai' && <div className="w-2 h-2 bg-white rounded-full" />}
                    </div>
                    <Sparkles className="w-5 h-5" style={{ color: mode === 'ai' ? 'var(--secondary-600)' : 'var(--secondary-600)' }} />
                    <div>
                      <div className="text-base font-medium" style={{ color: mode === 'ai' ? 'var(--secondary-900)' : 'var(--text-primary)' }}>
                        AI 자동 추천
                      </div>
                      <div className="text-sm" style={{ color: mode === 'ai' ? 'var(--secondary-700)' : 'var(--text-secondary)' }}>
                        PV Solar 데이터 전문 AI가 분석합니다
                      </div>
                    </div>
                  </div>
                </label>
              </div>
            </div>
          )}

          {/* PV Solar AI 추천 결과 표시 */}
          {renderPVSolarAIRecommendations()}

          {/* 3단계: PV Solar 특화 옵션 */}
          {selectedSheet && (
            <div className="mb-8">
              <div className="flex items-center gap-2 mb-4">
                <div className="w-8 h-8 bg-yellow-600 text-white rounded-full flex items-center justify-center text-sm font-bold">3</div>
                <h3 className="font-semibold" style={{ color: 'var(--text-primary)' }}>
                  PV Solar 설정
                </h3>
              </div>
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                <div className="space-y-4">
                  <div className="card p-4" style={{ backgroundColor: 'var(--card-bg)' }}>
                    <div className="flex items-center gap-2 mb-3">
                      <Zap className="w-5 h-5 text-yellow-600" />
                      <h3 className="font-medium" style={{ color: 'var(--text-primary)' }}>
                        에너지 단위 설정
                      </h3>
                    </div>
                    <select 
                      value={pvSolarOptions.energyUnit} 
                      onChange={(e) => setPvSolarOptions({...pvSolarOptions, energyUnit: e.target.value as any})}
                      className="w-full px-3 py-2 border rounded-lg input" 
                      style={{ borderColor: 'var(--card-border)', backgroundColor: 'var(--card-bg)', color: 'var(--text-primary)' }}
                    >
                      <option value="kW">킬로와트 (kW)</option>
                      <option value="MW">메가와트 (MW)</option>
                      <option value="Wh">와트시 (Wh)</option>
                      <option value="MWh">메가와트시 (MWh)</option>
                    </select>
                  </div>

                  <div className="card p-4" style={{ backgroundColor: 'var(--card-bg)' }}>
                    <div className="flex items-center gap-2 mb-3">
                      <Wind className="w-5 h-5 text-blue-600" />
                      <h3 className="font-medium" style={{ color: 'var(--text-primary)' }}>
                        데이터 정규화 옵션
                      </h3>
                    </div>
                    <div className="space-y-3">
                      <label className="flex items-center">
                        <input type="checkbox" checked={pvSolarOptions.normalizeDates} onChange={(e) => setPvSolarOptions({...pvSolarOptions, normalizeDates: e.target.checked})} className="mr-2" />
                        <span className="text-sm" style={{ color: 'var(--text-primary)' }}>
                          날짜 형식 정규화
                        </span>
                      </label>
                      <label className="flex items-center">
                        <input type="checkbox" checked={pvSolarOptions.detectWeatherPatterns} onChange={(e) => setPvSolarOptions({...pvSolarOptions, detectWeatherPatterns: e.target.checked})} className="mr-2" />
                        <span className="text-sm" style={{ color: 'var(--text-primary)' }}>
                          날씨 데이터 패턴 감지
                        </span>
                      </label>
                      <label className="flex items-center">
                        <input type="checkbox" checked={pvSolarOptions.normalizeLocations} onChange={(e) => setPvSolarOptions({...pvSolarOptions, normalizeLocations: e.target.checked})} className="mr-2" />
                        <span className="text-sm" style={{ color: 'var(--text-primary)' }}>
                          위치 정보 정규화
                        </span>
                      </label>
                    </div>
                  </div>
                </div>

                <div className="space-y-4">
                  <div className="card p-4" style={{ backgroundColor: 'var(--card-bg)' }}>
                    <div className="flex items-center gap-2 mb-3">
                      <Sun className="w-5 h-5 text-orange-600" />
                      <h3 className="font-medium" style={{ color: 'var(--text-primary)' }}>
                        프로젝트 정보
                      </h3>
                    </div>
                    <div className="space-y-3">
                      <select 
                        value={pvSolarOptions.customMetadata.projectType} 
                        onChange={(e) => setPvSolarOptions({
                          ...pvSolarOptions, 
                          customMetadata: {...pvSolarOptions.customMetadata, projectType: e.target.value as any}
                        })}
                        className="w-full px-3 py-2 border rounded-lg input" 
                        style={{ borderColor: 'var(--card-border)', backgroundColor: 'var(--card-bg)', color: 'var(--text-primary)' }}
                      >
                        <option value="residential">주거용</option>
                        <option value="commercial">상업용</option>
                        <option value="utility">공공용</option>
                        <option value="hybrid">복합용</option>
                      </select>
                      
                      <input 
                        type="text" 
                        placeholder="설치 일자" 
                        value={pvSolarOptions.customMetadata.installationDate}
                        onChange={(e) => setPvSolarOptions({
                          ...pvSolarOptions, 
                          customMetadata: {...pvSolarOptions.customMetadata, installationDate: e.target.value}
                        })}
                        className="w-full px-3 py-2 border rounded-lg input" 
                        style={{ borderColor: 'var(--card-border)', backgroundColor: 'var(--card-bg)', color: 'var(--text-primary)' }}
                      />
                      
                      <input 
                        type="text" 
                        placeholder="위치" 
                        value={pvSolarOptions.customMetadata.location}
                        onChange={(e) => setPvSolarOptions({
                          ...pvSolarOptions, 
                          customMetadata: {...pvSolarOptions.customMetadata, location: e.target.value}
                        })}
                        className="w-full px-3 py-2 border rounded-lg input" 
                        style={{ borderColor: 'var(--card-border)', backgroundColor: 'var(--card-bg)', color: 'var(--text-primary)' }}
                      />
                      
                      <input 
                        type="text" 
                        placeholder="용량 (kW)" 
                        value={pvSolarOptions.customMetadata.capacity}
                        onChange={(e) => setPvSolarOptions({
                          ...pvSolarOptions, 
                          customMetadata: {...pvSolarOptions.customMetadata, capacity: e.target.value}
                        })}
                        className="w-full px-3 py-2 border rounded-lg input" 
                        style={{ borderColor: 'var(--card-border)', backgroundColor: 'var(--card-bg)', color: 'var(--text-primary)' }}
                      />
                    </div>
                  </div>

                  <div className="card p-4" style={{ backgroundColor: 'var(--card-bg)' }}>
                    <label className="flex items-center">
                      <input type="checkbox" checked={pvSolarOptions.includeMetadata} onChange={(e) => setPvSolarOptions({...pvSolarOptions, includeMetadata: e.target.checked})} className="mr-2" />
                      <span className="text-sm" style={{ color: 'var(--text-primary)' }}>
                        PV Solar 메타데이터 문서에 포함
                      </span>
                    </label>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* 나머지 기능은 기본 Excel 전처리 모달과 동일 */}
          {/* 4단계: 열 선택 및 전처리 옵션 */}
          {/* 5단계: 미리보기 */}
          {/* ... (중략) ... */}
        </div>

        <div className="flex justify-end gap-3 p-6 border-t" style={{ backgroundColor: 'var(--sidebar-bg)', borderColor: 'var(--card-border)' }}>
          <button onClick={onClose} className="btn btn-secondary px-6 py-2" style={{ backgroundColor: 'var(--btn-outline-bg)', color: 'var(--btn-outline-fg)', border: '1px solid var(--btn-outline-border)' }}>
            취소
          </button>
          <button 
            onClick={handleProcess} 
            disabled={!file || selectedColumns.length === 0 || isProcessing}
            className="btn btn-primary px-6 py-2 disabled:opacity-50 disabled:cursor-not-allowed" 
            style={{ backgroundColor: 'var(--btn-primary-bg)', color: 'var(--btn-primary-fg)' }}
          >
            {isProcessing ? '처리 중...' : 'PV Solar 처리 시작'}
          </button>
        </div>
      </div>
    </div>
  );
}