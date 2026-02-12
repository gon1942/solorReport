"use client";

import React, { useEffect, useMemo, useRef, useState, type ChangeEvent, type DragEvent } from 'react';
import { FileSpreadsheet, FileText, Upload as UploadIcon } from 'lucide-react';
import capexGanttData from '@/data/capex-gantt.json';
import PVSolarExcelPreprocessModal from '@/components/PVSolarExcelPreprocessModal';
import styles from './page.module.css';

type UploadStatus = {
  status: 'idle' | 'processing' | 'success' | 'error';
  message: string;
  details?: any;
};

type PreviewChunk = { order: string; page: string; content: string };

type GanttTask = {
  id: string;
  name: string;
  type: string;
  parentId?: string;
  startWeekIndex: number;
  endWeekIndex: number;
};

type GanttPhase = 'phase1' | 'phase2';

type GanttPayload = Record<GanttPhase, GanttTask[]>;

type RawCapexBlock = {
  sn?: number;
  category: string;
  rows: Array<{
    id: string;
    name: string;
    startWeekIndex: number;
    endWeekIndex: number;
  }>;
};

const defaultPreview: PreviewChunk[] = [
  {
    order: '1',
    page: '-',
    content: '미리볼 수 있는 데이터가 없습니다. 업로드 또는 최근 파일을 선택하세요.'
  }
];

type RecentFile = {
  name: string;
  updatedAt?: string;
  size?: string;
  status?: 'active' | 'idle' | 'processing';
};

export default function PVSolarUploadPage() {
  const [now, setNow] = useState(() => new Date());
  const [isMounted, setIsMounted] = useState(false);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [uploadStatus, setUploadStatus] = useState<UploadStatus>({
    status: 'idle',
    message: '파일 업로드를 시작하세요.'
  });
  const [recentFiles, setRecentFiles] = useState<RecentFile[]>([]);
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [selectedSheet, setSelectedSheet] = useState<string>('');
  const [reportFormat, setReportFormat] = useState<'ppt' | 'pdf'>('ppt');
  const [isDragging, setIsDragging] = useState(false);
  const [pollTarget, setPollTarget] = useState<{ fileName: string; userId: string } | null>(null);
  const [isLoadingRecent, setIsLoadingRecent] = useState(false);
  const [isGeneratingReport, setIsGeneratingReport] = useState(false);
  const [deletingFile, setDeletingFile] = useState<string | null>(null);
  const [selectedRecentFile, setSelectedRecentFile] = useState<string | null>(null);
  const [previewData, setPreviewData] = useState<PreviewChunk[]>(defaultPreview);
  const [previewLoading, setPreviewLoading] = useState(false);
  const [previewError, setPreviewError] = useState<string | null>(null);
  const [isPreviewCollapsed, setIsPreviewCollapsed] = useState(true);
  const [isGanttCollapsed, setIsGanttCollapsed] = useState(true);
  const [processingContext, setProcessingContext] = useState<'upload' | 'report'>('upload');
  const [reportCategory, setReportCategory] = useState('');
  const [ganttPhase, setGanttPhase] = useState<GanttPhase>('phase1');
  const [ganttTasksByPhase, setGanttTasksByPhase] = useState<Record<GanttPhase, GanttTask[]>>({
    phase1: [],
    phase2: []
  });
  const [ganttSaveMessage, setGanttSaveMessage] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const saveCapexGanttToServer = async (payload: GanttPayload): Promise<boolean> => {
    try {
      const response = await fetch('/api/capex/gantt', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      return response.ok;
    } catch (error) {
      return false;
    }
  };

  const loadCapexGanttFromServer = async (): Promise<GanttPayload | null> => {
    try {
      const response = await fetch('/api/capex/gantt', { method: 'GET', cache: 'no-store' });
      if (!response.ok) return null;
      const payload = (await response.json()) as { phase1?: GanttTask[]; phase2?: GanttTask[] };
      if (!payload || (!Array.isArray(payload.phase1) && !Array.isArray(payload.phase2))) {
        return null;
      }
      return {
        phase1: Array.isArray(payload.phase1) ? payload.phase1 : [],
        phase2: Array.isArray(payload.phase2) ? payload.phase2 : []
      };
    } catch (error) {
      return null;
    }
  };

  const capexTemplateTasks = useMemo(() => {
    const data = capexGanttData as RawCapexBlock[];
    const tasks: GanttTask[] = [];

    data.forEach((block, idx) => {
      const groupId = `group-${block.sn ?? idx + 1}`;
      const childRows = block.rows || [];
      const hasChildren = childRows.length > 0;
      const groupStart = hasChildren ? Math.min(...childRows.map((r) => r.startWeekIndex)) : 0;
      const groupEnd = hasChildren ? Math.max(...childRows.map((r) => r.endWeekIndex)) : 0;

      tasks.push({
        id: groupId,
        name: block.category,
        type: 'group',
        startWeekIndex: groupStart,
        endWeekIndex: groupEnd
      });

      childRows.forEach((row) => {
        tasks.push({
          id: row.id,
          name: row.name,
          type: 'task',
          parentId: groupId,
          startWeekIndex: row.startWeekIndex,
          endWeekIndex: row.endWeekIndex
        });
      });
    });

    return tasks;
  }, []);
  const activeGanttTasks = useMemo(
    () => ganttTasksByPhase[ganttPhase] || [],
    [ganttPhase, ganttTasksByPhase]
  );
  const capexMaxWeek = useMemo(
    () => activeGanttTasks.reduce((max, task) => Math.max(max, task.endWeekIndex), 0),
    [activeGanttTasks]
  );
  const ganttBaseDate = useMemo(() => new Date(), []);
  const ganttMonthSegments = useMemo(() => {
    const totalWeeks = capexMaxWeek + 1;
    const bucketSize = 4;
    const segmentCount = Math.ceil(totalWeeks / bucketSize);
    const segments: Array<{ label: string; weeks: number }> = [];

    for (let i = 0; i < segmentCount; i += 1) {
      const d = new Date(ganttBaseDate.getFullYear(), ganttBaseDate.getMonth() + i, 1);
      const label = d.toLocaleDateString('en-US', { month: 'short', year: 'numeric' });
      const coveredWeeks = Math.min(bucketSize, totalWeeks - (i * bucketSize));
      segments.push({ label, weeks: coveredWeeks });
    }

    return segments;
  }, [capexMaxWeek, ganttBaseDate]);
  const ganttColWidth = 72;
  const hasInvalidGantt = useMemo(() => {
    return Object.values(ganttTasksByPhase).some((tasks) =>
      tasks.some((t) => t.type === 'task' && t.endWeekIndex < t.startWeekIndex)
    );
  }, [ganttTasksByPhase]);

  const formattedNow = useMemo(() => {
    return new Intl.DateTimeFormat('ko-KR', {
      dateStyle: 'medium',
      timeStyle: 'medium'
    }).format(now);
  }, [now]);

  const ganttRenderRows = useMemo(() => {
    const childrenByParent = activeGanttTasks.reduce<Record<string, GanttTask[]>>((acc, task) => {
      if (task.parentId) {
        acc[task.parentId] = acc[task.parentId] || [];
        acc[task.parentId].push(task);
      }
      return acc;
    }, {});

    const computedGroups = activeGanttTasks
      .filter((t) => t.type === 'group')
      .map((group) => {
        const children = childrenByParent[group.id] || [];
        if (!children.length) return group;
        const minStart = Math.min(...children.map((c) => c.startWeekIndex));
        const maxEnd = Math.max(...children.map((c) => c.endWeekIndex));
        return { ...group, startWeekIndex: minStart, endWeekIndex: maxEnd } as GanttTask;
      });

    const rows: Array<GanttTask & { isGroup?: boolean }> = [];
    computedGroups.forEach((group) => {
      rows.push({ ...group, isGroup: true });
      const children = (childrenByParent[group.id] || []).sort(
        (a, b) => a.startWeekIndex - b.startWeekIndex || a.name.localeCompare(b.name)
      );
      children.forEach((child) => rows.push(child));
    });

    const orphanTasks = activeGanttTasks.filter((t) => t.type !== 'group' && !t.parentId);
    return rows.concat(orphanTasks);
  }, [activeGanttTasks]);

  const statusApiKey = process.env.NEXT_PUBLIC_RAG_STATUS_KEY || '';
  const statusApiBase = process.env.NEXT_PUBLIC_RAG_STATUS_BASE || '';
  const recentFilesApi = process.env.NEXT_PUBLIC_RAG_FILES_API || '';
  const deleteApi = process.env.NEXT_PUBLIC_RAG_DELETE_API || '';
  const chunksApi = process.env.NEXT_PUBLIC_RAG_CHUNKS_API || '';
  const defaultUserId = process.env.NEXT_PUBLIC_RAG_DEFAULT_USER_ID || '';

  const parseMarkdownTable = (text: string) => {
    const lines = text.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
    if (!lines.length) return null;

    // Find the first block of pipe-delimited lines
    let start = -1;
    for (let i = 0; i < lines.length; i++) {
      if (lines[i].includes('|')) {
        start = i;
        break;
      }
    }
    if (start === -1) return null;

    const tableLines: string[] = [];
    for (let i = start; i < lines.length; i++) {
      const line = lines[i];
      if (!line.includes('|')) break;
      tableLines.push(line);
    }

    if (tableLines.length < 2) return null;

    const clean = (row: string) => row.replace(/^\|/, '').replace(/\|$/, '');
    const split = (row: string) => clean(row).split('|').map((c) => c.trim());

    const header = split(tableLines[0]);
    let body = tableLines.slice(1);
    if (body[0] && /---/.test(body[0])) {
      body = body.slice(1);
    }
    const rows = body.map((r) => split(r)).filter((r) => r.length > 0);
    if (!rows.length) return null;

    // Normalize column counts
    const colCount = Math.max(header.length, ...rows.map((r) => r.length));
    const normalizeRow = (r: string[]) => {
      const copy = [...r];
      while (copy.length < colCount) copy.push('');
      return copy.slice(0, colCount);
    };

    return {
      headers: normalizeRow(header),
      rows: rows.map(normalizeRow)
    };
  };

  const handleProcess = async (options: any) => {
    if (!options?.file) {
      setUploadStatus({ status: 'error', message: '파일이 선택되지 않았습니다.' });
      return;
    }

    setProcessingContext('upload');

    setUploadStatus({
      status: 'processing',
      message: 'PV Solar 데이터 전처리 시작...',
      details: { name: options.file.name }
    });

    try {
      const formData = new FormData();
      formData.append('file', options.file);
      formData.append('selected_sheet', options.selectedSheet || selectedSheet || options.file.name);
      formData.append('rag_pdf_backend', 'pdfplumber');
      formData.append('rag_generate_image_embeddings', 'no');
      formData.append('use_enhanced_embedding', 'no');
      formData.append('rag_process_text_only', 'no');
      formData.append('rag_mode', 'normal'); //'intelligence');
      formData.append('overwrite', 'true');
      formData.append('user_id', defaultUserId);
      formData.append('rag_enable_graphrag', 'no');

      formData.append(
        'pv_solar_specific_options',
        JSON.stringify(options.pvSolarSpecificOptions || {})
      );

      const response = await fetch('/api/rag/upload', {
        method: 'POST',
        body: formData,
      });

      const result = await response.json();
      const apiData = result?.data?.data || result?.data || {};
      const statusFromApi = (apiData.status as string | undefined)?.toLowerCase();
      const messageFromApi = apiData.message as string | undefined;
      const embeddingStatus = (apiData.metadata?.embedding_status as string | undefined)?.toLowerCase();
      const embeddedFlag = apiData.embedded === true || apiData.metadata?.embedded === true;
      const chunkCount = apiData.metadata?.chunk_count ?? apiData.chunk_count ?? 0;

      const completedKeywords = ['completed', 'ready', 'success', 'done', 'finished'];
      const processingKeywords = ['processing', 'process', 'embed', 'embedding', 'chunk', '임베딩', '처리 중'];
      const messageLower = (messageFromApi || '').toLowerCase();

      const isCompletedStatus = (statusFromApi && completedKeywords.some((k) => statusFromApi.includes(k)))
        || (embeddingStatus && completedKeywords.some((k) => embeddingStatus.includes(k)))
        || embeddedFlag
        || chunkCount > 0;

      const isProcessingStatus = (statusFromApi && processingKeywords.some((k) => statusFromApi.includes(k)))
        || (embeddingStatus && processingKeywords.some((k) => embeddingStatus.includes(k)));

      const messageHintsProcessing = processingKeywords.some((k) => messageLower.includes(k));

      const shouldTreatAsProcessing = (!isCompletedStatus && (isProcessingStatus || messageHintsProcessing)) || (!isCompletedStatus && (result?.success || apiData?.success));
      const normalizedStatus: UploadStatus['status'] = shouldTreatAsProcessing ? 'processing' : isCompletedStatus ? 'success' : 'error';

      setUploadStatus({
        status: normalizedStatus,
        message: normalizedStatus === 'processing'
          ? (messageFromApi || '파일 업로드됨, 임베딩 진행 중...')
          : (messageFromApi || 'PV Solar 데이터가 성공적으로 처리되었습니다.'),
        details: { ...apiData, status: apiData.status || (shouldTreatAsProcessing ? 'processing' : 'success'), name: options.file.name }
      });

      if (normalizedStatus === 'processing') {
        setPollTarget({ fileName: options.file.name, userId: defaultUserId });
      } else if (normalizedStatus === 'success') {
        setPollTarget(null);
        setSelectedRecentFile(options.file.name);
        setSelectedSheet(options.file.name);
        await fetchRecentFiles();
        await loadPreviewChunks(options.file.name);
      }
    } catch (error) {
      setUploadStatus({
        status: 'error',
        message: 'PV Solar 데이터 처리 중 네트워크 오류가 발생했습니다.',
        details: error
      });
      setPollTarget(null);
    }
  };

  const loadPreviewChunks = async (fileName: string) => {
    if (!fileName) return;
    try {
      setPreviewLoading(true);
      setPreviewError(null);
      const url = `${chunksApi}?filename=${encodeURIComponent(fileName)}&user_id=${encodeURIComponent(defaultUserId)}`;
      const res = await fetch(url, {
        method: 'GET',
        headers: {
          accept: 'application/json',
          'X-API-Key': statusApiKey
        }
      });

      if (!res.ok) throw new Error(`청크 조회 실패: ${res.status}`);

      const json = await res.json();
      const chunks = (json?.chunks || json?.data?.chunks || json?.data || []) as any[];

      if (!Array.isArray(chunks) || chunks.length === 0) {
        setPreviewData([{ order: '-', page: '-', content: '미리볼 수 있는 데이터가 없습니다.' }]);
        return;
      }

      const rows: PreviewChunk[] = chunks.slice(0, 10).map((chunk, idx) => {
        const content = typeof chunk === 'string'
          ? chunk
          : chunk?.content || chunk?.text || chunk?.chunk || JSON.stringify(chunk);
        const page = chunk?.metadata?.page_number ?? chunk?.page_number ?? '-';
        const order = chunk?.metadata?.chunk_index ?? chunk?.index ?? idx + 1;
        return { order: String(order), page: String(page), content };
      });

      setPreviewData(rows);
    } catch (error) {
      setPreviewError('청크 미리보기를 불러오지 못했습니다.');
      setPreviewData(defaultPreview);
    } finally {
      setPreviewLoading(false);
    }
  };

  const startUpload = (file: File) => {
    setProcessingContext('upload');
    setSelectedFile(file);
    setSelectedRecentFile(file.name);
    setSelectedSheet(file.name);
    setUploadStatus({ status: 'processing', message: '파일 업로드 중...', details: { name: file.name } });
    handleProcess({
      file,
      selectedSheet: file.name,
      pvSolarSpecificOptions: {
        energyUnit: 'MW',
        normalizeDates: true,
        detectWeatherPatterns: true,
        normalizeLocations: true,
        includeMetadata: true
      }
    });
  };

  const onFileChange = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      startUpload(file);
    }
  };

  const onDrop = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files?.[0];
    if (file) {
      startUpload(file);
    }
  };

  const statusClassName = (() => {
    if (uploadStatus.status === 'success') return `${styles.statusCard} ${styles.statusSuccess}`;
    if (uploadStatus.status === 'error') return `${styles.statusCard} ${styles.statusError}`;
    if (uploadStatus.status === 'processing') return styles.statusCard;
    return `${styles.statusCard} ${styles.statusIdle}`;
  })();

  const statusChip = (() => {
    const label = uploadStatus.details?.status || (uploadStatus.status === 'success' ? '완료' : uploadStatus.status === 'error' ? '오류' : '진행 중');
    const time = uploadStatus.details?.timestamp || uploadStatus.details?.updatedAt;
    const baseClass = uploadStatus.status === 'success'
      ? `${styles.statusChip} ${styles.statusChipSuccess}`
      : uploadStatus.status === 'error'
        ? `${styles.statusChip} ${styles.statusChipError}`
        : `${styles.statusChip} ${styles.statusChipProcessing}`;

    return {
      label,
      time,
      className: baseClass
    };
  })();

  const targetFile = selectedSheet || selectedRecentFile;
  const totalGanttTaskCount = ganttTasksByPhase.phase1.length + ganttTasksByPhase.phase2.length;
  const isGenerateDisabled =
    isGeneratingReport ||
    previewLoading ||
    !targetFile ||
    (reportCategory === 'capex' && hasInvalidGantt);
  const generatedFiles = uploadStatus.details?.generatedFiles;
  const reportPath = generatedFiles?.report as string | undefined;
  const reportDownloadUrl = reportPath ? `/api/report/download?file=${encodeURIComponent(reportPath)}` : '';
  const generationLogs = Array.isArray(uploadStatus.details?.logs)
    ? (uploadStatus.details.logs as Array<{ step?: string; stdout?: string; stderr?: string }>)
    : [];
  const previewSummary = useMemo(() => {
    if (previewLoading) {
      return {
        title: '미리보기 준비 중',
        details: ['데이터를 불러오는 중입니다.']
      };
    }
    if (previewError) {
      return {
        title: '미리보기 불러오기 실패',
        details: [previewError]
      };
    }
    const chunkCount = previewData.length;
    const tableChunkCount = previewData.filter((chunk) => Boolean(parseMarkdownTable(chunk.content))).length;
    const textChunkCount = Math.max(0, chunkCount - tableChunkCount);
    const pages = previewData
      .map((chunk) => chunk.page)
      .filter((page) => page && page !== '-');
    const firstPage = pages[0] || '-';
    return {
      title: `총 ${chunkCount}개 청크`,
      details: [
        `표 형식 ${tableChunkCount}개`,
        `텍스트 형식 ${textChunkCount}개`,
        `첫 페이지 ${firstPage}`
      ]
    };
  }, [previewData, previewError, previewLoading]);

  useEffect(() => {
    if (!pollTarget || uploadStatus.status !== 'processing') return;

    let cancelled = false;

    const checkStatus = async () => {
      try {
        const url = `${statusApiBase}?fileName=${encodeURIComponent(pollTarget.fileName)}&userId=${encodeURIComponent(pollTarget.userId)}`;
        const res = await fetch(url, {
          method: 'GET',
          headers: {
            accept: 'application/json',
            'X-API-Key': statusApiKey
          }
        });

        if (!res.ok) throw new Error(`Status check failed: ${res.status}`);

        const json = await res.json();
        const apiData = json?.data?.data || json?.data || {};
        console.log("apiData=================="+JSON.stringify(apiData));

        const rawStatusFromApi = (
          apiData.status ||
          apiData.embedding_status ||
          apiData.embeddingStatus ||
          apiData?.metadata?.embedding_status ||
          apiData?.metadata?.embeddingStatus
        ) as string | undefined;
        const statusFromApi = rawStatusFromApi?.toLowerCase();
        const messageFromApi = (apiData.message || apiData?.metadata?.message) as string | undefined;

        if (cancelled) return;

        const processingStatuses = ['processing', 'pending', 'queued', 'running', 'init', 'embedding'];

        if (statusFromApi && !processingStatuses.includes(statusFromApi)) {
          const isSuccess = ['completed', 'ready', 'success'].includes(statusFromApi);
          setUploadStatus({
            status: isSuccess ? 'success' : 'error',
            message: messageFromApi || (isSuccess ? '임베딩이 완료되었습니다.' : '임베딩 처리 중 오류가 발생했습니다.'),
            details: { ...apiData, name: pollTarget.fileName }
          });
          setPollTarget(null);
          if (isSuccess) {
            setSelectedRecentFile(pollTarget.fileName);
            await fetchRecentFiles();
            await loadPreviewChunks(pollTarget.fileName);
          }
        } else {
          setUploadStatus((prev) => ({
            ...prev,
            status: 'processing',
            message: messageFromApi || prev.message,
            details: { ...apiData, name: pollTarget.fileName }
          }));
        }
      } catch (err) {
        if (cancelled) return;
        setUploadStatus({
          status: 'error',
          message: '임베딩 상태 조회 실패',
          details: err
        });
        setPollTarget(null);
      }
    };

    checkStatus();
    const intervalId = setInterval(checkStatus, 5000);

    return () => {
      cancelled = true;
      clearInterval(intervalId);
    };
  }, [pollTarget, statusApiBase, statusApiKey, uploadStatus.status]);

  const fetchRecentFiles = async () => {
    try {
      setIsLoadingRecent(true);
      const url = `${recentFilesApi}?userId=${encodeURIComponent(defaultUserId)}`;
      console.log("Recent Files URL=================="+url);
      const res = await fetch(url, {
        method: 'GET',
        headers: {
          accept: 'application/json',
          'X-API-Key': statusApiKey
        }
      });

      if (!res.ok) throw new Error(`Recent files fetch failed: ${res.status}`);

      const json = await res.json();
      const files = (json?.data?.files || json?.data || []) as any[];
      const mapped: RecentFile[] = files.map((f) => ({
        name: f?.file || f?.name || '알 수 없는 파일',
        updatedAt: f?.updatedAt || f?.timestamp || '',
        size: f?.size || '',
        status: f?.status === 'active' ? 'active' : f?.status === 'processing' ? 'processing' : 'idle'
      }));
      setRecentFiles(mapped);
    } catch (error) {
      setRecentFiles([]);
    } finally {
      setIsLoadingRecent(false);
    }
  };

  const deleteRecentFile = async (fileName: string) => {
    if (!fileName) return;
    try {
      setDeletingFile(fileName);
      const res = await fetch(deleteApi, {
        method: 'POST',
        headers: {
          accept: 'application/json',
          'Content-Type': 'application/json',
          'X-API-Key': statusApiKey
        },
        body: JSON.stringify({ file_paths: fileName, user_id: defaultUserId })
      });

      if (!res.ok) throw new Error(`삭제 실패: ${res.status}`);

      await fetchRecentFiles();
      setSelectedRecentFile((prev) => (prev === fileName ? null : prev));
      setSelectedSheet((prev) => (prev === fileName ? '' : prev));
      if (selectedRecentFile === fileName || selectedSheet === fileName) {
        setPreviewData(defaultPreview);
        setPreviewError(null);
      }
    } catch (error) {
      // 상태 카드로 오류만 표시
      setUploadStatus({
        status: 'error',
        message: '파일 삭제 중 오류가 발생했습니다.',
        details: { name: fileName, error: (error as Error)?.message }
      });
    } finally {
      setDeletingFile(null);
    }
  };

  useEffect(() => {
    fetchRecentFiles();
  }, [recentFilesApi, statusApiKey]);

  useEffect(() => {
    const timer = setInterval(() => {
      setNow(new Date());
    }, 1000);
    return () => clearInterval(timer);
  }, []);

  useEffect(() => {
    setIsMounted(true);
  }, []);

  useEffect(() => {
    if (!ganttSaveMessage) return;
    const timer = setTimeout(() => setGanttSaveMessage(null), 2000);
    return () => clearTimeout(timer);
  }, [ganttSaveMessage]);



  useEffect(() => {
    let cancelled = false;

    const loadCapexState = async () => {
      if (reportCategory !== 'capex') {
        setGanttPhase('phase1');
        setGanttTasksByPhase({ phase1: [], phase2: [] });
        return;
      }

      const clonedTasks = capexTemplateTasks.map((t) => ({ ...t }));
      setGanttPhase('phase1');

      const stored = await loadCapexGanttFromServer();
      if (cancelled) return;

      if (stored?.phase1?.length || stored?.phase2?.length) {
        setGanttTasksByPhase({
          phase1: stored?.phase1?.length ? stored.phase1 : clonedTasks.map((t) => ({ ...t })),
          phase2: stored?.phase2?.length ? stored.phase2 : clonedTasks.map((t) => ({ ...t }))
        });
      } else {
        setGanttTasksByPhase({
          phase1: clonedTasks.map((t) => ({ ...t })),
          phase2: clonedTasks.map((t) => ({ ...t }))
        });
      }
    };

    loadCapexState();
    return () => {
      cancelled = true;
    };
  }, [reportCategory]);

  return (
    <div className={styles.page}>
      <main className={styles.layout}>
        <section className={styles.leftPane}>
          <div className={styles.timeCard}>
            <div className={styles.timeLabel}>현재 시간</div>
            <div className={styles.timeValue} aria-live="polite">
              {isMounted ? formattedNow : '--'}
            </div>
          </div>
          <div className={styles.stepRow}>
            <span className={`${styles.stepBadge} ${styles.stepActive}`}>단계 1</span>
            <div>
              <div className={styles.stepTitle}>파일 업로드</div>
              <div className={styles.subTitle}>.xlsx 또는 .csv</div>
            </div>
          </div>

          <div
            className={`${styles.dropZone} ${isDragging ? styles.dropZoneActive : ''}`}
            onDragOver={(e) => {
              e.preventDefault();
              setIsDragging(true);
            }}
            onDragLeave={() => setIsDragging(false)}
            onDrop={onDrop}
          >
            <div className={styles.dropIcon}>
              <UploadIcon size={28} />
            </div>
            <div className={styles.dropTitle}>파일을 놓아주세요</div>
            <p className={styles.dropHint}>.xlsx 또는 .csv</p>
            {selectedFile && <p className={styles.dropHint}>선택된 파일: {selectedFile.name}</p>}
            <button className={styles.browseButton} onClick={() => fileInputRef.current?.click()}>
              찾아보기
            </button>
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls,.csv"
              className={styles.fileInput}
              onChange={onFileChange}
            />
          </div>

      {uploadStatus.message && (
        <div className={statusClassName}>
          <div className={styles.statusRow}>
            {uploadStatus.status === 'processing' && <span className={styles.spinner} aria-label="처리 중" />}
            <strong>{uploadStatus.details?.name || '업로드 상태'}</strong>
          </div>
          <span>{uploadStatus.message}</span>
          {uploadStatus.status === 'processing' && (
            <span className={styles.statusNote}>
              {processingContext === 'report'
                ? '리포트 생성 진행 중입니다. 완료되면 생성 결과가 자동으로 표시됩니다.'
                : '임베딩 진행 중입니다. 완료 시 자동으로 갱신됩니다.'}
            </span>
          )}
          <div className={styles.statusMeta}>
            <span className={statusChip.className}>{statusChip.label}</span>
            {statusChip.time && <span>{statusChip.time}</span>}
          </div>
        </div>
      )}

          <div className={styles.recentHeader}>최근 파일</div>
          <div className={styles.recentList}>
            {isLoadingRecent && <div className={styles.recentEmpty}>불러오는 중...</div>}
            {!isLoadingRecent && recentFiles.length === 0 && (
              <div className={styles.recentEmpty}>최근 업로드한 파일이 없습니다.</div>
            )}
            {!isLoadingRecent && recentFiles.map((file) => (
              <div
                key={file.name}
                className={`${styles.recentItem} ${file.status === 'active' ? styles.recentItemActive : ''} ${selectedRecentFile === file.name ? styles.recentItemSelected : ''}`}
                onClick={() => {
                  setSelectedRecentFile(file.name);
                  setSelectedSheet(file.name);
                  loadPreviewChunks(file.name);
                }}
              >
                <div className={styles.recentFileIcon}>
                  {file.status === 'active' ? <FileSpreadsheet size={20} /> : <FileText size={20} />}
                </div>
                <div className={styles.recentFileInfo}>
                  <div className={styles.recentFileName}>{file.name}</div>
                  <div className={styles.recentFileMeta}>
                    {file.updatedAt ? file.updatedAt : '최근 업데이트 정보 없음'}
                    {file.size ? ` • ${file.size}` : ''}
                  </div>
                  {file.status === 'active' && (
                    <span className={styles.statusPill}>Active</span>
                  )}
                  <div className={styles.recentActions}>
                    <button
                      className={styles.deleteButton}
                      onClick={(e) => {
                        e.stopPropagation();
                        deleteRecentFile(file.name);
                      }}
                      disabled={deletingFile === file.name}
                    >
                      {deletingFile === file.name ? '삭제 중...' : '삭제'}
                    </button>
                  </div>
                </div>
              </div>
            ))}
          </div>
        </section>

        <section className={styles.rightPane}>
          <div className={styles.stepCard}>
            <div className={styles.stepRow}>
              <span className={`${styles.stepBadge} ${styles.stepSecondary}`}>단계 2</span>
              <div>
                <div className={styles.stepTitle}>시트 및 프로젝트 선택</div>
                <div className={styles.subTitle}>타깃 시트와 생성 조건을 확인하세요.</div>
              </div>
            </div>

            <div>
              <div className={styles.sectionLabel}>대상 스프레드시트</div>
              <select
                className={styles.selectBox}
                value={selectedSheet}
                onChange={(e) => {
                  const value = e.target.value;
                  setSelectedSheet(value);
                  setSelectedRecentFile(value || null);
                  if (value) {
                    loadPreviewChunks(value);
                  }
                }}
              >
                <option value="" disabled>
                  대상 스프레드시트를 선택하세요
                </option>
                {recentFiles.length === 0 && <option value="" disabled>최근 업로드된 스프레드시트가 없습니다</option>}
                {recentFiles.map((file) => (
                  <option key={file.name} value={file.name}>
                    {file.name}
                  </option>
                ))}
              </select>
            </div>

            <div>
              <div className={styles.sectionLabel}>생성 체크리스트</div>
              <div className={styles.readinessCard}>
                <div className={styles.readinessRow}>
                  <span className={targetFile ? styles.readinessOk : styles.readinessWarn}>{targetFile ? '완료' : '필수'}</span>
                  <span>대상 파일 선택</span>
                  <span className={styles.readinessMeta}>{targetFile || '선택 필요'}</span>
                </div>
                <div className={styles.readinessRow}>
                  <span className={reportCategory ? styles.readinessOk : styles.readinessWarn}>{reportCategory ? '완료' : '필수'}</span>
                  <span>카테고리 선택</span>
                  <span className={styles.readinessMeta}>{reportCategory || '선택 필요'}</span>
                </div>
                <div className={styles.readinessRow}>
                  <span className={previewLoading ? styles.readinessWarn : styles.readinessOk}>{previewLoading ? '대기' : '완료'}</span>
                  <span>미리보기 로딩 상태</span>
                  <span className={styles.readinessMeta}>{previewLoading ? '로딩 중' : '준비됨'}</span>
                </div>
                {reportCategory === 'capex' && (
                  <>
                    <div className={styles.readinessRow}>
                      <span className={totalGanttTaskCount > 0 ? styles.readinessOk : styles.readinessWarn}>
                        {totalGanttTaskCount > 0 ? '완료' : '확인'}
                      </span>
                      <span>간트 차트 데이터 구성</span>
                      <span className={styles.readinessMeta}>{!hasInvalidGantt ? `${totalGanttTaskCount}개 작업` : '시작/종료 주차 확인 필요'}</span>
                    </div>
                  </>
                )}
                {isGenerateDisabled && (
                  <div className={styles.readinessHint}>조건이 충족되면 리포트 생성 버튼이 자동으로 활성화됩니다.</div>
                )}
              </div>
            </div>
          </div>

          <div className={styles.stepCard}>
            <div className={styles.previewHeader}>
              <div>
                <div className={styles.sectionLabel}>데이터 미리보기</div>
              </div>
              <button
                type="button"
                className={styles.previewToggleButton}
                onClick={() => setIsPreviewCollapsed((prev) => !prev)}
              >
                {isPreviewCollapsed ? '펼치기' : '접기'}
              </button>
            </div>
            {isPreviewCollapsed ? (
              <div className={styles.previewCollapsedHint}>
                <div className={styles.previewSummaryTitle}>{previewSummary.title}</div>
                <div className={styles.previewSummaryList}>
                  {previewSummary.details.map((line, idx) => (
                    <span key={`preview-summary-${idx}`} className={styles.previewSummaryItem}>{line}</span>
                  ))}
                </div>
              </div>
            ) : (
              <>
                {previewLoading && <div className={styles.recentEmpty}>미리보기를 불러오는 중...</div>}
                {previewError && <div className={styles.recentEmpty}>{previewError}</div>}
                <div className={styles.chunkList}>
                  {previewData.map((chunk) => {
                    const parsed = parseMarkdownTable(chunk.content);
                    const isNumeric = (val: string) => {
                      const v = val.replace(/,/g, '');
                      return /^-?\d+(\.\d+)?$/.test(v) || /^\d{4}-?\d{2}-?\d{2}/.test(v);
                    };

                    return (
                      <div key={`${chunk.order}-${chunk.page}`} className={styles.chunkCard}>
                        <div className={styles.chunkMeta}>
                          <span className={styles.chunkTag}>청크 #{chunk.order}</span>
                          <span className={styles.chunkTag}>페이지 {chunk.page}</span>
                        </div>
                        {parsed ? (
                          <div className={styles.chunkTableScroll}>
                            <table className={styles.chunkTable}>
                              <thead>
                                <tr>
                                  {parsed.headers.map((h, idx) => (
                                    <th key={idx}>{h || `열 ${idx + 1}`}</th>
                                  ))}
                                </tr>
                              </thead>
                              <tbody>
                                {parsed.rows.map((row, rIdx) => (
                                  <tr key={rIdx}>
                                    {row.map((cell, cIdx) => (
                                      <td key={cIdx} className={isNumeric(cell) ? styles.cellNumeric : styles.cellText}>
                                        {cell || '-'}
                                      </td>
                                    ))}
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          </div>
                        ) : (
                          <pre className={styles.chunkContent}>{chunk.content}</pre>
                        )}
                      </div>
                    );
                  })}
                </div>
              </>
            )}
          </div>
              
          <div className={styles.stepCard}>
            <div className={`${styles.stepRow} ${styles.stepGroup} ${styles.stepSpacing}`}>
              <span className={`${styles.stepBadge} ${styles.stepTertiary}`}>단계 3</span>
              <div>
                <div className={styles.stepTitle}>리포트 생성</div>
                <div className={styles.subTitle}>선택한 파일로 PPT/PDF 리포트를 만듭니다.</div>
              </div>
            </div>
            <div className={styles.reportBar}>
              <div className={styles.reportControls}>
                <label className={styles.reportLabel} htmlFor="reportCategory">카테고리</label>
                <select
                  id="reportCategory"
                  className={styles.reportSelect}
                  value={reportCategory}
                  onChange={(e) => setReportCategory(e.target.value)}
                >
                  <option value="">선택</option>
                  <option value="capex">CAPEX</option>
                  <option value="depex">DEPEX</option>
                </select>
              </div>
              <div className={styles.reportActions}>
                <button
                  className={styles.reportButton}
                  type="button"
                  onClick={async () => {
                    const targetFile = selectedSheet || selectedRecentFile;
                    if (!targetFile) return;

                    setIsGeneratingReport(true);
                    setProcessingContext('report');
                    setUploadStatus({
                      status: 'processing',
                      message: '리포트 데이터(JSON) 생성 중...',
                      details: { name: targetFile }
                    });

                    const ganttPayload = reportCategory === 'capex'
                      ? { phase1: ganttTasksByPhase.phase1, phase2: ganttTasksByPhase.phase2 }
                      : [];

                    try {
                      const response = await fetch('/api/report/generate', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({
                          format: reportFormat,
                          category: reportCategory,
                          fileName: targetFile,
                          gantt: ganttPayload
                        })
                      });

                      const result = await response.json();
                      if (!response.ok || !result?.success) {
                        throw new Error(result?.error?.message || `리포트 데이터 생성 실패 (${response.status})`);
                      }

                      setUploadStatus({
                        status: 'success',
                        message: `리포트 데이터 생성 완료: ${result?.data?.outputDir || ''}`,
                        details: {
                          name: targetFile,
                          generatedFiles: result?.data?.files,
                          logs: result?.data?.logs
                        }
                      });
                    } catch (error) {
                      setUploadStatus({
                        status: 'error',
                        message: error instanceof Error ? error.message : '리포트 데이터 생성 중 오류가 발생했습니다.',
                        details: { name: targetFile }
                      });
                    } finally {
                      setIsGeneratingReport(false);
                    }
                  }}
                  disabled={
                    isGenerateDisabled
                  }
                >
                  {isGeneratingReport ? '데이터 생성 중...' : '리포트 생성'}
                </button>
              </div>
            </div>
            <div className={styles.resultPanel}>
              <div className={styles.resultHeader}>
                <div className={styles.sectionLabel}>리포트 생성 결과</div>
                {reportPath && (
                  <a className={styles.resultDownloadButton} href={reportDownloadUrl} download>
                    다운로드
                  </a>
                )}
              </div>
              {reportPath ? (
                <>
                  <div className={styles.resultRow}>
                    <span className={styles.resultLabel}>문서 경로</span>
                    <code className={styles.resultPath}>{reportPath}</code>
                  </div>
                  <div className={styles.resultRow}>
                    <span className={styles.resultLabel}>JSON 산출물</span>
                    <span className={styles.resultMetaText}>{generatedFiles?.capex}, {generatedFiles?.opex_year1}, {generatedFiles?.general}</span>
                  </div>
                  <div className={styles.resultLogBox}>
                    {generationLogs.slice(-4).map((log, idx) => (
                      <div key={`${log.step || 'step'}-${idx}`} className={styles.resultLogRow}>
                        <span className={styles.resultStep}>{log.step || `step-${idx + 1}`}</span>
                        <span className={styles.resultLogStatus}>{(log.stderr || '').trim() ? '경고/확인 필요' : '완료'}</span>
                      </div>
                    ))}
                  </div>
                </>
              ) : (
                <div className={styles.resultEmpty}>리포트를 생성하면 문서 경로와 처리 로그 요약이 표시됩니다.</div>
              )}
            </div>
          </div>

          {reportCategory === 'capex' && (
            <div className={styles.stepCard}>
              <div className={styles.previewHeader}>
                <div>
                  <div className={styles.sectionLabel}>CAPEX 간트 차트</div>
                  <div className={styles.subTitle}>주차 기준(Week Index)으로 진행 구간 표시</div>
                </div>
                <button
                  type="button"
                  className={styles.previewToggleButton}
                  onClick={() => setIsGanttCollapsed((prev) => !prev)}
                >
                  {isGanttCollapsed ? '펼치기' : '접기'}
                </button>
              </div>
              {isGanttCollapsed ? (
                <div className={styles.previewCollapsedHint}>
                  <div className={styles.previewSummaryTitle}>간트 차트 요약</div>
                  <div className={styles.previewSummaryList}>
                    <span className={styles.previewSummaryItem}>작업 수 {totalGanttTaskCount}개</span>
                    <span className={styles.previewSummaryItem}>{ganttPhase === 'phase1' ? '현재 Phase 1' : '현재 Phase 2'}</span>
                    <span className={styles.previewSummaryItem}>{hasInvalidGantt ? '오류 있음' : '유효성 정상'}</span>
                  </div>
                </div>
              ) : (
                <>
                  <div className={styles.ganttPhaseTabs}>
                    <button
                      type="button"
                      className={`${styles.ganttPhaseTab} ${ganttPhase === 'phase1' ? styles.ganttPhaseTabActive : ''}`}
                      onClick={() => setGanttPhase('phase1')}
                    >
                      Phase 1
                    </button>
                    <button
                      type="button"
                      className={`${styles.ganttPhaseTab} ${ganttPhase === 'phase2' ? styles.ganttPhaseTabActive : ''}`}
                      onClick={() => setGanttPhase('phase2')}
                    >
                      Phase 2
                    </button>
                  </div>
                  <div className={styles.ganttWrapper}>
                <div className={styles.ganttHeaderRow}>
                  <div className={styles.ganttHeaderLabelSpacer} />
                  <div>
                    <div className={styles.ganttMonthHeader} style={{ width: `${(capexMaxWeek + 1) * ganttColWidth}px` }}>
                      {ganttMonthSegments.map((segment, idx) => (
                        <div
                          key={`${segment.label}-${idx}`}
                          className={styles.ganttMonthCell}
                          style={{ width: `${segment.weeks * ganttColWidth}px` }}
                        >
                          {segment.label}
                        </div>
                      ))}
                    </div>
                    <div className={styles.ganttHeader} style={{ width: `${(capexMaxWeek + 1) * ganttColWidth}px` }}>
                      {Array.from({ length: capexMaxWeek + 1 }).map((_, idx) => (
                        <div key={idx} className={styles.ganttWeek} style={{ width: `${ganttColWidth}px` }}>
                          W{idx + 1}
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
                <div className={styles.ganttBody}>
                  {ganttRenderRows.map((task) => {
                    const duration = task.endWeekIndex - task.startWeekIndex + 1;
                    const left = task.startWeekIndex * ganttColWidth;
                    const width = duration * ganttColWidth;
                    const isGroup = (task as any).isGroup;
                    return (
                      <div key={task.id} className={styles.ganttRow}>
                        <div className={`${styles.ganttLabel} ${isGroup ? styles.ganttLabelGroup : styles.ganttLabelTask}`}>
                          {task.name}
                        </div>
                        <div className={styles.ganttTrack} style={{ width: `${(capexMaxWeek + 1) * ganttColWidth}px` }}>
                          <div
                            className={isGroup ? styles.ganttBarGroup : styles.ganttBar}
                            style={{
                              left: `${left}px`,
                              width: `${width}px`
                            }}
                          >
                            <span className={styles.ganttDuration}>{duration}w</span>
                            <span className={styles.ganttRange}>W{task.startWeekIndex + 1} - W{task.endWeekIndex + 1}</span>
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
              <div className={styles.ganttEditWrapper}>
                <div className={styles.ganttEditHeaderRow}>
                  <div className={styles.sectionLabel}>기간 편집 (주차 Index · {ganttPhase === 'phase1' ? 'Phase 1' : 'Phase 2'})</div>
                  <div className={styles.ganttEditActions}>
                    <button
                      type="button"
                      className={styles.ganttSaveButton}
                      onClick={async () => {
                        const ok = await saveCapexGanttToServer(ganttTasksByPhase);
                        setGanttSaveMessage(ok ? '저장되었습니다.' : '저장에 실패했습니다.');
                      }}
                    >
                      저장
                    </button>
                    {ganttSaveMessage && <span className={styles.ganttSaveMessage}>{ganttSaveMessage}</span>}
                  </div>
                </div>
                <div className={styles.ganttEditTable}>
                  <div className={styles.ganttEditHeader}>작업</div>
                  <div className={styles.ganttEditHeader}>Start W</div>
                  <div className={styles.ganttEditHeader}>End W</div>
                  {activeGanttTasks.filter((task) => task.type !== 'group').map((task) => (
                    <React.Fragment key={task.id}>
                      <div className={styles.ganttEditCell}>{task.name}</div>
                      <div className={styles.ganttEditCell}>
                        <input
                          type="number"
                          min={0}
                          value={task.startWeekIndex}
                          onChange={(e) => {
                            const val = Number(e.target.value);
                            setGanttTasksByPhase((prev) => ({
                              ...prev,
                              [ganttPhase]: prev[ganttPhase].map((t) =>
                                t.id === task.id ? { ...t, startWeekIndex: val } : t
                              )
                            }));
                          }}
                        />
                      </div>
                      <div className={styles.ganttEditCell}>
                        <input
                          type="number"
                          min={0}
                          value={task.endWeekIndex}
                          onChange={(e) => {
                            const val = Number(e.target.value);
                            setGanttTasksByPhase((prev) => ({
                              ...prev,
                              [ganttPhase]: prev[ganttPhase].map((t) =>
                                t.id === task.id ? { ...t, endWeekIndex: val } : t
                              )
                            }));
                          }}
                        />
                      </div>
                    </React.Fragment>
                  ))}
                </div>
                {hasInvalidGantt && (
                  <div className={styles.ganttError}>End W는 Start W 이상이어야 합니다.</div>
                )}
              </div>
                </>
              )}
            </div>
          )}
        </section>

      </main>

      <PVSolarExcelPreprocessModal
        isOpen={isModalOpen}
        onClose={() => setIsModalOpen(false)}
        onProcess={handleProcess}
      />
    </div>
  );
}
