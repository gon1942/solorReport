"use client";

import { useEffect, useMemo, useState } from 'react';
import { Activity, AlertTriangle, CheckCircle2, Clock3, RefreshCcw } from 'lucide-react';
import styles from './page.module.css';

type PlannerStatus = {
  id: string;
  name: string;
  status: 'queued' | 'running' | 'completed' | 'failed';
  progress: number;
  owner: string;
  priority: 'low' | 'normal' | 'high';
  updatedAt: string;
  message?: string;
};

type StatusResponse = {
  success: boolean;
  source: string;
  data: PlannerStatus[];
  error?: string;
};

const statusLabels: Record<PlannerStatus['status'], string> = {
  queued: '대기',
  running: '진행 중',
  completed: '완료',
  failed: '실패'
};

const priorityLabels: Record<PlannerStatus['priority'], string> = {
  high: '높음',
  normal: '보통',
  low: '낮음'
};

export default function PlannerStatusPage() {
  const [items, setItems] = useState<PlannerStatus[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [filterStatus, setFilterStatus] = useState<'all' | PlannerStatus['status']>('all');
  const [filterPriority, setFilterPriority] = useState<'all' | PlannerStatus['priority']>('all');
  const [query, setQuery] = useState('');
  const [lastUpdated, setLastUpdated] = useState<string | null>(null);
  const [source, setSource] = useState<string>('');

  const fetchStatuses = async () => {
    setLoading(true);
    setError(null);
    try {
      const res = await fetch('/api/planner/status');
      const json = (await res.json()) as StatusResponse;
      if (!json.success) {
        setError(json.error || '상태를 불러오지 못했습니다.');
      }
      setItems(json.data || []);
      setSource(json.source || '');
      setLastUpdated(new Date().toISOString());
    } catch (err: any) {
      setError(err?.message || '상태를 불러오지 못했습니다.');
      setItems([]);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchStatuses();
    const timer = setInterval(fetchStatuses, 10000);
    return () => clearInterval(timer);
  }, []);

  const filtered = useMemo(() => {
    return items.filter((item) => {
      const matchStatus = filterStatus === 'all' || item.status === filterStatus;
      const matchPriority = filterPriority === 'all' || item.priority === filterPriority;
      const matchQuery = query.trim() === '' || item.name.toLowerCase().includes(query.trim().toLowerCase()) || item.id.toLowerCase().includes(query.trim().toLowerCase());
      return matchStatus && matchPriority && matchQuery;
    });
  }, [items, filterStatus, filterPriority, query]);

  const summary = useMemo(() => {
    const base = { running: 0, queued: 0, completed: 0, failed: 0 } as Record<PlannerStatus['status'], number>;
    items.forEach((item) => {
      base[item.status] += 1;
    });
    return base;
  }, [items]);

  const formatTime = (value: string) => {
    if (!value) return '';
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return value;
    return new Intl.DateTimeFormat('ko-KR', {
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit'
    }).format(date);
  };

  return (
    <div className={styles.page}>
      <header className={styles.header}>
        <div>
          <h1 className={styles.title}>Planner 상태 관리</h1>
          <p className={styles.subtitle}>REC sales under discussion with Samcheok Bluepower(SBP) at KRW 189/kWh (20-year fixed term)</p>
        </div>
        <div className={styles.headerActions}>
          <button className={styles.refreshButton} onClick={fetchStatuses} disabled={loading}>
            <RefreshCcw size={16} />
            <span>{loading ? '갱신 중...' : '새로고침'}</span>
          </button>
          {source && <span className={styles.sourceBadge}>{source}</span>}
        </div>
      </header>

      <section className={styles.summaryGrid}>
        <div className={`${styles.summaryCard} ${styles.cardRunning}`}>
          <div className={styles.summaryIcon}><Activity size={18} /></div>
          <div className={styles.summaryLabel}>진행</div>
          <div className={styles.summaryValue}>{summary.running}</div>
        </div>
        <div className={`${styles.summaryCard} ${styles.cardQueued}`}>
          <div className={styles.summaryIcon}><Clock3 size={18} /></div>
          <div className={styles.summaryLabel}>대기</div>
          <div className={styles.summaryValue}>{summary.queued}</div>
        </div>
        <div className={`${styles.summaryCard} ${styles.cardCompleted}`}>
          <div className={styles.summaryIcon}><CheckCircle2 size={18} /></div>
          <div className={styles.summaryLabel}>완료</div>
          <div className={styles.summaryValue}>{summary.completed}</div>
        </div>
        <div className={`${styles.summaryCard} ${styles.cardFailed}`}>
          <div className={styles.summaryIcon}><AlertTriangle size={18} /></div>
          <div className={styles.summaryLabel}>실패</div>
          <div className={styles.summaryValue}>{summary.failed}</div>
        </div>
      </section>

      <section className={styles.filters}>
        <div className={styles.filterGroup}>
          <label>상태</label>
          <select value={filterStatus} onChange={(e) => setFilterStatus(e.target.value as any)}>
            <option value="all">전체</option>
            <option value="running">진행 중</option>
            <option value="queued">대기</option>
            <option value="completed">완료</option>
            <option value="failed">실패</option>
          </select>
        </div>
        <div className={styles.filterGroup}>
          <label>우선순위</label>
          <select value={filterPriority} onChange={(e) => setFilterPriority(e.target.value as any)}>
            <option value="all">전체</option>
            <option value="high">높음</option>
            <option value="normal">보통</option>
            <option value="low">낮음</option>
          </select>
        </div>
        <div className={styles.searchBox}>
          <input
            type="text"
            placeholder="작업명 또는 ID 검색"
            value={query}
            onChange={(e) => setQuery(e.target.value)}
          />
        </div>
      </section>

      {error && <div className={styles.errorBox}>{error}</div>}

      <section className={styles.list}>
        {filtered.length === 0 && !loading && <div className={styles.empty}>표시할 작업이 없습니다.</div>}
        {filtered.map((item) => {
          const statusClass = {
            running: styles.badgeRunning,
            queued: styles.badgeQueued,
            completed: styles.badgeCompleted,
            failed: styles.badgeFailed
          }[item.status];

          const priorityClass = {
            high: styles.priorityHigh,
            normal: styles.priorityNormal,
            low: styles.priorityLow
          }[item.priority];

          const progress = Math.min(100, Math.max(0, item.progress));

          return (
            <article key={item.id} className={styles.card}>
              <header className={styles.cardHeader}>
                <div>
                  <div className={styles.cardTitle}>{item.name}</div>
                  <div className={styles.cardMeta}>ID: {item.id} · 담당: {item.owner}</div>
                </div>
                <div className={styles.chips}>
                  <span className={`${styles.badge} ${statusClass}`}>{statusLabels[item.status]}</span>
                  <span className={`${styles.badge} ${priorityClass}`}>{priorityLabels[item.priority]}</span>
                </div>
              </header>

              <div className={styles.progressBar}>
                <div className={styles.progressTrack}>
                  <div className={styles.progressFill} style={{ width: `${progress}%` }} />
                </div>
                <span className={styles.progressLabel}>{progress}%</span>
              </div>

              {item.message && <p className={styles.message}>{item.message}</p>}

              <footer className={styles.cardFooter}>
                <span>업데이트: {formatTime(item.updatedAt)}</span>
              </footer>
            </article>
          );
        })}
      </section>

      {lastUpdated && (
        <div className={styles.timestamp}>마지막 갱신: {formatTime(lastUpdated)}</div>
      )}
    </div>
  );
}
