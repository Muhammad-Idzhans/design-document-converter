"use client";

import React, { useState, useEffect, useMemo } from 'react';
import {
    Card,
    Row,
    Col,
    Typography,
    Tag,
    Progress,
    Table,
    Tooltip,
    Button,
    Spin,
    Empty,
    message,
    Space,
    Select,
    Segmented
} from 'antd';
import {
    DollarOutlined,
    FileTextOutlined,
    CheckCircleOutlined,
    ThunderboltOutlined,
    SyncOutlined,
    ArrowRightOutlined,
    EyeOutlined,
    PieChartOutlined,
    RobotOutlined,
    PictureOutlined,
    ReadOutlined,
    RiseOutlined,
    CalendarOutlined
} from '@ant-design/icons';
import {
    ResponsiveContainer,
    AreaChart,
    Area,
    XAxis,
    YAxis,
    CartesianGrid,
    Tooltip as RechartsTooltip,
} from 'recharts';
import { useRouter } from 'next/navigation';
import dayjs from 'dayjs';
import { resolveCost } from '@/lib/cost';

const { Title, Text } = Typography;

// Color palette for dynamically discovered document types
const DYNAMIC_PALETTE = ['blue', 'purple', 'cyan', 'geekblue', 'orange', 'green', 'magenta', 'gold'];

const PRESET_COLORS: Record<string, string> = {
    'Design Document': 'blue',
    'Final Configuration': 'purple',
};

function normalizeDocType(type?: string): string {
    if (!type) return 'Design Document';
    const lower = type.toLowerCase().trim();
    if (lower === 'final-config' || lower === 'final configuration document' || lower === 'final configuration') {
        return 'Final Configuration';
    }
    if (lower === 'design-document' || lower === 'design document') {
        return 'Design Document';
    }
    return type;
}

const MONTH_NAMES = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const MONTH_NAMES_FULL = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
];

interface MonthlyDataPoint {
    month: string;
    monthFull: string;
    monthIndex: number;
    spendingMyr: number;
    spendingUsd: number;
    documents: number;
    completedDocs: number;
    failedDocs: number;
    tokens: number;
}

export default function DashboardPage() {
    const router = useRouter();
    const [tasks, setTasks] = useState<any[]>([]);
    const [settings, setSettings] = useState<any>(null);
    const [loading, setLoading] = useState(true);
    const [refreshing, setRefreshing] = useState(false);
    const [lastUpdated, setLastUpdated] = useState<string>('');
    const [selectedYear, setSelectedYear] = useState<number>(dayjs().year());
    const [activeMetric, setActiveMetric] = useState<'spending' | 'documents' | 'tokens'>('spending');
    const [isMounted, setIsMounted] = useState(false);

    useEffect(() => {
        setIsMounted(true);
    }, []);

    const fetchData = async (isManual = false) => {
        if (isManual) setRefreshing(true);
        try {
            const [historyRes, settingsRes] = await Promise.all([
                fetch('/api/history'),
                fetch('/api/settings')
            ]);

            const historyData = await historyRes.json();
            const settingsData = await settingsRes.json();

            if (historyData.data) setTasks(historyData.data);
            if (settingsData && !settingsData.error) setSettings(settingsData);
            setLastUpdated(dayjs().format('h:mm:ss A'));
        } catch (error) {
            console.error("Dashboard fetch error:", error);
            message.error("Failed to load dashboard data.");
        } finally {
            setLoading(false);
            setRefreshing(false);
        }
    };

    useEffect(() => {
        fetchData();
    }, []);

    // ─────────────────────────────────────────────────────────────────────────
    // COMPUTED ANALYTICS & METRICS
    // ─────────────────────────────────────────────────────────────────────────

    // 1. Overall Totals & Spending
    const metrics = useMemo(() => {
        let totalMyr = 0;
        let totalUsd = 0;
        let totalTokens = 0;
        let totalPromptTokens = 0;
        let totalCompletionTokens = 0;
        let completedCount = 0;
        let failedCount = 0;
        let inProgressCount = 0;

        let totalVisionUsd = 0;
        let totalLlmUsd = 0;
        let totalCuUsd = 0;

        const exchangeRate = Number(settings?.exchangeRate || 4.45);

        tasks.forEach((task) => {
            // Cost calculation
            const cost = resolveCost(task, settings);
            if (cost) {
                totalMyr += cost.costMyr;
                totalUsd += cost.costUsd;
            }

            // Status counts
            if (task.status === 'completed') {
                completedCount++;
            } else if (task.status === 'failed') {
                failedCount++;
            } else {
                inProgressCount++;
            }

            // Token analytics
            const tMetrics = task.cost_metrics;
            if (tMetrics) {
                const pTokens = (tMetrics.llm_tokens_prompt || 0) + (tMetrics.vision_tokens_prompt || 0);
                const cTokens = (tMetrics.llm_tokens_completion || 0) + (tMetrics.vision_tokens_completion || 0);
                totalPromptTokens += pTokens;
                totalCompletionTokens += cTokens;
                totalTokens += (tMetrics.total_tokens || (pTokens + cTokens));
            }

            // Service Breakdown
            if (task.cost_summary?.breakdown_usd) {
                totalVisionUsd += task.cost_summary.breakdown_usd.vision || 0;
                totalLlmUsd += task.cost_summary.breakdown_usd.llm || 0;
                totalCuUsd += task.cost_summary.breakdown_usd.content_understanding || 0;
            } else if (tMetrics) {
                const vInRate = (settings?.visionInput ?? 2.50) / 1_000_000;
                const vOutRate = (settings?.visionOutput ?? 10.00) / 1_000_000;
                const lInRate = (settings?.llmInput ?? 2.00) / 1_000_000;
                const lOutRate = (settings?.llmCompletion ?? 8.00) / 1_000_000;
                const cuRate = (settings?.contentUnderstanding ?? 5.00) / 1_000;

                totalVisionUsd += ((tMetrics.vision_tokens_prompt || 0) * vInRate) + ((tMetrics.vision_tokens_completion || 0) * vOutRate);
                totalLlmUsd += ((tMetrics.llm_tokens_prompt || 0) * lInRate) + ((tMetrics.llm_tokens_completion || 0) * lOutRate);
                totalCuUsd += (tMetrics.content_understanding_pages || 0) * cuRate;
            }
        });

        const totalDocs = tasks.length;
        const successRate = totalDocs > 0 ? Math.round((completedCount / totalDocs) * 100) : 100;
        const avgCostMyr = completedCount > 0 ? (totalMyr / completedCount) : 0;

        // Service breakdown in MYR
        const visionMyr = totalVisionUsd * exchangeRate;
        const llmMyr = totalLlmUsd * exchangeRate;
        const cuMyr = totalCuUsd * exchangeRate;
        const serviceTotalUsd = totalVisionUsd + totalLlmUsd + totalCuUsd;

        return {
            totalMyr,
            totalUsd,
            totalDocs,
            completedCount,
            failedCount,
            inProgressCount,
            successRate,
            avgCostMyr,
            totalTokens,
            totalPromptTokens,
            totalCompletionTokens,
            services: {
                llm: { usd: totalLlmUsd, myr: llmMyr, pct: serviceTotalUsd > 0 ? Math.round((totalLlmUsd / serviceTotalUsd) * 100) : 0 },
                vision: { usd: totalVisionUsd, myr: visionMyr, pct: serviceTotalUsd > 0 ? Math.round((totalVisionUsd / serviceTotalUsd) * 100) : 0 },
                cu: { usd: totalCuUsd, myr: cuMyr, pct: serviceTotalUsd > 0 ? Math.round((totalCuUsd / serviceTotalUsd) * 100) : 0 },
            }
        };
    }, [tasks, settings]);

    // 2. Dynamic Document Type Aggregation (Supports any number of future document types)
    const documentTypeStats = useMemo(() => {
        if (!tasks.length) return [];

        const groupMap: Record<string, { count: number; completed: number; totalCostMyr: number }> = {};

        tasks.forEach((task) => {
            const type = normalizeDocType(task.documentType);
            if (!groupMap[type]) {
                groupMap[type] = { count: 0, completed: 0, totalCostMyr: 0 };
            }
            groupMap[type].count += 1;
            if (task.status === 'completed') groupMap[type].completed += 1;

            const cost = resolveCost(task, settings);
            if (cost) groupMap[type].totalCostMyr += cost.costMyr;
        });

        return Object.entries(groupMap)
            .map(([type, data], index) => ({
                type,
                count: data.count,
                completed: data.completed,
                totalCostMyr: data.totalCostMyr,
                percent: Math.round((data.count / tasks.length) * 100),
                color: PRESET_COLORS[type] || DYNAMIC_PALETTE[index % DYNAMIC_PALETTE.length],
            }))
            .sort((a, b) => b.count - a.count);
    }, [tasks, settings]);

    // 3. Dynamic Year Options (includes current year and all years from tasks, sorted descending)
    const availableYears = useMemo(() => {
        const currentYear = dayjs().year();
        const taskYears = tasks
            .map((t) => (t.createdAt ? dayjs(t.createdAt).year() : null))
            .filter((y): y is number => typeof y === 'number' && !isNaN(y));
        const allYears = Array.from(new Set([currentYear, ...taskYears]));
        return allYears.sort((a, b) => b - a);
    }, [tasks]);

    // 4. Monthly Aggregation for Selected Year (12 months: Jan - Dec)
    const monthlyChartData = useMemo(() => {
        const buckets: MonthlyDataPoint[] = MONTH_NAMES.map((name, idx) => ({
            month: name,
            monthFull: MONTH_NAMES_FULL[idx],
            monthIndex: idx,
            spendingMyr: 0,
            spendingUsd: 0,
            documents: 0,
            completedDocs: 0,
            failedDocs: 0,
            tokens: 0,
        }));

        tasks.forEach((task) => {
            if (!task.createdAt) return;
            const taskDate = dayjs(task.createdAt);
            if (taskDate.year() !== selectedYear) return;

            const monthIdx = taskDate.month();
            if (monthIdx < 0 || monthIdx > 11) return;

            const bucket = buckets[monthIdx];
            bucket.documents += 1;
            if (task.status === 'completed') {
                bucket.completedDocs += 1;
            } else if (task.status === 'failed') {
                bucket.failedDocs += 1;
            }

            const cost = resolveCost(task, settings);
            if (cost) {
                bucket.spendingMyr += cost.costMyr;
                bucket.spendingUsd += cost.costUsd;
            }

            const tMetrics = task.cost_metrics;
            if (tMetrics) {
                const pTokens = (tMetrics.llm_tokens_prompt || 0) + (tMetrics.vision_tokens_prompt || 0);
                const cTokens = (tMetrics.llm_tokens_completion || 0) + (tMetrics.vision_tokens_completion || 0);
                bucket.tokens += (tMetrics.total_tokens || (pTokens + cTokens));
            }
        });

        return buckets.map((b) => ({
            ...b,
            spendingMyr: Math.round(b.spendingMyr * 100) / 100,
            spendingUsd: Math.round(b.spendingUsd * 100) / 100,
        }));
    }, [tasks, selectedYear, settings]);

    // Annual cumulative totals for the selected year
    const yearTotal = useMemo(() => {
        return monthlyChartData.reduce(
            (acc, curr) => {
                acc.spendingMyr += curr.spendingMyr;
                acc.documents += curr.documents;
                acc.tokens += curr.tokens;
                return acc;
            },
            { spendingMyr: 0, documents: 0, tokens: 0 }
        );
    }, [monthlyChartData]);

    // Active metric display configuration
    const metricConfig = useMemo(() => {
        switch (activeMetric) {
            case 'spending':
                return {
                    title: 'Total Spending',
                    dataKey: 'spendingMyr',
                    strokeColor: '#1677ff',
                    fillGradientId: 'colorSpending',
                    gradientStopColor: '#1677ff',
                    yAxisFormatter: (v: number) => `RM ${v}`,
                };
            case 'documents':
                return {
                    title: 'Documents Generated',
                    dataKey: 'documents',
                    strokeColor: '#722ed1',
                    fillGradientId: 'colorDocs',
                    gradientStopColor: '#722ed1',
                    yAxisFormatter: (v: number) => `${v}`,
                };
            case 'tokens':
                return {
                    title: 'AI Tokens Consumed',
                    dataKey: 'tokens',
                    strokeColor: '#13c2c2',
                    fillGradientId: 'colorTokens',
                    gradientStopColor: '#13c2c2',
                    yAxisFormatter: (v: number) => {
                        if (v >= 1_000_000) return `${(v / 1_000_000).toFixed(1)}M`;
                        if (v >= 1_000) return `${Math.round(v / 1_000)}k`;
                        return `${v}`;
                    },
                };
        }
    }, [activeMetric]);

    // Custom floating tooltip for chart
    const CustomChartTooltip = ({ active, payload }: any) => {
        if (active && payload && payload.length) {
            const data: MonthlyDataPoint = payload[0].payload;
            return (
                <div
                    style={{
                        backgroundColor: '#ffffff',
                        border: '1px solid #f0f0f0',
                        borderRadius: '8px',
                        padding: '12px 16px',
                        boxShadow: '0 6px 16px rgba(0, 0, 0, 0.08)',
                        minWidth: '220px',
                    }}
                >
                    <div className="fw-bold text-dark mb-2" style={{ fontSize: '13px' }}>
                        {data.monthFull} {selectedYear}
                    </div>
                    <div className="d-flex align-items-center justify-content-between mb-1">
                        <span className="text-muted small">Total Spending:</span>
                        <span className="fw-bold" style={{ color: '#1677ff', fontSize: '13px' }}>
                            RM {data.spendingMyr.toFixed(2)}
                            <span className="text-muted fw-normal ms-1" style={{ fontSize: '11px' }}>
                                (${data.spendingUsd.toFixed(2)})
                            </span>
                        </span>
                    </div>
                    <div className="d-flex align-items-center justify-content-between mb-1">
                        <span className="text-muted small">Documents:</span>
                        <span className="fw-bold" style={{ color: '#722ed1', fontSize: '13px' }}>
                            {data.documents} doc{data.documents === 1 ? '' : 's'}
                        </span>
                    </div>
                    {data.documents > 0 && (
                        <div className="d-flex align-items-center justify-content-between mb-1 ps-2" style={{ fontSize: '11px' }}>
                            <span className="text-muted">Breakdown:</span>
                            <span className="text-secondary">
                                {data.completedDocs} ok{data.failedDocs > 0 ? `, ${data.failedDocs} failed` : ''}
                            </span>
                        </div>
                    )}
                    <div className="d-flex align-items-center justify-content-between pt-1 mt-1 border-top">
                        <span className="text-muted small">AI Tokens:</span>
                        <span className="fw-semibold" style={{ color: '#13c2c2', fontSize: '13px' }}>
                            {data.tokens > 1_000_000
                                ? `${(data.tokens / 1_000_000).toFixed(2)}M`
                                : data.tokens.toLocaleString()}
                        </span>
                    </div>
                </div>
            );
        }
        return null;
    };

    // 5. Recent 5 Tasks for the quick view table
    const recentTasks = useMemo(() => {
        return tasks.slice(0, 5);
    }, [tasks]);

    // Mini table columns for Recent Tasks
    const recentColumns = [
        {
            title: <span style={{ whiteSpace: 'nowrap' }}>Task ID</span>,
            dataIndex: 'taskId',
            key: 'taskId',
            width: 120,
            render: (text: string) => {
                const shortId = text ? (text.length > 8 ? `${text.substring(0, 8)}...` : text) : '-';
                return (
                    <Tooltip title={text}>
                        <Text type="secondary" style={{ fontFamily: 'monospace', fontSize: '13px', whiteSpace: 'nowrap' }}>
                            {shortId}
                        </Text>
                    </Tooltip>
                );
            }
        },
        {
            title: 'Document',
            key: 'document',
            render: (_: any, record: any) => {
                const name = record.generated_filename || record.filename || "Generated_Document";
                return (
                    <Tooltip title={name}>
                        <div
                            className="fw-medium text-dark"
                            style={{
                                maxWidth: '240px',
                                overflow: 'hidden',
                                textOverflow: 'ellipsis',
                                whiteSpace: 'nowrap'
                            }}
                        >
                            {name}
                        </div>
                    </Tooltip>
                );
            }
        },
        {
            title: <span style={{ whiteSpace: 'nowrap' }}>Type</span>,
            dataIndex: 'documentType',
            key: 'documentType',
            width: 160,
            render: (text: string) => {
                const norm = normalizeDocType(text);
                const color = PRESET_COLORS[norm] || 'geekblue';
                return (
                    <Tag
                        color={color}
                        style={{
                            borderRadius: '4px',
                            border: 'none',
                            padding: '2px 8px',
                            fontWeight: 500,
                            whiteSpace: 'nowrap'
                        }}
                    >
                        {norm}
                    </Tag>
                );
            }
        },
        {
            title: <span style={{ whiteSpace: 'nowrap' }}>Status</span>,
            dataIndex: 'status',
            key: 'status',
            width: 130,
            render: (status: string, record: any) => {
                const color = status === 'completed' ? 'success' : (status === 'failed' ? 'error' : 'processing');
                const label = status === 'completed' ? 'Completed' : (status === 'failed' ? 'Failed' : 'Processing');
                return (
                    <div className="d-flex align-items-center gap-1">
                        <Tag color={color} style={{ borderRadius: '4px', border: 'none', padding: '2px 8px', whiteSpace: 'nowrap' }}>
                            {label}
                        </Tag>
                        {record.isDeleted && (
                            <Tooltip title="File purged from storage. Cost & token metrics retained.">
                                <Tag color="default" style={{ borderRadius: '4px', border: 'none', padding: '1px 5px', fontSize: '11px', whiteSpace: 'nowrap', color: '#8c8c8c' }}>
                                    Purged
                                </Tag>
                            </Tooltip>
                        )}
                    </div>
                );
            }
        },
        {
            title: <span style={{ whiteSpace: 'nowrap' }}>Cost</span>,
            key: 'cost',
            width: 130,
            align: 'right' as const,
            render: (_: any, record: any) => {
                const cost = resolveCost(record, settings);
                if (!cost) return '-';
                return (
                    <Tooltip
                        title={
                            cost.isLocked
                                ? `Rate locked at 1 USD = RM ${cost.exchangeRate.toFixed(2)}`
                                : `Estimated using current rate: 1 USD = RM ${cost.exchangeRate.toFixed(2)}`
                        }
                    >
                        <span className="fw-medium" style={{ whiteSpace: 'nowrap', cursor: 'help' }}>
                            RM{cost.costMyr.toFixed(2)}{' '}
                            <Text type="secondary" style={{ fontSize: '12px' }}>
                                (${cost.costUsd.toFixed(2)})
                            </Text>
                        </span>
                    </Tooltip>
                );
            }
        },
        {
            title: <span style={{ whiteSpace: 'nowrap' }}>Date</span>,
            dataIndex: 'createdAt',
            key: 'date',
            width: 120,
            render: (text: string) => text ? dayjs(text).format('D MMM, h:mm A') : '-'
        }
    ];

    if (loading) {
        return (
            <div className="container-fluid bg-light d-flex justify-content-center align-items-center flex-grow-1" style={{ minHeight: '80vh' }}>
                <Spin size="large" />
            </div>
        );
    }

    return (
        <div className="container-fluid bg-light d-flex flex-column flex-grow-1 py-4 pb-5">
            <div className="container animate__animated animate__fadeIn w-100" style={{ maxWidth: '1200px' }}>

                {/* Header (Matching Design Document & Final Config typography) */}
                <div className="d-flex justify-content-between align-items-center mt-3 mb-4 flex-wrap gap-2">
                    <div>
                        <h2 className="fw-bold mb-2">Dashboard</h2>
                        <p className="text-muted mb-0">
                            Document generation analytics, volume statistics, and spending overview.
                        </p>
                    </div>
                    <div className="d-flex align-items-center gap-3">
                        {lastUpdated && (
                            <span className="text-muted small">
                                Last updated: <span className="fw-medium">{lastUpdated}</span>
                            </span>
                        )}
                        <Button
                            icon={<SyncOutlined spin={refreshing} />}
                            onClick={() => fetchData(true)}
                            loading={refreshing}
                            style={{ borderRadius: '6px' }}
                        >
                            Refresh
                        </Button>
                    </div>
                </div>

                {/* Top Row: 4 Primary KPI Summary Cards */}
                <Row gutter={[16, 16]} className="mb-4">
                    {/* 1. Total Spending */}
                    <Col xs={24} sm={12} lg={6}>
                        <Card className="shadow-sm border-0 h-100" style={{ borderRadius: '12px' }}>
                            <div className="d-flex justify-content-between align-items-start mb-2">
                                <span className="text-muted fw-semibold text-uppercase" style={{ fontSize: '12px', letterSpacing: '0.5px' }}>
                                    Total Spend
                                </span>
                                <div
                                    style={{
                                        width: '36px',
                                        height: '36px',
                                        borderRadius: '8px',
                                        backgroundColor: '#e6f4ff',
                                        color: '#1677ff',
                                        display: 'flex',
                                        alignItems: 'center',
                                        justifyContent: 'center',
                                        fontSize: '18px'
                                    }}
                                >
                                    <DollarOutlined />
                                </div>
                            </div>
                            <div className="fw-bold" style={{ fontSize: '26px', color: '#1c2b36', lineHeight: 1.2 }}>
                                RM {metrics.totalMyr.toFixed(2)}
                            </div>
                            <div className="text-muted small mt-2 d-flex align-items-center gap-1">
                                <span>≈ ${metrics.totalUsd.toFixed(2)} USD</span>
                                <span>•</span>
                                <span>Avg RM {metrics.avgCostMyr.toFixed(2)}/doc</span>
                            </div>
                        </Card>
                    </Col>

                    {/* 2. Total Documents */}
                    <Col xs={24} sm={12} lg={6}>
                        <Card className="shadow-sm border-0 h-100" style={{ borderRadius: '12px' }}>
                            <div className="d-flex justify-content-between align-items-start mb-2">
                                <span className="text-muted fw-semibold text-uppercase" style={{ fontSize: '12px', letterSpacing: '0.5px' }}>
                                    Total Generated
                                </span>
                                <div
                                    style={{
                                        width: '36px',
                                        height: '36px',
                                        borderRadius: '8px',
                                        backgroundColor: '#f0f5ff',
                                        color: '#2f54eb',
                                        display: 'flex',
                                        alignItems: 'center',
                                        justifyContent: 'center',
                                        fontSize: '18px'
                                    }}
                                >
                                    <FileTextOutlined />
                                </div>
                            </div>
                            <div className="fw-bold" style={{ fontSize: '26px', color: '#1c2b36', lineHeight: 1.2 }}>
                                {metrics.totalDocs}
                            </div>
                            <div className="text-muted small mt-2">
                                <span>{documentTypeStats.length} document type{documentTypeStats.length === 1 ? '' : 's'} active</span>
                            </div>
                        </Card>
                    </Col>

                    {/* 3. Success Rate */}
                    <Col xs={24} sm={12} lg={6}>
                        <Card className="shadow-sm border-0 h-100" style={{ borderRadius: '12px' }}>
                            <div className="d-flex justify-content-between align-items-start mb-2">
                                <span className="text-muted fw-semibold text-uppercase" style={{ fontSize: '12px', letterSpacing: '0.5px' }}>
                                    Success Rate
                                </span>
                                <div
                                    style={{
                                        width: '36px',
                                        height: '36px',
                                        borderRadius: '8px',
                                        backgroundColor: metrics.successRate >= 80 ? '#f6ffed' : '#fff1f0',
                                        color: metrics.successRate >= 80 ? '#52c41a' : '#f5222d',
                                        display: 'flex',
                                        alignItems: 'center',
                                        justifyContent: 'center',
                                        fontSize: '18px'
                                    }}
                                >
                                    <CheckCircleOutlined />
                                </div>
                            </div>
                            <div className="fw-bold" style={{ fontSize: '26px', color: '#1c2b36', lineHeight: 1.2 }}>
                                {metrics.successRate}%
                            </div>
                            <div className="small mt-2 d-flex align-items-center gap-1">
                                <span className="text-success fw-medium">{metrics.completedCount} Success</span>
                                <span className="text-muted">•</span>
                                <span className={metrics.failedCount > 0 ? "text-danger fw-medium" : "text-muted"}>
                                    {metrics.failedCount} Failed
                                </span>
                            </div>
                        </Card>
                    </Col>

                    {/* 4. Total Tokens */}
                    <Col xs={24} sm={12} lg={6}>
                        <Card className="shadow-sm border-0 h-100" style={{ borderRadius: '12px' }}>
                            <div className="d-flex justify-content-between align-items-start mb-2">
                                <span className="text-muted fw-semibold text-uppercase" style={{ fontSize: '12px', letterSpacing: '0.5px' }}>
                                    Total AI Tokens
                                </span>
                                <div
                                    style={{
                                        width: '36px',
                                        height: '36px',
                                        borderRadius: '8px',
                                        backgroundColor: '#f9f0ff',
                                        color: '#722ed1',
                                        display: 'flex',
                                        alignItems: 'center',
                                        justifyContent: 'center',
                                        fontSize: '18px'
                                    }}
                                >
                                    <ThunderboltOutlined />
                                </div>
                            </div>
                            <div className="fw-bold" style={{ fontSize: '26px', color: '#1c2b36', lineHeight: 1.2 }}>
                                {metrics.totalTokens > 1_000_000
                                    ? `${(metrics.totalTokens / 1_000_000).toFixed(2)}M`
                                    : metrics.totalTokens.toLocaleString()}
                            </div>
                            <div className="text-muted small mt-2">
                                <span>LLM reasoning & Vision prompt tokens</span>
                            </div>
                        </Card>
                    </Col>
                </Row>

                {/* Middle Section: Dynamic Document Types Breakdown & AI Service Distribution */}
                <Row gutter={[16, 16]} className="mb-4">
                    {/* Left: Dynamic Document Types Breakdown */}
                    <Col xs={24} lg={12}>
                        <Card
                            className="shadow-sm border-0 h-100"
                            style={{ borderRadius: '12px' }}
                            title={
                                <div className="d-flex align-items-center gap-2 py-1">
                                    <PieChartOutlined style={{ color: '#1677ff', fontSize: '18px' }} />
                                    <span className="fw-semibold" style={{ fontSize: '16px' }}>Volume by Document Type</span>
                                </div>
                            }
                        >
                            {documentTypeStats.length === 0 ? (
                                <Empty image={Empty.PRESENTED_IMAGE_SIMPLE} description="No documents generated yet" />
                            ) : (
                                <div className="d-flex flex-column gap-4">
                                    {documentTypeStats.map((item) => (
                                        <div key={item.type}>
                                            <div className="d-flex justify-content-between align-items-center mb-1">
                                                <div className="d-flex align-items-center gap-2">
                                                    <Tag color={item.color} style={{ borderRadius: '4px', border: 'none', padding: '1px 8px', fontWeight: 500 }}>
                                                        {item.type}
                                                    </Tag>
                                                </div>
                                                <div className="d-flex align-items-center gap-3">
                                                    <span className="fw-medium text-dark">{item.count} doc{item.count === 1 ? '' : 's'}</span>
                                                    <span className="text-muted small" style={{ minWidth: '40px', textAlign: 'right' }}>
                                                        {item.percent}%
                                                    </span>
                                                    <span className="fw-semibold text-primary small" style={{ minWidth: '70px', textAlign: 'right' }}>
                                                        RM {item.totalCostMyr.toFixed(2)}
                                                    </span>
                                                </div>
                                            </div>
                                            <Progress
                                                percent={item.percent}
                                                showInfo={false}
                                                strokeColor={item.color === 'purple' ? '#722ed1' : (item.color === 'blue' ? '#1677ff' : '#13c2c2')}
                                                size="small"
                                            />
                                        </div>
                                    ))}
                                </div>
                            )}
                        </Card>
                    </Col>

                    {/* Right: AI Service Spending Breakdown */}
                    <Col xs={24} lg={12}>
                        <Card
                            className="shadow-sm border-0 h-100"
                            style={{ borderRadius: '12px' }}
                            title={
                                <div className="d-flex align-items-center gap-2 py-1">
                                    <RobotOutlined style={{ color: '#722ed1', fontSize: '18px' }} />
                                    <span className="fw-semibold" style={{ fontSize: '16px' }}>AI Service Cost Distribution</span>
                                </div>
                            }
                        >
                            <div className="d-flex flex-column gap-4">
                                {/* LLM Generation */}
                                <div>
                                    <div className="d-flex justify-content-between align-items-center mb-1">
                                        <div className="d-flex align-items-center gap-2">
                                            <div style={{ width: '28px', height: '28px', borderRadius: '6px', backgroundColor: '#f0f5ff', color: '#2f54eb', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                                                <RobotOutlined />
                                            </div>
                                            <div>
                                                <div className="fw-medium text-dark" style={{ fontSize: '14px' }}>LLM Synthesis</div>
                                                <div className="text-muted" style={{ fontSize: '11px' }}>GPT-4.1 / GPT-4o drafting</div>
                                            </div>
                                        </div>
                                        <div className="text-end">
                                            <div className="fw-semibold text-dark">RM {metrics.services.llm.myr.toFixed(2)}</div>
                                            <div className="text-muted small">{metrics.services.llm.pct}% of total</div>
                                        </div>
                                    </div>
                                    <Progress percent={metrics.services.llm.pct} showInfo={false} strokeColor="#2f54eb" size="small" />
                                </div>

                                {/* Vision Analysis */}
                                <div>
                                    <div className="d-flex justify-content-between align-items-center mb-1">
                                        <div className="d-flex align-items-center gap-2">
                                            <div style={{ width: '28px', height: '28px', borderRadius: '6px', backgroundColor: '#e6f4ff', color: '#1677ff', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                                                <PictureOutlined />
                                            </div>
                                            <div>
                                                <div className="fw-medium text-dark" style={{ fontSize: '14px' }}>Vision Analysis</div>
                                                <div className="text-muted" style={{ fontSize: '11px' }}>GPT-4o slide & diagram vision</div>
                                            </div>
                                        </div>
                                        <div className="text-end">
                                            <div className="fw-semibold text-dark">RM {metrics.services.vision.myr.toFixed(2)}</div>
                                            <div className="text-muted small">{metrics.services.vision.pct}% of total</div>
                                        </div>
                                    </div>
                                    <Progress percent={metrics.services.vision.pct} showInfo={false} strokeColor="#1677ff" size="small" />
                                </div>

                                {/* Content Understanding */}
                                <div>
                                    <div className="d-flex justify-content-between align-items-center mb-1">
                                        <div className="d-flex align-items-center gap-2">
                                            <div style={{ width: '28px', height: '28px', borderRadius: '6px', backgroundColor: '#f6ffed', color: '#52c41a', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                                                <ReadOutlined />
                                            </div>
                                            <div>
                                                <div className="fw-medium text-dark" style={{ fontSize: '14px' }}>Content Understanding</div>
                                                <div className="text-muted" style={{ fontSize: '11px' }}>Azure presentation parsing</div>
                                            </div>
                                        </div>
                                        <div className="text-end">
                                            <div className="fw-semibold text-dark">RM {metrics.services.cu.myr.toFixed(2)}</div>
                                            <div className="text-muted small">{metrics.services.cu.pct}% of total</div>
                                        </div>
                                    </div>
                                    <Progress percent={metrics.services.cu.pct} showInfo={false} strokeColor="#52c41a" size="small" />
                                </div>
                            </div>
                        </Card>
                    </Col>
                </Row>

                {/* Middle Row: Interactive Monthly Trends Chart */}
                <Card
                    className="shadow-sm border-0 mb-4"
                    style={{ borderRadius: '12px' }}
                >
                    <div className="d-flex flex-column flex-md-row justify-content-between align-items-md-center gap-3 mb-3">
                        <div>
                            <div className="d-flex align-items-center gap-2 mb-1">
                                <RiseOutlined style={{ color: '#1677ff', fontSize: '18px' }} />
                                <span className="fw-semibold" style={{ fontSize: '16px' }}>Monthly Trends & Utilization</span>
                            </div>
                            <div className="text-muted small">
                                Track financial spend, generation volume, and token consumption throughout {selectedYear}.
                            </div>
                        </div>

                        <div className="d-flex flex-wrap align-items-center gap-2">
                            <div className="d-flex align-items-center gap-1">
                                <CalendarOutlined className="text-muted small" />
                                <Select
                                    value={selectedYear}
                                    onChange={(val) => setSelectedYear(val)}
                                    style={{ width: 105 }}
                                    options={availableYears.map((yr) => ({ label: `${yr}`, value: yr }))}
                                />
                            </div>
                            <Segmented
                                value={activeMetric}
                                onChange={(val) => setActiveMetric(val as any)}
                                options={[
                                    { label: 'Spending (RM)', value: 'spending', icon: <DollarOutlined /> },
                                    { label: 'Documents', value: 'documents', icon: <FileTextOutlined /> },
                                    { label: 'AI Tokens', value: 'tokens', icon: <ThunderboltOutlined /> },
                                ]}
                            />
                        </div>
                    </div>

                    {/* Annual Total Stat Tag for active metric */}
                    <div className="d-flex align-items-center gap-2 mb-3">
                        <Tag
                            color={activeMetric === 'spending' ? 'blue' : (activeMetric === 'documents' ? 'purple' : 'cyan')}
                            style={{ borderRadius: '4px', border: 'none', padding: '3px 10px', fontSize: '12px' }}
                        >
                            <span className="fw-medium">{selectedYear} Total: </span>
                            <span className="fw-bold">
                                {activeMetric === 'spending'
                                    ? `RM ${yearTotal.spendingMyr.toFixed(2)}`
                                    : (activeMetric === 'documents'
                                        ? `${yearTotal.documents} document${yearTotal.documents === 1 ? '' : 's'}`
                                        : `${yearTotal.tokens.toLocaleString()} tokens`)}
                            </span>
                        </Tag>
                    </div>

                    {/* Chart Container */}
                    {isMounted ? (
                        <div style={{ width: '100%', height: 320 }}>
                            <ResponsiveContainer width="100%" height="100%">
                                <AreaChart
                                    data={monthlyChartData}
                                    margin={{ top: 10, right: 15, left: -10, bottom: 0 }}
                                >
                                    <defs>
                                        <linearGradient id={metricConfig.fillGradientId} x1="0" y1="0" x2="0" y2="1">
                                            <stop offset="5%" stopColor={metricConfig.gradientStopColor} stopOpacity={0.25} />
                                            <stop offset="95%" stopColor={metricConfig.gradientStopColor} stopOpacity={0.0} />
                                        </linearGradient>
                                    </defs>
                                    <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f0f0f0" />
                                    <XAxis
                                        dataKey="month"
                                        stroke="#8c8c8c"
                                        fontSize={12}
                                        tickLine={false}
                                        axisLine={{ stroke: '#f0f0f0' }}
                                    />
                                    <YAxis
                                        stroke="#8c8c8c"
                                        fontSize={12}
                                        tickLine={false}
                                        axisLine={false}
                                        tickFormatter={metricConfig.yAxisFormatter}
                                    />
                                    <RechartsTooltip content={<CustomChartTooltip />} />
                                    <Area
                                        type="monotone"
                                        dataKey={metricConfig.dataKey}
                                        stroke={metricConfig.strokeColor}
                                        strokeWidth={2.5}
                                        fillOpacity={1}
                                        fill={`url(#${metricConfig.fillGradientId})`}
                                        dot={{ r: 3, fill: metricConfig.strokeColor, strokeWidth: 1 }}
                                        activeDot={{ r: 6, stroke: '#ffffff', strokeWidth: 2 }}
                                    />
                                </AreaChart>
                            </ResponsiveContainer>
                        </div>
                    ) : (
                        <div className="d-flex justify-content-center align-items-center" style={{ width: '100%', height: 320 }}>
                            <Spin />
                        </div>
                    )}
                </Card>

                {/* Bottom Row: Recent Generations Feed */}
                <Card
                    className="shadow-sm border-0"
                    style={{ borderRadius: '12px' }}
                    title={
                        <div className="d-flex justify-content-between align-items-center py-1">
                            <span className="fw-semibold" style={{ fontSize: '16px' }}>Recent Generations</span>
                            <Button
                                type="link"
                                onClick={() => router.push('/history')}
                                className="p-0 d-flex align-items-center gap-1"
                            >
                                View All History <ArrowRightOutlined style={{ fontSize: '12px' }} />
                            </Button>
                        </div>
                    }
                >
                    {recentTasks.length === 0 ? (
                        <Empty image={Empty.PRESENTED_IMAGE_SIMPLE} description="No recent generation activity" />
                    ) : (
                        <Table
                            columns={recentColumns}
                            dataSource={recentTasks}
                            rowKey="taskId"
                            pagination={false}
                            className="custom-table"
                            rowClassName={() => 'align-middle'}
                            scroll={{ x: 'max-content' }}
                        />
                    )}
                </Card>

            </div>
        </div>
    );
}

