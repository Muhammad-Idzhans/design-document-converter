"use client";

import React, { useState, useEffect, useMemo } from 'react';
import { Table, Tag, Typography, Dropdown, MenuProps, message, Spin, Modal, Tooltip } from 'antd';
import { EyeOutlined, DownloadOutlined, EllipsisOutlined, DeleteOutlined, ExclamationCircleOutlined } from '@ant-design/icons';
import { useRouter } from 'next/navigation';
import dayjs from 'dayjs';
import { resolveCost } from '@/lib/cost';

const { Title, Text } = Typography;

export default function HistoryPage() {
    const router = useRouter();
    const [tasks, setTasks] = useState<any[]>([]);
    const [settings, setSettings] = useState<any>(null);
    const [loading, setLoading] = useState(true);

    useEffect(() => {
        const fetchData = async () => {
            try {
                const [historyRes, settingsRes] = await Promise.all([
                    fetch('/api/history?active_only=true'),
                    fetch('/api/settings')
                ]);

                const historyData = await historyRes.json();
                const settingsData = await settingsRes.json();

                if (historyData.data) {
                    setTasks(historyData.data.filter((t: any) => !t.isDeleted));
                }
                if (settingsData && !settingsData.error) setSettings(settingsData);
            } catch (error) {
                console.error("Failed to fetch data:", error);
                message.error("Failed to load history.");
            } finally {
                setLoading(false);
            }
        };

        fetchData();
    }, []);

    const handleDelete = (taskId: string) => {
        Modal.confirm({
            title: 'Are you sure you want to delete this document?',
            icon: <ExclamationCircleOutlined />,
            content: 'This will permanently delete the document from cloud storage and remove it from your history list. Historical cost and generation metrics will be retained on the dashboard.',
            okText: 'Yes, delete it',
            okType: 'danger',
            cancelText: 'Cancel',
            onOk: async () => {
                try {
                    const response = await fetch(`/api/tasks/${taskId}`, { method: 'DELETE' });

                    if (!response.ok) throw new Error("Failed to delete task");

                    setTasks(prev => prev.filter(t => t.taskId !== taskId));
                    message.success('Document deleted successfully');
                } catch (error) {
                    console.error("Delete error:", error);
                    message.error('Failed to delete document');
                }
            }
        });
    };

    // Column definitions for the Ant Design Table
    const columns = [
        {
            title: <span style={{ whiteSpace: 'nowrap' }}>Task ID</span>,
            dataIndex: 'taskId',
            key: 'taskId',
            width: 120,
            render: (text: string) => {
                if (!text) return '-';
                const shortId = text.length > 8 ? `${text.substring(0, 8)}...` : text;
                return (
                    <Tooltip title={text} placement="topLeft">
                        <Text
                            type="secondary"
                            style={{
                                fontSize: '13px',
                                fontFamily: 'SFMono-Regular, Menlo, Monaco, Consolas, monospace',
                                whiteSpace: 'nowrap',
                                cursor: 'default',
                            }}
                        >
                            {shortId}
                        </Text>
                    </Tooltip>
                );
            },
        },
        {
            title: 'Document',
            key: 'document',
            render: (_: any, record: any) => {
                const formattedName = record.generated_filename || "Generated_Design_Document";
                const originalName = record.filename || "Generated_Design_Document";

                return (
                    <div style={{ minWidth: 0, maxWidth: '300px' }}>
                        <Tooltip title={formattedName} placement="topLeft">
                            <div
                                className="fw-medium text-dark"
                                style={{
                                    overflow: 'hidden',
                                    textOverflow: 'ellipsis',
                                    whiteSpace: 'nowrap',
                                    cursor: 'default',
                                }}
                            >
                                {formattedName}
                            </div>
                        </Tooltip>
                        <Tooltip title={originalName} placement="topLeft">
                            <Text
                                type="secondary"
                                style={{
                                    fontSize: '12px',
                                    display: 'block',
                                    overflow: 'hidden',
                                    textOverflow: 'ellipsis',
                                    whiteSpace: 'nowrap',
                                    cursor: 'default',
                                }}
                            >
                                {originalName}
                            </Text>
                        </Tooltip>
                    </div>
                );
            },
        },
        {
            title: <span style={{ whiteSpace: 'nowrap' }}>Type</span>,
            dataIndex: 'documentType',
            key: 'documentType',
            width: 170,
            render: (text: string) => {
                const isFinalConfig = text?.toLowerCase().includes('final');
                const label = isFinalConfig ? 'Final Configuration' : (text || 'Design Document');
                return (
                    <Tag
                        color={isFinalConfig ? 'purple' : 'blue'}
                        style={{
                            borderRadius: '4px',
                            border: 'none',
                            padding: '2px 8px',
                            fontSize: '12px',
                            fontWeight: 500,
                            whiteSpace: 'nowrap',
                        }}
                    >
                        {label}
                    </Tag>
                );
            },
        },
        {
            title: <span style={{ whiteSpace: 'nowrap' }}>Date</span>,
            dataIndex: 'createdAt',
            key: 'date',
            width: 130,
            render: (text: string) => {
                if (!text) return '-';
                return (
                    <div style={{ whiteSpace: 'nowrap' }}>
                        <div className="fw-medium text-dark">{dayjs(text).format('D MMM YYYY')}</div>
                        <Text type="secondary" style={{ fontSize: '12px' }}>{dayjs(text).format('h:mm A')}</Text>
                    </div>
                );
            },
        },
        {
            title: <span style={{ whiteSpace: 'nowrap' }}>Status</span>,
            dataIndex: 'status',
            key: 'status',
            width: 110,
            render: (status: string) => {
                let color = 'default';
                let text = status;
                if (status === 'completed') {
                    color = 'success';
                    text = 'Completed';
                } else if (status === 'failed') {
                    color = 'error';
                    text = 'Failed';
                } else if (status === 'processing_upload') {
                    color = 'processing';
                    text = 'In review';
                }
                return (
                    <Tag
                        color={color}
                        style={{
                            borderRadius: '4px',
                            border: 'none',
                            padding: '2px 8px',
                            whiteSpace: 'nowrap',
                            fontWeight: 500,
                        }}
                    >
                        {text}
                    </Tag>
                );
            },
        },
        {
            title: <span style={{ whiteSpace: 'nowrap' }}>Cost</span>,
            key: 'cost',
            width: 140,
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
                        <div style={{ whiteSpace: 'nowrap', cursor: 'help' }}>
                            <Text className="fw-medium">
                                RM{cost.costMyr.toFixed(2)}{' '}
                                <Text type="secondary" style={{ fontSize: '13px' }}>
                                    (${cost.costUsd.toFixed(2)})
                                </Text>
                            </Text>
                        </div>
                    </Tooltip>
                );
            },
        },
        {
            title: '',
            key: 'action',
            render: (_: any, record: any) => {
                const items: MenuProps['items'] = [
                    {
                        key: 'view',
                        label: 'View Document',
                        icon: <EyeOutlined />,
                        onClick: () => router.push(`/history/${record.taskId}`)
                    },
                    {
                        key: 'download',
                        label: 'Download Document',
                        icon: <DownloadOutlined />,
                        onClick: async () => {
                            try {
                                const API_BASE_URL = process.env.NEXT_PUBLIC_API_URL || "http://localhost:8000";
                                const docName = record.generated_filename || "Generated_Design_Document";
                                const response = await fetch(`${API_BASE_URL}/api/download/${record.taskId}?filename=${encodeURIComponent(docName)}`);

                                if (!response.ok) throw new Error("Download failed");

                                const blob = await response.blob();
                                const url = window.URL.createObjectURL(blob);
                                const a = document.createElement("a");
                                a.href = url;
                                a.download = `${docName}.docx`;
                                document.body.appendChild(a);
                                a.click();
                                a.remove();
                                window.URL.revokeObjectURL(url);
                            } catch (error) {
                                console.error("Download Error:", error);
                                message.error("Failed to download document.");
                            }
                        }
                    },
                    {
                        type: 'divider'
                    },
                    {
                        key: 'delete',
                        label: 'Delete Document',
                        icon: <DeleteOutlined />,
                        danger: true,
                        onClick: () => handleDelete(record.taskId || record.id)
                    }
                ];

                return (
                    <Dropdown menu={{ items }} trigger={['click']} placement="bottomRight">
                        <div style={{ cursor: 'pointer', padding: '4px 8px' }}>
                            <EllipsisOutlined style={{ fontSize: '18px', color: '#555' }} />
                        </div>
                    </Dropdown>
                );
            },
            width: 50,
            align: 'center' as const,
        }
    ];

    return (
        <div className="container-fluid bg-light d-flex flex-column flex-grow-1 py-4 pb-5">
            <div className="container animate__animated animate__fadeIn w-100" style={{ maxWidth: '1200px' }}>
                <div className="mt-3 mb-4">
                    <h2 className="fw-bold mb-2">Generation History</h2>
                    <p className="text-muted mb-0">
                        View all past generated documents and their associated AI pricing.
                    </p>
                </div>

                <div className="bg-white rounded shadow-sm border p-4">
                    {loading ? (
                        <div className="d-flex justify-content-center py-5">
                            <Spin size="large" />
                        </div>
                    ) : (
                        <Table
                            columns={columns}
                            dataSource={tasks}
                            rowKey="taskId"
                            pagination={{ pageSize: 10, showSizeChanger: false }}
                            className="custom-table"
                            rowClassName={() => 'align-middle'}
                            scroll={{ x: 'max-content' }}
                        />
                    )}
                </div>
            </div>
        </div>
    );
}
