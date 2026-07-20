"use client";

import React, { useState, useEffect, useMemo } from 'react';
import { Table, Tag, Typography, Dropdown, MenuProps, message, Spin, Modal } from 'antd';
import { EyeOutlined, DownloadOutlined, EllipsisOutlined, DeleteOutlined, ExclamationCircleOutlined } from '@ant-design/icons';
import { useRouter } from 'next/navigation';
import dayjs from 'dayjs';

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
                    fetch('/api/history'),
                    fetch('/api/settings')
                ]);

                const historyData = await historyRes.json();
                const settingsData = await settingsRes.json();

                if (historyData.data) setTasks(historyData.data);
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
            content: 'This will permanently delete the record from the database and the files from cloud storage. This action cannot be undone.',
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
            title: 'Task ID',
            dataIndex: 'taskId',
            key: 'taskId',
            render: (text: string) => (
                <Text type="secondary" style={{ fontSize: '13px' }}>
                    {text.substring(0, 8)}...
                </Text>
            ),
            width: '10%',
        },
        {
            title: 'Document',
            key: 'document',
            render: (_: any, record: any) => {
                const formattedName = record.generated_filename || "Generated_Design_Document";

                return (
                    <div>
                        <div className="fw-medium text-dark">{formattedName}</div>
                        <Text type="secondary" style={{ fontSize: '12px' }}>{record.filename || "Generated_Design_Document"}</Text>
                    </div>
                );
            },
            width: '35%',
        },
        {
            title: 'Type',
            dataIndex: 'documentType',
            key: 'documentType',
            render: (text: string) => (
                <Text type="secondary">{text || 'Design Document'}</Text>
            ),
            width: '15%',
        },
        {
            title: 'Date',
            dataIndex: 'createdAt',
            key: 'date',
            render: (text: string) => {
                if (!text) return '-';
                return (
                    <div>
                        <div className="fw-medium text-dark">{dayjs(text).format('D MMM YYYY')}</div>
                        <Text type="secondary" style={{ fontSize: '12px' }}>{dayjs(text).format('h:mm A')}</Text>
                    </div>
                );
            },
            width: '15%',
        },
        {
            title: 'Status',
            dataIndex: 'status',
            key: 'status',
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
                    text = 'In review'; // matching your sample image loosely
                }
                return <Tag color={color} style={{ borderRadius: '4px', border: 'none', padding: '2px 8px' }}>{text}</Tag>;
            },
            width: '10%',
        },
        {
            title: 'Cost',
            key: 'cost',
            render: (_: any, record: any) => {
                if (!settings || !record.cost_metrics) return '-';

                const metrics = record.cost_metrics;
                const visionInputRate = (settings.visionInput || 0) / 1_000_000;
                const visionOutputRate = (settings.visionOutput || 0) / 1_000_000;
                const llmInputRate = (settings.llmInput || 0) / 1_000_000;
                const llmCompletionRate = (settings.llmCompletion || 0) / 1_000_000;
                const cuPageRate = (settings.contentUnderstanding || 0) / 1_000;
                const exchangeRate = settings.exchangeRate || 4.20;

                let costUsd = 0;
                costUsd += (metrics.vision_tokens_prompt || 0) * visionInputRate;
                costUsd += (metrics.vision_tokens_completion || 0) * visionOutputRate;
                costUsd += (metrics.llm_tokens_prompt || 0) * llmInputRate;
                costUsd += (metrics.llm_tokens_completion || 0) * llmCompletionRate;
                costUsd += (metrics.content_understanding_pages || 0) * cuPageRate;

                const costMyr = costUsd * exchangeRate;

                return (
                    <Text className="fw-medium">
                        RM{costMyr.toFixed(2)} <Text type="secondary" style={{ fontSize: '13px' }}>(${costUsd.toFixed(2)})</Text>
                    </Text>
                );
            },
            width: '15%',
            align: 'right' as const,
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
                        onClick: () => router.push(`/tasks/${record.taskId}/review?from=history`)
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
            width: '5%',
            align: 'center' as const,
        }
    ];

    return (
        <div className="container-fluid bg-light d-flex flex-column flex-grow-1 py-4 pb-5">
            <div className="container animate__animated animate__fadeIn w-100" style={{ maxWidth: '1200px' }}>
                <div className="mb-4">
                    <Title level={2} className="mb-1" style={{ letterSpacing: "-0.5px" }}>Generation History</Title>
                    <Text type="secondary" style={{ fontSize: "1.05rem" }}>
                        View all past generated documents and their associated AI pricing.
                    </Text>
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
                        />
                    )}
                </div>
            </div>
        </div>
    );
}
