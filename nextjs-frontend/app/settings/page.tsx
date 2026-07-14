"use client";

import React, { useState, useEffect } from 'react';
import { Card, Input, Button, Divider, Form, message, Row, Col, Typography } from 'antd';
import {
    DollarOutlined,
    SyncOutlined,
    RobotOutlined,
    EyeOutlined,
    CloudServerOutlined,
    SaveOutlined,
    ExportOutlined
} from '@ant-design/icons';

const { Title, Text } = Typography;

const CurrencyInput = ({ value, onChange, prefix, addonBefore }: { value?: any, onChange?: (val: number) => void, prefix?: React.ReactNode, addonBefore?: React.ReactNode }) => {
    const displayValue = value !== undefined && value !== null ? Number(value).toFixed(2) : "";

    const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
        const digitsOnly = e.target.value.replace(/\D/g, '');
        if (digitsOnly === "") {
            onChange?.(0);
            return;
        }
        const numericValue = parseInt(digitsOnly, 10) / 100;
        onChange?.(numericValue);
    };

    return (
        <Input
            size="large"
            addonBefore={addonBefore}
            prefix={prefix}
            value={displayValue}
            onChange={handleChange}
            inputMode="numeric"
            style={{ borderRadius: "8px" }}
        />
    );
};

export default function SettingsPage() {
    const [loadingRates, setLoadingRates] = useState(false);
    const [saving, setSaving] = useState(false);
    const [isDirty, setIsDirty] = useState(false);
    const [lastUpdated, setLastUpdated] = useState<string>("");
    const [form] = Form.useForm();

    useEffect(() => {
        // Fetch settings on mount
        fetch('/api/settings')
            .then(res => res.json())
            .then(data => {
                if (!data.error) {
                    form.setFieldsValue(data);
                    if (data.lastUpdated) {
                        setLastUpdated(data.lastUpdated);
                    }
                }
            })
            .catch(err => console.error("Failed to load settings", err));
    }, [form]);

    const handleFetchRates = async () => {
        setLoadingRates(true);
        try {
            const response = await fetch('https://api.exchangerate-api.com/v4/latest/USD');
            if (!response.ok) throw new Error('Failed to fetch');
            const data = await response.json();
            
            if (data && data.rates && data.rates.MYR) {
                const liveRate = data.rates.MYR;
                form.setFieldsValue({ exchangeRate: liveRate });
                setIsDirty(true);
                message.success(`Live exchange rate fetched: 1 USD = RM ${liveRate.toFixed(2)}`);
            } else {
                throw new Error('MYR rate not found');
            }
        } catch (error) {
            console.error("Fetch rate error:", error);
            message.error("Failed to fetch live exchange rate.");
        } finally {
            setLoadingRates(false);
        }
    };

    const handleSave = async (values: any) => {
        setSaving(true);
        try {
            const res = await fetch('/api/settings', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(values)
            });
            const data = await res.json();
            if (data.success) {
                message.success("Settings saved successfully!");
                setIsDirty(false);
                if (data.settings.lastUpdated) {
                    setLastUpdated(data.settings.lastUpdated);
                }
            } else {
                message.error("Failed to save settings.");
            }
        } catch (error) {
            console.error("Save error:", error);
            message.error("An error occurred while saving.");
        } finally {
            setSaving(false);
        }
    };

    return (
        <div className="container-fluid bg-light d-flex flex-column flex-grow-1 py-4 pb-5">
            <div className="container animate__animated animate__fadeIn w-100" style={{ maxWidth: '1200px' }}>
                <div className="mb-4 d-flex justify-content-between align-items-end">
                    <div>
                        <Title level={2} className="mb-1" style={{ letterSpacing: "-0.5px" }}>Settings & Configurations</Title>
                        <Text type="secondary" style={{ fontSize: "1.05rem" }}>
                            Manage application parameters, AI model pricing, and conversion rates.
                        </Text>
                    </div>
                    {lastUpdated && (
                        <div className="text-muted small">
                            Last updated: <span className="fw-medium">{lastUpdated}</span>
                        </div>
                    )}
                </div>

                <Form
                    form={form}
                    layout="vertical"
                    onFinish={handleSave}
                    onValuesChange={() => setIsDirty(true)}
                    initialValues={{
                        exchangeRate: 4.45,
                        visionInput: 2.50,
                        visionOutput: 10.00,
                        llmInput: 2.00,
                        llmCompletion: 8.00,
                        contentUnderstanding: 5.00
                    }}
                >
                    {/* Currency Conversion Section */}
                    <Card
                        className="mb-4 shadow-sm border-0 settings-card"
                        title={
                            <span className="fw-semibold d-flex align-items-center" style={{ fontSize: "16px" }}>
                                <div className="bg-primary bg-opacity-10 text-primary p-2 rounded me-3 d-flex">
                                    <DollarOutlined />
                                </div>
                                Currency Conversion
                            </span>
                        }
                        styles={{ header: { borderBottom: '1px solid #f0f0f0', padding: "16px 24px" }, body: { padding: "24px" } }}
                    >
                        <Row align="bottom" gutter={16}>
                            <Col xs={24} sm={16}>
                                <Form.Item
                                    label={<span className="fw-medium">USD to MYR Exchange Rate</span>}
                                    name="exchangeRate"
                                    tooltip="This rate is used to calculate the estimated MYR cost for generated documents."
                                    className="mb-0"
                                >
                                    <CurrencyInput
                                        addonBefore={<span className="fw-medium px-1" style={{ color: "#555" }}>1 US Dollar =</span>}
                                        prefix={<span className="text-muted fw-medium me-1">RM</span>}
                                    />
                                </Form.Item>
                            </Col>
                            <Col xs={24} sm={8}>
                                <Button
                                    type="dashed"
                                    size="large"
                                    icon={<SyncOutlined spin={loadingRates} />}
                                    onClick={handleFetchRates}
                                    block
                                    className="mt-3 mt-sm-0"
                                    style={{ borderRadius: "8px", borderColor: "#1677ff", color: "#1677ff" }}
                                >
                                    Auto Fetch Live Rate
                                </Button>
                            </Col>
                        </Row>
                    </Card>

                    {/* AI Model Pricing Section */}
                    <Card
                        className="mb-4 shadow-sm border-0 settings-card"
                        title={
                            <span className="fw-semibold d-flex align-items-center" style={{ fontSize: "16px" }}>
                                <div className="bg-success bg-opacity-10 text-success p-2 rounded me-3 d-flex">
                                    <RobotOutlined />
                                </div>
                                AI Model Pricing (USD)
                            </span>
                        }
                        extra={
                            <a
                                href="https://azure.microsoft.com/en-us/pricing/details/azure-openai/"
                                target="_blank"
                                rel="noopener noreferrer"
                                className="text-decoration-none d-flex align-items-center small"
                            >
                                <span className="me-1">Official Pricing</span> <ExportOutlined />
                            </a>
                        }
                        styles={{ header: { borderBottom: '1px solid #f0f0f0', padding: "16px 24px" }, body: { padding: "24px" } }}
                    >
                        <div className="mb-2">
                            <Text strong className="d-flex align-items-center mb-3" style={{ fontSize: "15px", color: "#2c3e50" }}>
                                <EyeOutlined className="me-2 text-info" /> Vision Model (GPT-4o)
                            </Text>
                            <Row gutter={24}>
                                <Col xs={24} sm={12}>
                                    <Form.Item label={<span className="text-muted">Input Price (per 1M tokens)</span>} name="visionInput">
                                        <CurrencyInput prefix={<span className="text-muted">$</span>} />
                                    </Form.Item>
                                </Col>
                                <Col xs={24} sm={12}>
                                    <Form.Item label={<span className="text-muted">Output Price (per 1M tokens)</span>} name="visionOutput">
                                        <CurrencyInput prefix={<span className="text-muted">$</span>} />
                                    </Form.Item>
                                </Col>
                            </Row>
                        </div>

                        <Divider dashed style={{ margin: '8px 0 24px 0' }} />

                        <div>
                            <Text strong className="d-flex align-items-center mb-3" style={{ fontSize: "15px", color: "#2c3e50" }}>
                                <RobotOutlined className="me-2 text-warning" /> LLM Model (GPT-4.1)
                            </Text>
                            <Row gutter={24}>
                                <Col xs={24} sm={12}>
                                    <Form.Item label={<span className="text-muted">Input Price (per 1M tokens)</span>} name="llmInput">
                                        <CurrencyInput prefix={<span className="text-muted">$</span>} />
                                    </Form.Item>
                                </Col>
                                <Col xs={24} sm={12}>
                                    <Form.Item label={<span className="text-muted">Completion Price (per 1M tokens)</span>} name="llmCompletion">
                                        <CurrencyInput prefix={<span className="text-muted">$</span>} />
                                    </Form.Item>
                                </Col>
                            </Row>
                        </div>
                    </Card>

                    {/* Azure Service Pricing Section */}
                    <Card
                        className="mb-4 shadow-sm border-0 settings-card"
                        title={
                            <span className="fw-semibold d-flex align-items-center" style={{ fontSize: "16px" }}>
                                <div className="bg-info bg-opacity-10 text-info p-2 rounded me-3 d-flex">
                                    <CloudServerOutlined />
                                </div>
                                Azure Service Pricing (USD)
                            </span>
                        }
                        extra={
                            <a
                                href="https://azure.microsoft.com/en-us/pricing/details/content-understanding/?msockid=1b3ed87184dd64901c70ce4f854965a6"
                                target="_blank"
                                rel="noopener noreferrer"
                                className="text-decoration-none d-flex align-items-center small"
                            >
                                <span className="me-1">Official Pricing</span> <ExportOutlined />
                            </a>
                        }
                        styles={{ header: { borderBottom: '1px solid #f0f0f0', padding: "16px 24px" }, body: { padding: "24px" } }}
                    >
                        <Row gutter={16}>
                            <Col xs={24} md={12}>
                                <Form.Item
                                    label={<span className="fw-medium">Content Understanding (per 1k pages)</span>}
                                    name="contentUnderstanding"
                                    className="mb-0"
                                >
                                    <CurrencyInput prefix={<span className="text-muted fw-medium me-1">$</span>} />
                                </Form.Item>
                            </Col>
                            <Col xs={24} md={12} className="d-flex align-items-center">
                                <Text type="secondary" className="mt-3 mt-md-0 ms-0 ms-md-2" style={{ fontSize: '13px', lineHeight: "1.5" }}>
                                    Base extraction cost applied by the Azure AI Content Understanding service per 1,000 pages processed.
                                </Text>
                            </Col>
                        </Row>
                    </Card>

                    {/* Submit Actions */}
                    <div className="d-flex justify-content-end mb-5 mt-4">
                        <Button
                            type="primary"
                            danger
                            size="large"
                            className="me-3 border-0 shadow-sm"
                            style={{ borderRadius: "8px" }}
                            disabled={!isDirty}
                            onClick={() => {
                                form.resetFields();
                                setIsDirty(false);
                            }}
                        >
                            Discard Changes
                        </Button>
                        <Button
                            type="primary"
                            size="large"
                            htmlType="submit"
                            icon={<SaveOutlined />}
                            loading={saving}
                            disabled={!isDirty}
                            style={{ borderRadius: "8px", padding: "0 32px" }}
                        >
                            Save Settings
                        </Button>
                    </div>
                </Form>
            </div>
        </div>
    );
}
