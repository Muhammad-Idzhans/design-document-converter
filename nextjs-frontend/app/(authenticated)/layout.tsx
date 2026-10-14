"use client";

import React from 'react';
import { Layout } from 'antd';
import Header from "../../components/layout/Header";
import Sidebar from "../../components/layout/Sidebar";

const { Content } = Layout;

export default function AuthenticatedLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <Layout style={{ height: '100vh', overflow: 'hidden' }}>
      {/* Header is fixed at 56px height (defined in Header.tsx) */}
      <Header />
      
      {/* Inner Layout takes the remaining height */}
      <Layout style={{ height: 'calc(100vh - 56px)' }}>
        <Sidebar />
        
        {/* Content area is the ONLY scrollable part of the page */}
        <Content
          style={{
            height: '100%',
            overflowY: 'auto',
            backgroundColor: '#f5f5f5',
            display: 'flex',
            flexDirection: 'column',
          }}
        >
          {children}
        </Content>
      </Layout>
    </Layout>
  );
}
