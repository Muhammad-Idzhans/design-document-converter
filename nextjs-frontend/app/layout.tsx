
import type { Metadata } from "next";
import { Inter } from "next/font/google";
import { AntdRegistry } from '@ant-design/nextjs-registry';
import 'bootstrap/dist/css/bootstrap.min.css';
import "./globals.css";

import NextAuthProvider from "@/components/auth/NextAuthProvider";

const inter = Inter({
  subsets: ["latin"],
  display: 'swap',
});

export const metadata: Metadata = {
  title: "Enfrasys Document Generator",
  description: "Convert design slides (pptx) to Word Document (docx) & Final Configuration (HTML) using AI.",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en" suppressHydrationWarning>
      <body className={`${inter.className} d-flex flex-column vh-100 overflow-hidden m-0 p-0`} suppressHydrationWarning>
        <NextAuthProvider>
          <AntdRegistry>
            {children}
          </AntdRegistry>
        </NextAuthProvider>
      </body>
    </html>
  );
}