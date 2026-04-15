// // import type { Metadata } from "next";
// // import { Inter } from "next/font/google";
// // import { AntdRegistry } from '@ant-design/nextjs-registry';
// // import 'bootstrap/dist/css/bootstrap.min.css';
// // import "./globals.css";
// // import Header from "../components/layout/Header";

// // const inter = Inter({
// //   subsets: ["latin"],
// //   display: 'swap',
// // });

// // export const metadata: Metadata = {
// //   title: "Design Document Converter",
// //   description: "Convert design slides (pptx) to Word Document (docx) using AI.",
// // };

// // export default function RootLayout({
// //   children,
// // }: Readonly<{
// //   children: React.ReactNode;
// // }>) {
// //   return (
// //     <html lang="en">
// //       <body className={inter.className}>
// //         <AntdRegistry>
// //           <Header />
// //           {children}
// //         </AntdRegistry>
// //       </body>
// //     </html>
// //   );
// // }


// import type { Metadata } from "next";
// import { Inter } from "next/font/google";
// import { AntdRegistry } from '@ant-design/nextjs-registry';
// import 'bootstrap/dist/css/bootstrap.min.css';
// import "./globals.css";
// import Header from "../components/layout/Header";

// const inter = Inter({
//   subsets: ["latin"],
//   display: 'swap',
// });

// export const metadata: Metadata = {
//   title: "Design Document Converter",
//   description: "Convert design slides (pptx) to Word Document (docx) using AI.",
// };

// export default function RootLayout({
//   children,
// }: Readonly<{
//   children: React.ReactNode;
// }>) {
//   return (
//     <html lang="en">
//       {/* 1. STRICT HEIGHT: vh-100 and overflow-hidden locks the browser window from scrolling entirely */}
//       <body className={`${inter.className} d-flex flex-column vh-100 overflow-hidden m-0 p-0`}>
//         <AntdRegistry>

//           <Header />

//           {/* 2. SCROLLABLE MAIN: flex-grow-1 fills the rest of the screen, overflow-auto lets ONLY the content scroll */}
//           <main className="flex-grow-1 overflow-auto d-flex flex-column position-relative">
//             {children}
//           </main>

//         </AntdRegistry>
//       </body>
//     </html>
//   );
// }

import type { Metadata } from "next";
import { Inter } from "next/font/google";
import { AntdRegistry } from '@ant-design/nextjs-registry';
import 'bootstrap/dist/css/bootstrap.min.css';
import "./globals.css";
import Header from "../components/layout/Header";

const inter = Inter({
  subsets: ["latin"],
  display: 'swap',
});

export const metadata: Metadata = {
  title: "Design Document Converter",
  description: "Convert design slides (pptx) to Word Document (docx) using AI.",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en">
      {/* 1. Make the body a flex-column that takes up at least 100vh */}
      <body className={`${inter.className} d-flex flex-column min-vh-100 m-0`}>
        <AntdRegistry>

          <Header />

          {/* 2. Wrap children in a <main> tag that flex-grows to fill the remaining space */}
          <main className="flex-grow-1 d-flex flex-column">
            {children}
          </main>

        </AntdRegistry>
      </body>
    </html>
  );
}