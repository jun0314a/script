import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "보험 상담 스크립트 생성기",
  description: "나이, 성별, 기존보험 유무에 따른 맞춤 상담 스크립트",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="ko" className="h-full">
      <body className="min-h-full bg-gray-50 text-gray-900 antialiased">
        {children}
      </body>
    </html>
  );
}
