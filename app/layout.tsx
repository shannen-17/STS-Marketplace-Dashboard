import "./globals.css";
import type { Metadata } from "next";

export const metadata: Metadata = {
  title: "STS Marketplace Performance Dashboard",
  description: "Executive dashboard for STS marketplace reporting"
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}